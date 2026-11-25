Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiAOP2
    Public Lokal As Boolean = True
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjAOP As New ClsTransaksiAOP
    Public baru As Boolean
    Public vQuickCreateStockIn As Boolean = False
    Public Objsupplier As New clsMasterSupplier
    Public ObjCurrency As New Core_Module.clsMasterCurrency
    Public LanjutCreateAOP As Boolean = True
    Public LangsungClose As Boolean = False

    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    'Dim ObjSupplierProductCode As New clsTransaksiSettingSupplierProductCodeNo
    Dim objGRN As New ClsTransaksiGRN
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjAOPItem As New ClsTransaksiAOPItem
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjPR As New ClsTransaksiPurchaseReturn
    Dim ObjPRItem As New ClsTransaksiPurchaseReturnItem
    Dim ObjCL As New ClsTransaksiMonthlyClosing
    Dim ObjAkun As New ClsMasterAkun_Purchase
    Dim ObjLicense As New ClsLicense
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjStockIn As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjTaxRate As New Sales_Module.ClsTransaksiSalesInvoiceTaxRate
    Dim DefaultPrecision As Integer = 0
    Dim buffer As String = ""
    'Dim SedangRevisi As Boolean = False
    Public TBL As New DseTransaksiGRN.DataTable8DataTable
    Public rowGRN1 As DseTransaksiGRN.DataTable8Row
    Dim TBL2 As New DseTransaksiAOP.DataTable9DataTable
    Dim TBLPR As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_view1DataTable
    Dim DTAPO As New DseTransaksiGRNTableAdapters.tbltrans_po_item_view5TableAdapter
    Dim TBLPO As New DseTransaksiGRN.tbltrans_po_item_view5DataTable

    Dim TBLGRN As New DseTransaksiGRN.DataTable3DataTable
    Dim DTAGRN As New DseTransaksiGRNTableAdapters.DataTable3TableAdapter
    Dim rowGRN As DseTransaksiGRN.DataTable3Row

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Dim DseAccount As New AccModule.DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccount As New AccModule.DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccount As AccModule.DseMultiAccount.tblisi_cbo_ar_akunRow

    Dim DseCurrTax As New AccModule.DseMultiAccountTableAdapters.tblCurrencyTaxTableAdapter
    Dim tblCurrTax As New AccModule.DseMultiAccount.tblCurrencyTaxDataTable
    Dim rowCurrTax As AccModule.DseMultiAccount.tblCurrencyTaxRow

    '2025 Core Tax Tahap 2.1
    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo

    Public Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Liabilities
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Payable
        Dim objMulti As New AccModule.ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboAccount.DataSource = Nothing
            Me.cboAccount.Items.Clear()
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryAccount.Items.Clear()
            Me.tblAccount.Rows.Clear()

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccount, Me.txtPrimaryCurrency.Text)

            Me.cboAccount.DataSource = Me.tblAccount.DefaultView
            Me.cboAccount.DisplayMember = "Account Code"
            Me.cboPrimaryAccount.DataSource = Me.tblAccount.DefaultView
            Me.cboPrimaryAccount.DisplayMember = "Primary"
            Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiCBOCurrencyTax()
        '- isicbo customer
        Dim objMulti As New AccModule.ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboCurrTax.DataSource = Nothing
            Me.cboCurrTax.Items.Clear()
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboPrimaryCurrTax.Items.Clear()
            objMulti.IsiTBLCurrTax(Me.txtPrimaryCurrency.Text, Me.tblCurrTax)

            'Dim I As New Form5
            'I.DataGridView1.DataSource = Nothing
            'I.DataGridView1.DataSource = Me.tblAR
            'I.ShowDialog()

            Me.cboCurrTax.DataSource = Me.tblCurrTax.DefaultView
            Me.cboCurrTax.DisplayMember = "Currency Code"
            Me.cboPrimaryCurrTax.DataSource = Me.tblCurrTax.DefaultView
            Me.cboPrimaryCurrTax.DisplayMember = "Primary"
            Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged

        Try
            If Me.TxtRemark.Enabled = True AndAlso _
            IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso _
            (Me.txtPrimaryCurrency.Text) > 0 Then
                Me.IsiCBOCurrencyTax()
                Me.IsiCBOArAkunByTipePerkiraan1()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.TxtDutyStamp.Text) Then
                    Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                If Me.TxtRemark.Enabled = True Then
                    Me.Hitung_Total_Multi()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub cboPrimaryCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.TxtDutyStamp.Text) Then
                    Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                If Me.TxtRemark.Enabled = True Then
                    Me.Hitung_Total_Multi()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCurrTax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrTax.TextChanged
        Dim PrimaryCurrTax As Double = -1

        Try
            If IsNumeric(Me.txtPrimaryCurrTax.Text) Then
                PrimaryCurrTax = CDbl(Me.txtPrimaryCurrTax.Text)
            End If

            If Me.tblCurrTax.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblCurrTax.Rows.Count - 1
                    Me.rowCurrTax = Me.tblCurrTax.Rows(IntA)
                    With Me.rowCurrTax
                        If .Primary = PrimaryCurrTax Then
                            Me.cboPrimaryCurrTax.SelectedIndex = IntA
                            Me.cboCurrTax.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryAccount.SelectedIndexChanged
        Try
            ' Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryAccount.TextChanged
        Dim PrimaryARAkun As Double = -1

        Try
            If IsNumeric(Me.txtPrimaryAccount.Text) Then
                PrimaryARAkun = CDbl(Me.txtPrimaryAccount.Text)
            End If

            If Me.tblAccount.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblAccount.Rows.Count - 1
                    Me.rowAccount = Me.tblAccount.Rows(IntA)
                    With Me.rowAccount
                        If .Primary = PrimaryARAkun Then
                            Me.cboPrimaryAccount.SelectedIndex = IntA
                            Me.cboAccount.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub NewData_Procedure()
        Dim UseTaxSetting, TaxPercentSetting As String

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            'Me.SedangRevisi = False
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TabControl1.SelectedIndex = 0
            Me.dtpDate.Focus()
            '  Me.txtPrimaryAccount.Visible = True
            Me.ObjAOP.ResetObject(Me.dgvItem, Me.DGVdp, Me.dgvExchageRates, Me.DGVitemPR, Me.dgvCost)
            'pertama kali dianggap supplier lokal dulu
            Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, -1)
            'Me.txtNoFakturPajak.Text = Me.ObjAOP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
            Me.ObjAOP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP

            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.AOP, ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
            Else
                If UseTaxSetting = "0" Then
                    Me.chkTax.Checked = False
                Else
                    Me.chkTax.Checked = True
                    TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.AOP, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(TaxPercentSetting) = True Then
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping(TaxPercentSetting, 2)
                    End If
                End If
            End If

            Me.txtExhangeRateToIDR.ReadOnly = True
            Me.SetDefaultPrecision()
            Me.SetCalculateBy()
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiCBOCurrencyTax()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.NewData_Procedure()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)

            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.DGVdp.AllowUserToDeleteRows = True
            Me.DGVitemPR.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvExchageRates.ReadOnly = False
            'set dgv cost, boleh edit, bisa tambahkan baris baru
            Me.dgvCost.ReadOnly = False
            Me.dgvCost.AllowUserToAddRows = True
            Me.dgvCost.AllowUserToDeleteRows = True

            Me.llbAddress.Enabled = True
            Me.llbAddDP.Enabled = True
            Me.llbAddSR.Enabled = True

            If Me.dgvItem.RowCount > 0 Then
                Me.txtSupplierCode.ReadOnly = True
                Me.txtSupplierCode.BackColor = Me.TxtAOP.BackColor
                Me.btnFindSupplier.Enabled = False
            Else
                Me.txtSupplierCode.ReadOnly = False
                Me.txtSupplierCode.BackColor = Me.TxtRemark.BackColor
                Me.btnFindSupplier.Enabled = True
            End If

            Me.cboAccount.Enabled = True
            Me.cboCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
            End If

            Me.SetTaxRate(False)

            ''@rosadinnn
            ''@MUaT
            Me.FormatDGV()
            Me.dgvItem.MultiSelect = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.ObjAOP
                Me.cboCalculateBy.SelectedIndex = .CalculatesBy
                Me.TxtPrimaryTransaksiAOP.Text = .Primary
                Me.TxtAOP.Text = .NoTransaksi
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision(False)
                Me.txtPrimarySupplier.Text = .PrimarySupplier
                Me.txtSupplierInvoiceNo.Text = .SupplierInvoiceNo
                Me.DtpInvDate.Value = .SupplierInvoiceDate
                Me.txtSupplierCode.Text = Me.Objsupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.txtSupplierAddIdentifier.Text = .SuppIdentifier
                Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.cboSupplierAddress.Text = Me.Objsupplier.GetKodeAddress(Me.txtSupplierAddIdentifier.Text)
                Me.IsiCboDeliveryOrder2(Me.txtPrimarySupplier.Text)
                Me.TxtPaymentsTerms.Text = Me.ObjCD.DigitGrouping(.PaymentTerms)

                'Exit Sub
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.DownPayment, Me.DefaultPrecision))
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubtotalBeforeDP, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxAmount, _
                Me.ObjCurrency.GetDecimal_DefaultCurrency))
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxExchangeRate, Me.DefaultPrecision))

            
                Me.txtExhangeRateToIDR.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExchangRateToIDR, Me.DefaultPrecision))
                Me.txtExchangeRateToUSD.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExhangeRateToUSD, Me.DefaultPrecision))
                If .TaxExclude = ClsTransaksiAOP.EnumTaxExclude.Exclude Then
                    Me.chkTax.Checked = True
                Else
                    Me.chkTax.Checked = False
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(.Tax, 2)
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(.TaxExchangeRate, 2)
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                Me.IsiDGVReturnItem(False)
                Me.TxtNPWPNo.Text = .NoNPWPSupplier
                Me.TxtTaxInvoiceNumber.Text = .NoFakturPajak
                Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = .EFak_KodeJenisTransaksiPajak
                If .VATDipungut = True Then
                    Me.cboPPNDipungut.SelectedIndex = 0
                Else
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If


                Me.IsiCBOArAkunByTipePerkiraan1()
                If .PrimaryAccount > -1 Then
                    Me.txtPrimaryAccount.Text = .PrimaryAccount
                    Me.cboPrimaryAccount.Text = .PrimaryAccount
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                End If

                Me.IsiCBOCurrencyTax()
                If .PrimaryCurrTax > -1 Then
                    Me.txtPrimaryCurrTax.Text = .PrimaryCurrTax
                    Me.cboPrimaryCurrTax.Text = .PrimaryCurrTax
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryCurrTax.DataSource = Nothing
                    Me.txtPrimaryCurrTax.Text = -1
                End If

                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
                Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(.DutyStamp, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))

                '2025 Core Tax Tahap 2.2
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(.DPPNilaiLain, Me.DefaultPrecision)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()
            'tampilan exchange rate
            Me.SetDGVExchangeRate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjAOP
                .Primary = CDec(Me.TxtPrimaryTransaksiAOP.Text)
                .NoTransaksi = Me.TxtAOP.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimarySupplier = CDec(Me.txtPrimarySupplier.Text)
                .SupplierInvoiceNo = Me.txtSupplierInvoiceNo.Text
                .SupplierInvoiceDate = Me.DtpInvDate.Value
                If Me.Lokal = True Then
                    .Lokal = ClsTransaksiAOP.EnumLokal.Lokal
                Else
                    .Lokal = ClsTransaksiAOP.EnumLokal.Import
                End If
                .SupplierInvoiceNo = Me.txtSupplierInvoiceNo.Text
                .PrimaryCurrency = CDec(Me.txtPrimaryCurrency.Text)
                .CalculatesBy = Me.cboCalculateBy.SelectedIndex
                .DownPayment = Me.txtDownPayment.Text
                .DutyStamp = Me.TxtDutyStamp.Text
                .SubTotal = Me.txtSubTotalBeforeDP.Text
                .Total = Me.TxtTotal.Text
                .PaymentTerms = Me.TxtPaymentsTerms.Text

                If Me.chkTax.Checked = True Then
                    .TaxExclude = ClsTransaksiAOP.EnumTaxExclude.Exclude
                Else
                    .TaxExclude = ClsTransaksiAOP.EnumTaxExclude.Include
                End If
                .SubTotal = Me.txtSubTotalAfterDP.Text
                .Tax = Me.txtTax.Text
                .Remarks = Me.TxtRemark.Text
                .ExchangRateToIDR = Me.txtExhangeRateToIDR.Text
                .ExhangeRateToUSD = Me.txtExchangeRateToUSD.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .SubtotalBeforeDP = Me.txtSubTotalBeforeDP.Text
                .DownPayment = Me.txtDownPayment.Text
                .TaxExchangeRate = Me.txtTaxExchangeRate.Text
                .SuppIdentifier = Me.txtSupplierAddIdentifier.Text
                '.Lokal = Me.CboLokal.SelectedIndex
                '.NoNPWP = Me.txtNoNPWP.Text
                '.NoSeriFakturPajak = Me.txtNoSeriFakturPajak.Text
                ' .Revisi = Me.txtRevisiNo.Text
                '.JenisFakturPajak = Me.CboFakturPajak.SelectedIndex
                .NoNPWPSupplier = Me.TxtNPWPNo.Text
                .NoFakturPajak = Me.TxtTaxInvoiceNumber.Text
                '2025 Core Tax Tahap 2.2
                '.EFak_KodeJenisTransaksiPajak = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
                If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    .VATDipungut = True
                Else
                    .VATDipungut = False
                End If
                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
                .PrimaryAccountDP = -1

                '2025 Core Tax Tahap 2.2
                .DPPNilaiLain = Me.txtDPPNilaiLain.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim i As New FrmCariMasterCurrency

        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetTaxRate()
                    Me.SetDefaultPrecision()
                    'tampilan exchange rate
                    Me.SetDGVExchangeRate(True)
                    FilterDP(Me.txtCurrencyCode.Text)
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetTaxRate()
                        Me.SetDefaultPrecision()
                        'tampilan exchange rate
                        FilterDP(Me.txtCurrencyCode.Text)
                        Me.SetDGVExchangeRate(True)
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.SetTaxRate()
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
            End If

            ' If Me.txtPrimaryCurrTax.Text = -1 Then
            'If Me.txtPrimaryCurrency.Text <= 1 Then
            '    Me.cboCurrTax.Enabled = False
            'Else
            '    Me.cboCurrTax.Enabled = True
            'End If

            Me.IsiCBOCurrencyTax()
            Me.IsiCBOArAkunByTipePerkiraan1()
            '  End If
            Me.SetCalculateBy()

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
                Me.cboCalculateBy.Focus()
                Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                Me.SetTaxRate()
                Me.SetDefaultPrecision()
                'tampilan exchange rate
                Me.SetDGVExchangeRate(True)
                FilterDP(Me.txtCurrencyCode.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.txtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetTaxRate()
                    Me.SetDefaultPrecision()
                    'tampilan exchange rate
                    Me.SetDGVExchangeRate(True)
                    FilterDP(Me.txtCurrencyCode.Text)
                    Me.cboCalculateBy.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetTaxRate()
                    Me.SetDefaultPrecision()
                End If
            End If

            ' If Me.txtPrimaryCurrTax.Text = -1 Then
            'If Me.txtPrimaryCurrency.Text <= 1 Then
            '    Me.cboCurrTax.Enabled = False
            'Else
            '    Me.cboCurrTax.Enabled = True
            'End If

            Me.IsiCBOCurrencyTax()
            Me.IsiCBOArAkunByTipePerkiraan1()
            ' End If

            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub


    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetTaxRate()
                    Me.SetDefaultPrecision()
                    Me.cboCalculateBy.Focus()
                    'tampilan exchange rate
                    Me.SetDGVExchangeRate(True)
                    FilterDP(Me.txtCurrencyCode.Text)
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetTaxRate()
                        Me.SetDefaultPrecision()
                        Me.cboCalculateBy.Focus()
                        'tampilan exchange rate
                        Me.SetDGVExchangeRate(True)
                        FilterDP(Me.txtCurrencyCode.Text)
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.SetTaxRate()
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If
            End If

            ' If Me.txtPrimaryCurrTax.Text = -1 Then
            If Me.txtPrimaryCurrency.Text <= 1 Then
                Me.cboCurrTax.Enabled = False
            Else
                Me.cboCurrTax.Enabled = True
            End If

            Me.IsiCBOCurrencyTax()
            Me.IsiCBOArAkunByTipePerkiraan1()
            ' End If

            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub Hitung_Total_Multi()
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0
        Dim exchangRate As Decimal = 0
        Dim IDR_DefPrec As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Total_Multi_MiddlePoint()
                Exit Sub
            End If
            IDR_DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.dgvExchageRates.Rows.Count > 0 Then
                'hitung sub total
                For Each row In Me.dgvItem.Rows
                    'yang diperhitungkan hanya yang tidak pake secondary unit
                    If Me.cboCalculateBy.SelectedIndex = 0 Then
                        'berdasarkan sales invoice
                        exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
                    ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
                        'berdasarkan delivery order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
                    ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
                        'berdasarkan saler order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
                    End If

                    'hitung purchase return jika ada
                    Dim DefaultPrecisionRow As Decimal = 1
                    Dim SubTotal1 As Decimal = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    Dim SecQty As Decimal
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                        SecQty = row.Cells(Me.ColSecQty.Index).Value
                    Else
                        SecQty = -1
                    End If
                    Dim JmlReturn As Decimal = 0

                    If SecQty <= 0 Then
                        If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                            JmlReturn = Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            row.Cells(Me.ColPrimaryGRN.Index).Value, SecQty, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)

                            SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                             row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
                        Else
                            JmlReturn = Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            row.Cells(Me.ColPrimaryGRN.Index).Value, SecQty, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)

                            SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                             row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
                        End If
                    Else
                        SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                        row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, DefaultPrecisionRow))
                    End If

                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value)
                Next
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.DGVdp.Rows
                    Me.txtDownPayment.Text = CDec(Me.txtDownPayment.Text) + CDec(row.Cells(ColDP_DPAMount.Index).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtDownPayment.Text, DefaultPrecision))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                'Me.txtSubTotalAfterDP.Text = Me.txtSubTotalBeforeDP.Text

                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.txtPrimaryCurrency.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                            CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrency.Text))
                        Else
                            If Me.txtPrimaryCurrTax.Text = "1" Then
                                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                                      CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                            Else

                                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                                       CDec(Me.txtSubTotalAfterDP.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                            End If
                        End If
                    Else
                        If Me.txtPrimaryCurrTax.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                        End If
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                    End If
                End If


                'hitung total = sub total + taxj 
                If Me.txtPrimaryCurrency.Text = "1" Then
                    'kalo IDR, duty stamp diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text) + _
                    CDec(Me.txtTaxAmount.Text) + CDec(Me.TxtDutyStamp.Text), DefaultPrecision))
                Else
                    If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                        'kalo Currencynya sama diperhitungkan duty stamp dan tax
                        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text) + _
                        CDec(Me.txtTaxAmount.Text) + CDec(Me.TxtDutyStamp.Text), DefaultPrecision))
                    Else
                        'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text), _
                        DefaultPrecision))
                    End If
                End If
            End If
            'hitung ulang unit price
            Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub Hitung_Total_Multi_MiddlePoint()
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0
        Dim exchangRate As Decimal = 0
        Dim IDR_DefPrec As Integer = 0

        Try
            IDR_DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.dgvExchageRates.Rows.Count > 0 Then
                'hitung sub total
                For Each row In Me.dgvItem.Rows
                    'yang diperhitungkan hanya yang tidak pake secondary unit
                    If Me.cboCalculateBy.SelectedIndex = 0 Then
                        'berdasarkan sales invoice
                        exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
                    ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
                        'berdasarkan delivery order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
                    ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
                        'berdasarkan saler order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
                    End If

                    'hitung purchase return jika ada
                    Dim DefaultPrecisionRow As Decimal = 1
                    Dim SubTotal1 As Decimal = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    Dim SecQty As Decimal
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                        SecQty = row.Cells(Me.ColSecQty.Index).Value
                    Else
                        SecQty = -1
                    End If
                    Dim JmlReturn As Decimal = 0

                    If SecQty <= 0 Then
                        If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                            JmlReturn = Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            row.Cells(Me.ColPrimaryGRN.Index).Value, SecQty, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)

                            SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                             row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
                        Else
                            JmlReturn = Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            row.Cells(Me.ColPrimaryGRN.Index).Value, SecQty, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)

                            SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                             row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
                        End If
                    Else
                        SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                        row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, DefaultPrecisionRow))
                    End If

                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value)
                Next
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.DGVdp.Rows
                    Me.txtDownPayment.Text = CDec(Me.txtDownPayment.Text) + CDec(row.Cells(ColDP_DPAMount.Index).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtDownPayment.Text, DefaultPrecision))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                'Me.txtSubTotalAfterDP.Text = Me.txtSubTotalBeforeDP.Text

                '2025 Core Tax Tahap 2.2
                Me.UbahDPPNilaiLain()

                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.txtPrimaryCurrency.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                            CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrency.Text))
                        Else
                            If Me.txtPrimaryCurrTax.Text = "1" Then
                                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                            Else

                                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                CDec(Me.txtSubTotalAfterDP.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                            End If
                        End If
                    Else
                        If Me.txtPrimaryCurrTax.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                        End If
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                    End If
                End If


                'hitung total = sub total + taxj 
                If Me.txtPrimaryCurrency.Text = "1" Then
                    'kalo IDR, duty stamp diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text) + _
                    CDec(Me.txtTaxAmount.Text) + CDec(Me.TxtDutyStamp.Text), DefaultPrecision))
                Else
                    If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                        'kalo Currencynya sama diperhitungkan duty stamp dan tax
                        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text) + _
                        CDec(Me.txtTaxAmount.Text) + CDec(Me.TxtDutyStamp.Text), DefaultPrecision))
                    Else
                        'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text), _
                        DefaultPrecision))
                    End If
                End If
            End If
            'hitung ulang unit price
            Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage4.Controls, True, True, True)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, False)


            Me.TxtPrimaryTransaksiAOP.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtPrimaryAccount.Text = "-1"
            Me.txtPrimarySupplier.Text = "-1"
            Me.txtSubTotalBeforeDP.Text = "0"
            Me.txtSubTotalAfterDP.Text = "0"
            Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", 2)
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.txtDownPayment.Text = "0"
            Me.TxtDutyStamp.Text = "0"
            Me.TxtPaymentsTerms.Text = "0"
            Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = "-1"
            Me.txtExchangeRateToUSD.Text = "1"
            Me.txtExhangeRateToIDR.Text = "1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.DGVdp.Rows.Clear()
            Me.dgvExchageRates.Rows.Clear()
            Me.DGVitemPR.Rows.Clear()
            Me.dgvPR.Rows.Clear()
            Me.dgvCost.Rows.Clear()

            Me.chkTax.Checked = True
            Me.cboSupplierAddress.DataSource = Nothing
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.cboCurrTax.DataSource = Nothing

            Me.cboSupplierAddress.Text = ""
            Me.cboPrimaryAccount.Text = ""
            Me.cboAccount.Text = ""
            Me.cboCurrTax.Text = ""

            Me.cboSupplierAddress.Items.Clear()
            Me.cboPrimaryAccount.Items.Clear()
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.cboAccount.Items.Clear()
            Me.cboCurrTax.Items.Clear()

            '2025 Core Tax Tahap 2.2 
            Me.txtDPPNilaiLain.Text = "0"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)

            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit

            Me.DGVdp.AllowUserToDeleteRows = False
            Me.DGVitemPR.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvExchageRates.ReadOnly = True
            Me.dgvCost.ReadOnly = True
            Me.dgvCost.AllowUserToDeleteRows = False
            Me.dgvCost.AllowUserToAddRows = False

            Me.llbAddress.Enabled = True
            Me.llbAddDP.Enabled = False
            Me.llbAddSR.Enabled = False


            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False

            Me.dgvItem.MultiSelect = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim ListPrimaryGRN As New List(Of Double)
        Dim GRNDummy As String = ""
        Dim rowDGV As DataGridViewRow
        Dim TanggalPerbandingan As Date = #1/1/1950#
        Dim Dummy As String = ""

        Dim ConvertedCurr As String = ""
        Dim row_dgv As DataGridViewRow
        Dim i As Integer = 0
        Dim TipeAkun As ClsMasterAkun_Purchase.EnumTipeAkun = ClsMasterAkun_Purchase.EnumTipeAkun.Asset
        Dim TipeAkun2 As ClsMasterAkun_Purchase.EnumTipeAkun2 = ClsMasterAkun_Purchase.EnumTipeAkun2.Account_Receivable
        Dim PrimaryAccountAP1, PrimaryAccountAP2 As Double ', PrimaryAccountAP3 As Double
        Dim ListPrimaryProduct As New List(Of Double)

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimaryProduct.Contains(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) = False Then
                    ListPrimaryProduct.Add(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value)
                End If
            Next
            If Me.ObjProduct.ApakahSudahVerifikasi(ListPrimaryProduct, True) = False Then
                Return False
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjAOP.TimeStamp <> Me.ObjAOP.GetTimestamp(Me.ObjAOP.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjAOP.CariDetail(Me.ObjAOP.Primary, Me.dgvItem, Me.DGVdp, _
                        Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("AOP number")
            KoleksiTextBox.Add(Me.TxtAOP)
            KoleksiNamaData.Add("supplier code")
            KoleksiTextBox.Add(Me.txtSupplierCode)
            KoleksiNamaData.Add("currency")
            KoleksiTextBox.Add(Me.txtCurrencyCode)
            KoleksiNamaData.Add("Supplier Invoice Number")
            KoleksiTextBox.Add(Me.txtSupplierInvoiceNo)
            If Me.chkTax.Checked = True Then
                KoleksiNamaData.Add("Supplier Tax Invoice Number")
                KoleksiTextBox.Add(Me.TxtTaxInvoiceNumber)
                KoleksiNamaData.Add("Supplier NPWP Number")
                KoleksiTextBox.Add(Me.TxtNPWPNo)
            End If

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'AOP tidak boleh mengambil GRN dari beda periode
            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimaryGRN.Contains(rowDGV.Cells(Me.ColPrimaryGRN.Index).Value) = False Then
                    'cek tanggal GRN
                    TanggalPerbandingan = Me.objGRN.GetTanggalTransaksi(rowDGV.Cells(Me.ColPrimaryGRN.Index).Value)
                    If Format(TanggalPerbandingan, "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                        If GRNDummy = "" Then
                            GRNDummy = rowDGV.Cells(Me.ColGRNNO.Index).Value
                        Else
                            GRNDummy = GRNDummy + ", " + rowDGV.Cells(Me.ColGRNNO.Index).Value
                        End If
                    End If
                    'tambahkan ke dalam list
                    ListPrimaryGRN.Add(rowDGV.Cells(Me.ColPrimaryGRN.Index).Value)
                End If
            Next
            If GRNDummy <> "" Then
                MsgBox("The GRN must be dated on the same month as AOP date. Please change your AOP date or fix following GRN data: " _
                + GRNDummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek jangan sampai ada DP yang terlupakan, khusus untuk order balance yang sudah habis
            Dim ij As New List(Of Double)
            Dim row As DataGridViewRow
            Dim DPNeeded As String = ""

            'tambahkan item DP
            Me.dgvDPDummy.Rows.Clear()
            For Each row In Me.dgvItem.Rows
                If ij.Contains(Me.dgvItem.Item(Me.ColPrimaryPO.Index, row.Index).Value) = False Then
                    ij.Add(Me.dgvItem.Item(Me.ColPrimaryPO.Index, row.Index).Value)
                End If
            Next
            Me.ObjAOP.IsiDgvDP(ij, Me.dgvDPDummy, Me.txtCurrencyCode.Text, Me.ObjCD, Me.TxtPrimaryTransaksiAOP.Text)
            If Me.dgvDPDummy.Rows.Count > Me.DGVdp.Rows.Count Then
                If MsgBox("There are down payment refering to this invoice, continue and ignore the downpayment?", _
                MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If


            'Account payable dan currency tax tidak boleh berbeda dengan data aop
            'Dim PrimaryPO As Double = -1
            'If IsNumeric(Me.txtPrimaryGRN.Text) AndAlso _
            '   IsNumeric(Me.txtPrimaryAccount.Text) AndAlso _
            '   IsNumeric(Me.txtPrimaryCurrTax.Text) Then
            '    If Me.ObjAOP.GetJmlAOPByPrimaryPO(Me.txtP.Text, _
            '                                   Me.txtPrimaryAccount.Text, _
            '                                   Me.txtPrimaryCurrTax.Text, ClsTransaksiAOP.EnumTipeAop.AOP) <= 0 Then
            '        MsgBox("There's is none of transaction AOP with PO No. " & Me.txtPONo.Text & _
            '        " using account payable " & Me.cboAccount.Text & " and currency tax " & Me.cboCurrTax.Text & _
            '        ", please choose similiar account payable and currency tax for journal needed.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            'Item DP yang dipilih harus memiliki account payable dan currency tax yang sama
            If DGVdp.Rows.Count > 0 AndAlso Not IsDBNull(DGVdp.Rows(0).Cells(0).Value) Then
                Dim ObjManualDP As New AccModule.ClsTransaksiBalanceManual_DP

                Dim strMsg As String = ""
                Dim strMsg2 As String = ""
                For Each row In Me.DGVdp.Rows
                    Dim PrimaryAOPDP As Double = -1
                    Dim PrimaryCurrTax As Double = -1
                    Dim PrimaryAccount As Double = -1
                    If row.Cells(Me.ColDPTipeDP.Index).Value = ClsTransaksiAOP.EnumTipeDP.AOPDP Then
                        PrimaryAOPDP = Me.ObjAOP.GetPrimary(row.Cells(Me.ColDP_AOP_DP_No.Index).Value)
                        PrimaryCurrTax = Me.ObjAOP.GetPrimaryCurrency(PrimaryAOPDP)
                        PrimaryAccount = Me.ObjAOP.GetPrimaryAccountAP(PrimaryAOPDP)
                    Else
                        PrimaryAOPDP = ObjManualDP.GetPrimary(row.Cells(Me.ColDP_AOP_DP_No.Index).Value, AccModule.ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP)
                        PrimaryCurrTax = ObjManualDP.GetPrimaryCurrency(PrimaryAOPDP)
                        PrimaryAccount = ObjManualDP.GetPrimaryAPAR_Acc(PrimaryAOPDP)
                    End If

                    If Me.txtPrimaryCurrency.Text <> PrimaryCurrTax Then 'Me.txtPrimaryAccount.Text <> PrimaryAccount OrElse 
                        If strMsg = "" Then
                            strMsg &= row.Cells(Me.ColDP_AOP_DP_No.Index).Value & " use currency " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
                        Else
                            strMsg &= vbNewLine & row.Cells(Me.ColDP_AOP_DP_No.Index).Value & _
                                      " use currency " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
                        End If
                    End If
                Next

                If strMsg <> "" Then
                    MsgBox("AOPDP No. " & strMsg & vbNewLine & " you just change account payable and currency " & _
                    "tax with similiar data for journal needed.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Item DP yang dipilih harus memiliki account payable dan currency tax yang sama
            'If Me.dgvPR.Rows.Count > 0 AndAlso Not IsDBNull(dgvPR.Rows(0).Cells(0).Value) Then
            '    Dim strMsg As String = ""
            '    Dim strMsg2 As String = ""
            '    Dim row2 As New DataGridViewRow
            '    For Each row2 In Me.dgvPR.Rows
            '        Dim PrimaryPR As Double = -1
            '        Dim PrimaryCurrTax As Double = -1
            '        Dim PrimaryAccount As Double = -1
            '        PrimaryPR = Me.ObjPR.GetPrimary(row2.Cells(Me.CopPRNo.Index).Value)
            '        PrimaryCurrTax = Me.ObjPR.GetPrimaryCurrTax(PrimaryPR)
            '        PrimaryAccount = Me.ObjPR.GetPrimaryAccountAP(PrimaryPR)

            '        If Me.txtPrimaryAccount.Text <> PrimaryAccount OrElse Me.txtPrimaryCurrTax.Text <> PrimaryCurrTax Then
            '            If strMsg = "" Then
            '                strMsg &= row2.Cells(Me.ColPRNumber.Index).Value & " use account payable " & _
            '                          ObjAkun.GetAccountCode(PrimaryAccount) & "-" & ObjAkun.GetDescription2(PrimaryAccount) & _
            '                          " and currency tax " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
            '            Else
            '                strMsg &= vbNewLine & row2.Cells(Me.ColPRNumber.Index).Value & _
            '                          " use account payable " & ObjAkun.GetAccountCode(PrimaryAccount) & _
            '                          "-" & ObjAkun.GetDescription2(PrimaryAccount) & _
            '                          " and currency tax " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
            '            End If
            '        End If
            '    Next

            '    If strMsg <> "" Then
            '        MsgBox("Purchase Return No. " & strMsg & vbNewLine & " you just change account payable and currency " & _
            '        "tax with similiar data for journal needed.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            'berkaitan dengan tanggal closing
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.ObjAOP.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjAOP.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtAOP.Text = Me.ObjAOP.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjAOP.TanggalTransaksi
                        Me.TxtAOP.Text = Me.ObjAOP.NoTransaksi
                        Return False
                    Else
                        Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                    End If
                End If
            End If

            'cek berkaitan dengan purchase return
            Dummy = Me.GetPRNumber
            If Dummy <> "" Then
                If MsgBox("The purchase return: " + Dummy + " refers to product in this AOP, ignore purchase return and continue process anyway?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            'jika currency yang digunakan adalah IDR, maka tax rate harus bernilai 1
            If Me.txtPrimaryCurrency.Text = "1" Then
                If CDec(Me.txtTaxExchangeRate.Text) <> 1 Then
                    MsgBox("For invoice with " + Me.ObjCurrency.GetCurrencyCode(1) + " currency, the tax exchange rate value must be 1.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'jika bukan IDR dan tax ratenya bernilai 1 maka berikan peringatan
                If CDec(Me.txtTaxExchangeRate.Text) = 1 Then
                    MsgBox("This invoice use non-" + Me.ObjCurrency.GetCurrencyCode(1) + " currency and the tax exchange rate value = 1. This is wrong data input.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If


            'If Me.chkTax.Checked = True Then
            '    If CDec(Me.txtTax.Text) <= 0 Then
            '        MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.txtTax.Focus()
            '        Return False
            '    End If
            '    If CDec(Me.txtTax.Text) <> 10 Then
            '        If MsgBox("Tax rate <> 10%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
            '            Me.txtTax.Focus()
            '            Return False
            '        End If
            '    End If
            'End If

            If Me.chkTax.Checked = True Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtTax.Focus()

                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
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

            If Me.dgvExchageRates.RowCount > 0 Then
                Dummy = ""
                i = 0
                row_dgv = Me.dgvExchageRates.Rows(0)
                If Me.txtCurrencyCode.Text.ToUpper <> row_dgv.Cells(Me.ColConvertedCurrecy.Index).Value.ToString.ToUpper Then
                    'berkaitan dengan exchange rate data, jika currecy yang digunakan tidak sama dengan exchange rate data, maka seharusnya nilainya
                    'tidak nol dan tidak 1
                    For Each row_dgv In Me.dgvExchageRates.Rows
                        If CDec(row_dgv.Cells(Me.ColExchangeRate.Index).Value) = 1 Then
                            i += 1
                            If Dummy = "" Then
                                Dummy = "The following document have exchange rate value = 1:" + vbCrLf + _
                                i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            Else
                                Dummy = Dummy + i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            End If
                        End If
                    Next
                Else
                    'sebaliknya, jika mata uangnya sama tetapi kemudian exchange rate yang digunakan adalah 1 maka data juga salah
                    For Each row_dgv In Me.dgvExchageRates.Rows
                        If CDec(row_dgv.Cells(Me.ColExchangeRate.Index).Value) <> 1 Then
                            i += 1
                            If Dummy = "" Then
                                Dummy = "The following document have exchange rate value <> 1:" + vbCrLf + _
                                i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            Else
                                Dummy = Dummy + i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            End If
                        End If
                    Next
                End If

                If Dummy <> "" Then
                    Dummy = Dummy + vbCrLf + "This is wrong data input."
                    MsgBox(Dummy, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'data ongkos harus memiliki keterangan yang jelas, dan tidak boleh bernilai nol
            For Each rowDGV In Me.dgvCost.Rows
                If rowDGV.Cells(Me.ColCostAmount.Index).Value = 0 And rowDGV.Index <> Me.dgvCost.RowCount - 1 Then
                    MsgBox("Please specify cost data on row " + (rowDGV.Index + 1).ToString + ". Cost amount must be a positive number", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvCost.ClearSelection()
                    rowDGV.Selected = True
                    Me.dgvCost.Focus()
                    Return False
                End If
            Next

            'jika lebih besar dari convertion date, maka harus ada account code
            If Me.ObjAOP.GetConvertionDate.Date <= Me.dtpDate.Value.Date Then
                If Me.dgvCost.Visible = True Then
                    For Each rowDGV In Me.dgvCost.Rows
                        If rowDGV.Cells(Me.ColPrimaryAccount.Index).Value = -1 And rowDGV.Cells(Me.ColCostAmount.Index).Value <> 0 Then
                            MsgBox("Please specify account code on row " + (rowDGV.Index + 1).ToString, _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvCost.ClearSelection()
                            rowDGV.Selected = True
                            Me.dgvCost.Focus()
                            Return False
                        End If
                    Next
                End If
            End If

            'cek data currency untuk DP
            Dummy = ""
            For Each row_dgv In Me.DGVdp.Rows
                If row_dgv.Cells(Me.ColDP_CurrencyDP.Index).Value.ToString <> Me.txtCurrencyCode.Text Then
                    If Dummy = "" Then
                        Dummy = row_dgv.Cells(Me.ColDP_AOP_DP_No.Index).Value.ToString
                    Else
                        Dummy = Dummy + ", " + row_dgv.Cells(Me.ColDP_AOP_DP_No.Index).Value.ToString
                    End If
                End If
            Next
            If Dummy <> "" Then
                MsgBox("The following downpayment used different currency with this AOP: " _
                + Dummy + ". We cannot enter the data.", MsgBoxStyle.Information, "Information")
            End If

            'update data payment terms
            Me.UpdatePaymentTerm()

            'nilai maksimal total sales adalah: 999,999,999,999
            If CDec(Me.TxtTotal.Text) * CDec(Me.txtTaxExchangeRate.Text) > 999999999999 Then
                MsgBox("Maximum sales value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'nilai pajak maksimal adalah 999,999,999,999
            If CDec(Me.txtTaxAmount.Text) > 999999999999 Then
                MsgBox("Maximum VAT value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'untuk setiap item cost, account code harus diisikan dan tipenya harusnya adalah asset atau payable
            For Each rowDGV In Me.dgvCost.Rows
                If rowDGV.Index <> Me.dgvCost.Rows.Count - 1 Then
                    'pastikan akunnya diisikan
                    If IsNumeric(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) = False Then
                        MsgBox("Please specify account code for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvCost.ClearSelection()
                        Me.dgvCost.Rows(rowDGV.Index).Selected = True

                        Return False
                    End If

                    'pastikan akunnya ada dan valid
                    If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(rowDGV.Cells(Me.ColAccounrCode.Index).Value) = False Then
                        MsgBox("Please specify valid account code for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvCost.ClearSelection()
                        Me.dgvCost.Rows(rowDGV.Index).Selected = True

                        Return False
                    Else
                        'pastikan tipe akunnya adalah asset(pre payment) atau hutang
                        TipeAkun = Me.ObjAkun.GetTipeAkun(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value)
                        If TipeAkun = ClsMasterAkun_Purchase.EnumTipeAkun.Asset Or TipeAkun = ClsMasterAkun_Purchase.EnumTipeAkun.Liabilities Then
                        Else
                            'kalo ga maka stop
                            MsgBox("Please specify account code with 'asset' or 'account payable' account type for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedIndex = 0
                            Me.dgvCost.ClearSelection()
                            Me.dgvCost.Rows(rowDGV.Index).Selected = True

                            Return False
                        End If

                        ''PrimaryCurrency = Me.ObjAkun.GetPrimaryCurrency(CDec(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value))
                        'If TipeAkun = ClsMasterAkun_Purchase.EnumTipeAkun.Liabilities Then
                        '    If Me.txtPrimaryCurrency.Text <> CDec(rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value) Then
                        '        Dummy = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        '        MsgBox("Please use account payable with " + Dummy + " currency on row " + _
                        '        (rowDGV.Index + 1).ToString + ".", MsgBoxStyle.Information, "Information")
                        '        Me.TabControl1.SelectedIndex = 0
                        '        Me.dgvCost.ClearSelection()
                        '        Me.dgvCost.Rows(rowDGV.Index).Selected = True
                        '        Return False
                        '    End If
                        'End If

                        TipeAkun2 = Me.ObjAkun.GetTipeAkun2(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value)
                        If TipeAkun2 = ClsMasterAkun_Purchase.EnumTipeAkun2.Account_Payable Then
                            'currency harus sama --> either currency pembayaran atau base currency
                            If Core_Module.PrimaryBaseCurrency <> CDec(rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value) And _
                            Me.txtPrimaryCurrency.Text <> CDec(rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value) Then
                                If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                                    Dummy = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                                Else
                                    Dummy = Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                                " or " + Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                                End If
                                MsgBox("Please use account payable with " + Dummy + " currency on row " + _
                                (rowDGV.Index + 1).ToString + ".", MsgBoxStyle.Information, "Information")

                                Me.TabControl1.SelectedIndex = 0
                                Me.dgvCost.ClearSelection()
                                Me.dgvCost.Rows(rowDGV.Index).Selected = True

                                Return False
                            End If

                            'jika tipenya account payable, maka primary akun harus primary akun ap untuk curency AOP atau primary akun ap dr IDR
                            'diganti per 30 maret 2016, account payable yang dipilih harus persis sama dengan account payable AOP
                            If IsNumeric(Me.txtPrimaryAccount.Text) Then
                                PrimaryAccountAP1 = CDbl(Me.txtPrimaryAccount.Text)
                            End If
                            'PrimaryAccountAP2 = Me.ObjCurrency.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)
                            'PrimaryAccountAP3 = Me.ObjCurrency.GetPrimaryAkunAP(rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value)

                            'If CDec(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP1 AndAlso _
                            '   CDec(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP2 Then
                            'CDec(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP3 Then
                            If CDec(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP1 Then
                                Dummy = ""
                                'If PrimaryAccountAP1 = PrimaryAccountAP2 Then
                                '    Dummy = "'" + Me.ObjAkun.GetDescription2(PrimaryAccountAP1) + "'"
                                'Else
                                '    Dummy = "'" + Me.ObjAkun.GetDescription2(PrimaryAccountAP1) + "' or '" + Me.ObjAkun.GetDescription2(PrimaryAccountAP2) + "'"
                                'End If
                                Dummy = "'" + Me.ObjAkun.GetDescription2(PrimaryAccountAP1) + "'"
                                MsgBox("Please use account payable: " + Dummy + " for cost data on row " + _
                                (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                Me.TabControl1.SelectedIndex = 0
                                Me.dgvCost.ClearSelection()
                                Me.dgvCost.Rows(rowDGV.Index).Selected = True
                                Return False
                            End If


                            If CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP1 Then
                                MsgBox("Please use account payable: '" + Me.ObjAkun.GetDescription2(PrimaryAccountAP1) + _
                                 "' for cost data on row " + _
                                 (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")

                                Me.TabControl1.SelectedIndex = 0
                                Me.dgvCost.ClearSelection()
                                Me.dgvCost.Rows(rowDGV.Index).Selected = True

                                Return False
                            End If

                            'jika account payablenya non idr/sama currency-nya dengan currency AOP, maka exchange ratenya harus 
                            'sama dengan tax exchange rate, jika tidak maka balance menjadi salah hitung
                            'jika kemudian hari ditemukan kasus seperti ini, maka yang harus dimodifikasi adalah clstransaksi aop pada bagian
                            'fungsi 'Transfer Jurnal'-nya
                            If CDec(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) = PrimaryAccountAP1 Then
                                If CDec(rowDGV.Cells(Me.ColExcRateCost.Index).Value) <> CDec(Me.txtTaxExchangeRate.Text) Then
                                    MsgBox("Please use exchange rate: " + Me.txtTaxExchangeRate.Text + _
                                           " for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                    Me.TabControl1.SelectedIndex = 0
                                    Me.dgvCost.ClearSelection()
                                    Me.dgvCost.Rows(rowDGV.Index).Selected = True
                                    Return False
                                End If
                            End If
                        Else
                            'jika tipenya bukan account payable maka currency harus idr
                            If Me.ObjAkun.GetPrimaryCurrency(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> Core_Module.PrimaryBaseCurrency Then
                                MsgBox("Please specify account code with " + Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                                " currency for cost data on row " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                Me.TabControl1.SelectedIndex = 0
                                Me.dgvCost.ClearSelection()
                                Me.dgvCost.Rows(rowDGV.Index).Selected = True
                                Return False
                            End If
                        End If
                    End If
                End If
            Next

            'nilai dp, tidak boleh melebih nilai invoice
            If CDec(Me.txtDownPayment.Text) > CDec(Me.txtSubTotalBeforeDP.Text) Then
                MsgBox("Maximum down payment amount is " + Me.txtSubTotalBeforeDP.Text + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.chkTax.Checked = True Then
                'kalo ada VAT, maka akun VAT out harus ada
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify VAT in account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '2025 Core Tax Tahap 2.2
                TxtTaxInvoiceNumber.Text = TxtTaxInvoiceNumber.Text.Trim
                If Me.TxtTaxInvoiceNumber.Text.StartsWith("01") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("02") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("03") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("04") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("06") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("07") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("08") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("09") = False And _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("10") = False Then
                    MsgBox("The tax invoice serial number must starts with one of the following alternatives: 01, 02, 03, 04, 06, 07, 08,09 or 10.", _
                    MsgBoxStyle.Information, "Information")
                    TxtTaxInvoiceNumber.Focus()
                    Return False
                End If

                If Me.TxtTaxInvoiceNumber.Text.StartsWith("07") = True Or _
                Me.TxtTaxInvoiceNumber.Text.StartsWith("08") = True Then
                    Dim ijk As New Sales_Module.FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur
                    ijk.TxtTaxInvoiceSeries.Text = Me.TxtTaxInvoiceNumber.Text.Substring(0, 3)
                    ijk.vIndexAcuan = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
                    ijk.ShowDialog()
                    Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = ijk.cboKeteranganTambahan.SelectedIndex
                End If
                'Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = Me.cboTaxType.SelectedText
                '2025 Core Tax Tahap 2.2
            End If

            'cek akun duty stamp
            If CDec(Me.TxtDutyStamp.Text) > 0 Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify dutystamp expense account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If CDec(Me.TxtDutyStamp.Text) > 0 Then
                If Me.txtPrimaryCurrency.Text <> "1" Then
                    MsgBox("Duty stamp is for IDR AOP only.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.ObjAOP.ApakahDPSudahDipakai(Me.TxtPrimaryTransaksiAOP.Text, True, Me) = True Then
                Return False
            End If

            '@at_amount tidak boleh terlalu kecil sehingga unit price per satuan terkecil barang tersebut menjadi bernilai 0
            If Me.CekATAmount = False Then
                Return False
            End If
            'cek konsistensi @ cost
            If Me.ApakahAtCostKonsisten = False Then
                Return False
            End If

            'Update NPWP 16 Digit
            If chkTax.Checked = True Then 'hanya dicek ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                '2025 Core Tax Tahap 2.1
                If ObjSetNPWP.SetNPWP_New(Me.dtpDate.Value.Date, Me.txtPrimarySupplier.Text, Me.TxtNPWPNo.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Supplier, _
                  frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionPurchase, Me.UserCode, Me.TxtPrimaryTransaksiAOP.Text) = False Then
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
            ListPrimaryGRN = Nothing
        End Try
    End Function

    Private Function CekATAmount() As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DefPrecIDR As Integer = 0
        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim TBL As New DseTransaksiAOP_TransfJurnal.DataTable1DataTable
        Dim row As DseTransaksiAOP_TransfJurnal.DataTable1Row
        Dim DefPrec As Integer = 0
        Dim QtyRef As Decimal = 0
        Dim PrimaryCodeNo As Double = -1 

        Try
            DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency

            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimaryCodeNo.Contains(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) = False Then
                    ListPrimaryCodeNo.Add(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value)
                End If
            Next
            If ListPrimaryCodeNo.Count > 0 Then
                For Each PrimaryCodeNo In ListPrimaryCodeNo
                    DefPrec = Me.ObjProduct.GetAngkaDecimal(PrimaryCodeNo)

                    row = TBL.NewDataTable1Row
                    row.Primary_Product = PrimaryCodeNo
                    row.DefPrec = DefPrec
                    TBL.Rows.Add(row)
                Next

                For Each rowDGV In Me.dgvItem.Rows
                    PrimaryCodeNo = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                    For Each row In TBL.Rows
                        If row.Primary_Product = PrimaryCodeNo Then
                            'cek at price pada qty ref
                            QtyRef = 1 / (10 ^ row.DefPrec)

                            'aul
                            Dim SecQty As Decimal
                            If IsNumeric(rowDGV.Cells(Me.ColSecQty.Index).Value) = True Then
                                SecQty = rowDGV.Cells(Me.ColSecQty.Index).Value
                            Else
                                SecQty = -1
                            End If
                            Dim QtyMulti_Return As Decimal = Me.GetJmlDGVPR_WithBatch1_Multi1(PrimaryCodeNo, CDbl(rowDGV.Cells(Me.ColPrimaryGRN.Index).Value), _
                            SecQty, rowDGV.Cells(Me.ColBatchNo.Index).Value, rowDGV.Cells(Me.ColPONO.Index).Value)
                            '----------------------
                            If CDec(rowDGV.Cells(Me.ColQtyMulti_Item.Index).Value) = QtyMulti_Return Then 'diubah 2020-08-18, tadi colTotalQty
                                If CDec(rowDGV.Cells(Me.Col_AtAmount.Index).Value) <> 0 Then
                                    Me.TabControl1.SelectedIndex = 1
                                    Me.dgvItem.ClearSelection()
                                    rowDGV.Selected = True

                                    MsgBox("Invalid at cost value on row: " + (rowDGV.Index + 1).ToString + vbCrLf + "at cost value in IDR must be 0", MsgBoxStyle.Information, "Information")

                                    Return False
                                End If
                            Else
                                If (QtyRef * CDec(rowDGV.Cells(Me.Col_AtAmount.Index).Value)) / CDec(rowDGV.Cells(Me.ColTotalQty.Index).Value) < (1 / 10 ^ DefPrecIDR) Then
                                    Me.TabControl1.SelectedIndex = 1
                                    Me.dgvItem.ClearSelection()
                                    rowDGV.Selected = True

                                    MsgBox("Invalid at cost value on row: " + (rowDGV.Index + 1).ToString + vbCrLf + "Minimum at cost value is IDR " + _
                                    Me.ObjCD.DigitGrouping(((1 / 10 ^ DefPrecIDR) * CDec(rowDGV.Cells(Me.ColTotalQty.Index).Value)) / QtyRef, DefPrecIDR), MsgBoxStyle.Information, "Information")

                                    Return False
                                End If
                            End If

                            Exit For
                        End If
                    Next
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            ListPrimaryCodeNo = Nothing
            TBL.Dispose()
        End Try
    End Function

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiAOP

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.TipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP
            If Me.Lokal = True Then
                ij.ForeignSupplier = 0
            Else
                ij.ForeignSupplier = 1
            End If
            ij.ShowDialog()

            If Not ij.PrimaryTransaksi = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjAOP.CariDetail(ij.PrimaryTransaksi, _
                        Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me) _
                        = True Then
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================

    'msgbox("belum")
    'Ingat jika GRN sudah masuk maka semua item dalam GRN tersebut harus di-AOP
    Private Sub dtpDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.GotFocus
        Me.dtpDate.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub dtpDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.LostFocus
        Me.dtpDate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
    End Sub

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2)
            If Me.TxtRemark.Enabled = True Then
                Me.Hitung_Total_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTax.ReadOnly = True Then
                    Me.txtTax.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.txtTax.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierInvoiceNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierInvoiceNo.GotFocus
        Try

            If Me.TxtRemark.Enabled = True Then
                If Me.txtSupplierInvoiceNo.ReadOnly = True Then
                    Me.txtSupplierInvoiceNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.txtSupplierInvoiceNo.BackColor = Me.BackColorAll
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierInvoiceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierInvoiceNo.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtSupplierInvoiceNo.ReadOnly = True Then
                    Me.txtSupplierInvoiceNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.txtSupplierInvoiceNo.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Col As Windows.Forms.DataGridViewColumn

        Try
            If vQuickCreateStockIn = False Then
                Me.LoadProcedure()
                Me.LblProgress.Visible = False
                Me.btnStopProcess.Visible = False
                Me.btnStopProcess.TabStop = False
                Me.prg1.Visible = False
            Else
                Me.txtSupplierInvoiceNo.Focus()
                Me.LblProgress.Visible = True
                Me.btnStopProcess.Visible = True
                Me.btnStopProcess.TabStop = True
                Me.prg1.Visible = True
            End If
            Me.ShowIcon = False

            Me.cboFillBy.Items.RemoveAt(1)
            Me.cboFillBy.Items.Add("Stock In Purchase")
            Me.dgvItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvItem.ColumnHeadersHeight = 25

            'Me.ColPriceByMultiUnit.Visible = True
            Me.llbAdjustDPP.Visible = False  '-> ini ditutup pada 2019-06-06 (iwan), karena update price by multi unit
            'kalo nanti dibuka lagi, harus dicek penerapannya jika menggunakan price tidak berdasarkan multi unit
            Me.txtSubTotalBeforeDP.Left = Me.txtSubTotalAfterDP.Left
            Me.txtSubTotalBeforeDP.Width = Me.txtSubTotalAfterDP.Width

            'For Each Col In Me.DGVitemPR.Columns
            '    Col.SortMode = DataGridViewColumnSortMode.NotSortable
            'Next
            For Each Col In Me.dgvPR.Columns
                Col.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtExchangeRateToUSD_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchangeRateToUSD.GotFocus
        txtExchangeRateToUSD.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtExchangeRateToUSD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExchangeRateToUSD.KeyDown
        Me.ObjCD.Number_KeyDown(txtExchangeRateToUSD, sender, e)
    End Sub

    Private Sub txtExchangeRateToUSD_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExchangeRateToUSD.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtExchangeRateToUSD, sender, e)
    End Sub

    Private Sub txtExchangeRateToUSD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchangeRateToUSD.LostFocus
        txtExchangeRateToUSD.BackColor = Me.txtPrimarySupplier.BackColor
    End Sub

    Private Sub txtExhangeRateToIDR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtExhangeRateToIDR.Enabled = True Then
                    If Me.txtExhangeRateToIDR.ReadOnly = True Then
                        Me.txtExhangeRateToIDR.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.txtExhangeRateToIDR.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtExhangeRateToIDR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.Leave
        Try
            Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtExhangeRateToIDR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.LostFocus
        Try
            If Me.txtExhangeRateToIDR.Enabled = True Then
                If Me.txtExhangeRateToIDR.ReadOnly = True Then
                    txtExhangeRateToIDR.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtExhangeRateToIDR.BackColor = Me.txtPrimarySupplier.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            Me.cboDO.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPPNDipungut_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPPNDipungut.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.Hitung_Total_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetPaymentDue() As Decimal
        Dim ListPrimaryPO As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim PrimaryPO As Decimal = -1
        Dim PaymentDue As Decimal = -1
        Dim Dummy As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimaryPO.Contains(rowDGV.Cells(Me.ColPrimaryPO.Index).Value) = False Then
                    ListPrimaryPO.Add(rowDGV.Cells(Me.ColPrimaryPO.Index).Value)
                End If
            Next

            For Each PrimaryPO In ListPrimaryPO
                Dummy = Me.ObjPO.GetPaymentDue(PrimaryPO)
                If PaymentDue = -1 Then
                    PaymentDue = Dummy
                Else
                    If Dummy < PaymentDue Then
                        PaymentDue = Dummy
                    End If
                End If
            Next

            Return PaymentDue
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryPO = Nothing
        End Try
    End Function

    Private Sub UpdatePaymentTerm()
        Dim PaymentDue As Decimal = 0

        Try
            PaymentDue = Me.GetPaymentDue

            If CDec(Me.TxtPaymentsTerms.Text) <> PaymentDue Then
                If MsgBox("Update payment terms?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                    Me.TxtPaymentsTerms.Text = PaymentDue
                    Me.TabControl1.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadProcedure()
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvExchageRates.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVdp.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDPDummy.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCost.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVitemPR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvPR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ObjAOP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP

            If Me.Lokal = True Then
                Me.ColATCost.Visible = True
                Me.ColShowDetail.Visible = True
            Else
                Me.ColATCost.Visible = True
                Me.ColShowDetail.Visible = True
            End If
            Me.cboPPNDipungut.SelectedIndex = 0
            Me.cboFillBy.SelectedIndex = 0

            Me.cboCalculateBy.SelectedIndex = 0
            Me.SplitContainer2.Panel1Collapsed = False
            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.SplitContainer3.TabStop = False
            Me.SplitContainer4.TabStop = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Try
    '        If Me.vQuickCreateStockIn = False Then
    '            Me.LoadProcedure()
    '        Else
    '            Me.txtSupplierInvoiceNo.Focus()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

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
                        'If Me.SedangRevisi = True Then
                        'If Me.ObjAOP.SimpanRevisi(Me.dgvItem, Me.DGVdp, Me.DGVitemSR, _
                        'Me.dgvExchageRates, Me.UserCode, Me.ObjAOP.Primary, Me.dgvCost) Then
                        'MsgBox("Invoice Revised", MsgBoxStyle.Information, "Information")
                        ' Me.SalinKeField()
                        ' Me.NonAktifkanDetail()
                        'Me.btnFind.Focus()
                        'End If
                        'Else
                        If Me.ObjAOP.SimpanBaru(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                                Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()

                            If Me.vQuickCreateStockIn = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                    'End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjAOP.SimpanPerubahan(Me.dgvItem, Me.DGVdp, _
                        Me.DGVitemPR, Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me) Then
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
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
                If Me.vQuickCreateStockIn = True Then
                    Me.Close()
                End If
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.ObjAOP.CariDetail(Me.TxtPrimaryTransaksiAOP.Text, Me.dgvItem, _
                        Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                If Me.vQuickCreateStockIn = True Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtAOP.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOP.Text <> "" Then
                If Me.ObjAOP.ApakahBolehEdit(Me.ObjAOP.Primary, Me.UserCode, True) = True Then
                    If Me.ObjAOP.ApakahQuickPurchase(Me.TxtPrimaryTransaksiAOP.Text) = True Then
                        MsgBox("This data is quick purchase. Use quick purchase form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.baru = False
                    ' Me.SedangRevisi = False
                    Me.ObjAOP.CariDetail(Me.ObjAOP.Primary, Me.dgvItem, Me.DGVdp, _
                      Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtAOP.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOP.Text <> "" Then
                If Me.ObjAOP.ApakahBolehHapus(Me.ObjAOP.Primary, True, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    If Me.ObjAOP.ApakahQuickPurchase(Me.TxtPrimaryTransaksiAOP.Text) = True Then
                        MsgBox("This data is quick purchase. Use quick purchase form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Dim NoInvLama As String = Me.TxtAOP.Text.Substring(0, 19)
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjAOP.Hapus(Me.TxtPrimaryTransaksiAOP.Text, Me.UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            'jika masih ada invoce yang revisi selanjutnya maka tampilkan
                            Dim PrimaryInvBaru As Decimal = Me.ObjAOP.GetPrimaryByNoTransaksiLIKE(NoInvLama)
                            If PrimaryInvBaru <> -1 Then
                                If MsgBox("Load previous revision?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Me.ObjAOP.CariDetail(PrimaryInvBaru, Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                                    Me.dgvExchageRates, Me.dgvCost, Me)
                                    Me.SalinKeField()
                                End If
                            Else
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            End If
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
            If Me.LangsungClose = True Then
                Exit Sub
            End If
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            'If Me.SedangRevisi = True Then
                            'If Me.ObjAOP.SimpanRevisi(Me.dgvItem, Me.DGVdp, Me.DGVitemSR, _
                            'Me.dgvExchageRates, Me.UserCode, Me.ObjAOP.Primary) = True Then
                            'MsgBox("Invoice Revised", MsgBoxStyle.Information, "Information")
                            'Me.SalinKeField()
                            ' Me.NonAktifkanDetail()
                            ' Me.btnFind.Focus()
                            'Else
                            'e.Cancel = True
                            'End If
                            'Else
                            If Me.ObjAOP.SimpanBaru(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                            Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                                If Me.vQuickCreateStockIn = True Then
                                    Me.Close()
                                End If
                            Else
                                e.Cancel = True
                            End If
                        Else
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjAOP.SimpanPerubahan(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                            Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me) = True Then
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
                ElseIf hasil = MsgBoxResult.No Then
                    If Me.vQuickCreateStockIn = True Then
                        Me.LangsungClose = True
                        Me.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim dummyPrimaryCurrency As Decimal = -1
        Dim DefaulPres As Decimal = 0
        Dim IsMultiUnit As Boolean = False
        Dim AdaPriceByDefaultUnit As Boolean = False

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'format dp
            For Each row In Me.DGVdp.Rows
                row.Cells(Me.ColDP_DPAMount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColDP_DPAMount.Index).Value)
            Next
            'yang diformat dgv-nya:
            'For Each row In Me.dgvItem.Rows
            '    If dummyPrimaryCurrency <> row.Cells(Me.ColPrimaryCurrency.Index).Value Then
            '        DefaulPres = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            '        dummyPrimaryCurrency = row.Cells(Me.ColPrimaryCurrency.Index).Value
            '    End If

            '    'digit grouping: qty, secondary qty,unit price, nett price, amount
            '    If IsNumeric(row.Cells(Me.ColTotalQty.Index).Value) = True Then
            '        row.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value)
            '    End If
            '    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
            '        If row.Cells(Me.ColSecQty.Index).Value > 0 Then
            '            row.Cells(Me.ColSecQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecQty.Index).Value)
            '        End If
            '    End If
            '    If IsNumeric(row.Cells(Me.ColDiscount.Index).Value) = True Then
            '        row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColDiscount.Index).Value, _
            '        2))
            '    End If
            '    If IsNumeric(row.Cells(Me.ColUnitPrice.Index).Value) = True Then
            '        row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value, _
            '        DefaulPres))
            '    End If
            '    If IsNumeric(row.Cells(Me.ColNettPrice.Index).Value) = True Then
            '        row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColNettPrice.Index).Value, _
            '        DefaulPres))
            '    End If
            '    If IsNumeric(row.Cells(Me.ColAmount.Index).Value) = True Then
            '        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColAmount.Index).Value, _
            '        DefaulPres))
            '    End If
            'Next

            'yang diformat dgv-nya:
            Dim TidakAdaSecondary As Boolean = True
            Dim TidakAdaBatchNumber As Boolean = True

            For Each row In Me.dgvItem.Rows
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                If row.Cells(Me.ColSecQty.Index).Value = "-1" Or row.Cells(Me.ColSecUnit.Index).Value = "0" Then
                    row.Cells(Me.ColSecQty.Index).Value = "-"
                    row.Cells(Me.ColSecUnit.Index).Value = "-"
                End If
                If row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    TidakAdaBatchNumber = False
                End If
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    AdaPriceByDefaultUnit = True
                End If
            Next
            'tidak ada batch no
            If TidakAdaBatchNumber = True Then
                Me.ColBatchNo.Visible = False
            Else
                Me.ColBatchNo.Visible = True
            End If
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = True
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = True
            End If
            'untuk exchange rate 
            For Each row In Me.dgvExchageRates.Rows
                'digit grouping exchange rate menurut mata uangnya
                row.Cells(Me.ColExchangeRate.Index).Value = _
                Me.ObjCD.DigitGrouping(row.Cells(Me.ColExchangeRate.Index).Value)
                'digit grouping cost amount menurut rupiah
            Next
            If AdaPriceByDefaultUnit = False Then
                Me.ColPriceByMultiUnit_str.Visible = False
            Else
                Me.ColPriceByMultiUnit_str.Visible = True
            End If

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            ''@rosadinnn
            ''@MUaT
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnit.Index).Value <> .Cells(Me.ColPrimaryUnitsMulti_Item.Index).Value Then
                        IsMultiUnit = True
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.ColTotalQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc_Item.Visible = False
            Else
                Me.ColTotalQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc_Item.Visible = True
            End If

            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Try
            If e.ColumnIndex = Me.ColShowDetail.Index Then
                If e.RowIndex <> -1 Then
                    Me.ShowDetailCalculation(Me.dgvItem.Item(Me.ColCodeNo.Index, _
                    e.RowIndex).Value, e.RowIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellErrorTextChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellErrorTextChanged

    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Dim PrimaryPO As Double = -1
        Dim row, row1 As DataGridViewRow
        Dim i As Integer = 0
        Dim Ketemu As Boolean = False
        Dim PrimaryStockIn As Double = -1

        Try
            'hapus semua item yang sama primary DO-nya
            While i <= Me.dgvItem.Rows.Count - 1
                row = Me.dgvItem.Rows(i)
                If row.Cells(Me.ColPrimaryGRN.Index).Value = Me.buffer Then
                    Me.dgvItem.Rows.Remove(row)
                Else
                    i += 1
                End If
            End While
            'tambahkan data DO ke dalam tabel sumber
            PrimaryStockIn = Me.objGRN.GetPrimaryStockIn(Me.buffer)
            Me.TBL.Rows.Add(Me.buffer, Me.objGRN.GetNoGRN(CDec(Me.buffer)), _
            Me.ObjStockIn.GetSupplierDONumber(PrimaryStockIn), Me.ObjStockIn.GetNoTransaksi(PrimaryStockIn))

            'hapus data sales return yang merujuk ke DO ybs
            Me.IsiDGVReturn()

            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)

            If Me.dgvItem.RowCount = 0 Then
                Me.txtSupplierCode.ReadOnly = False
                Me.txtSupplierCode.BackColor = Me.TxtRemark.BackColor
                Me.btnFindSupplier.Enabled = True
            End If

            'pastikan setiap item dp, memiliki padanan pada item
            If Me.DGVdp.RowCount > 0 Then
                For ij As Integer = Me.DGVdp.Rows.Count - 1 To 0 Step -1
                    row = Me.DGVdp.Rows(ij)
                    PrimaryPO = Me.ObjAOP.GetPrimaryPO4DP(row.Cells(Me.ColDP_PrimaryAOPDP.Index).Value)
                    Ketemu = False
                    For Each row1 In Me.dgvItem.Rows
                        If row1.Cells(Me.ColPrimaryPO.Index).Value = PrimaryPO Then
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        Me.DGVdp.Rows.RemoveAt(ij)
                    End If
                Next
                Me.ObjCD.SetNumber(Me.DGVdp)
            End If

            ''@rosadinnn
            ''@MUaT
            FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTax.ReadOnly = True Then
                    Me.txtTax.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.txtTax.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.ReadOnly = True Then
                    TxtRemark.BackColor = Me.txtReadOnly.BackColor
                Else
                    TxtRemark.BackColor = BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.ReadOnly = True Then
                    TxtRemark.BackColor = Me.txtReadOnly.BackColor
                Else
                    TxtRemark.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Enter
        Try
            Me.buffer = Me.txtPrimarySupplier.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtSupplierCode.ReadOnly = True Then
                    txtSupplierCode.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtSupplierCode.BackColor = BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtSupplierCode.ReadOnly = True Then
                    txtSupplierCode.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtSupplierCode.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtCurrencyCode.ReadOnly = True Then
                    txtCurrencyCode.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtCurrencyCode.BackColor = BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtCurrencyCode.ReadOnly = True Then
                    txtCurrencyCode.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtCurrencyCode.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlDGVPR_WithBatch1_Default1(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryGRN As Decimal, _
    ByVal vSecQty As Decimal, ByVal vBatchNo As String, ByVal vPONO As String) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvPR.Rows
                If IsNumeric(row.Cells(Me.ColSecQty_PR.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN And _
                    row.Cells(Me.ColSecQty_PR.Index).Value = vSecQty And _
                    row.Cells(Me.ColBatchNo_PR.Index).Value = vBatchNo And _
                    row.Cells(Me.ColPONo_PR.Index).Value = vPONO Then
                        Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN And _
                    row.Cells(Me.ColBatchNo_PR.Index).Value = vBatchNo And _
                    row.Cells(Me.ColPONo_PR.Index).Value = vPONO Then
                        Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJmlDGVPR_WithBatch1_Multi1(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryGRN As Decimal, _
            ByVal vSecQty As Decimal, ByVal vBatchNo As String, ByVal vPONo As String) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvPR.Rows
                If IsNumeric(row.Cells(Me.ColSecQty_PR.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN And _
                    row.Cells(Me.ColSecQty_PR.Index).Value = vSecQty And _
                    row.Cells(Me.ColBatchNo_PR.Index).Value = vBatchNo And _
                    row.Cells(Me.ColPONo_PR.Index).Value = vPONo Then
                        Jml = Jml + CDec(row.Cells(Me.ColQtyMulti_PR.Index).Value) ''-->>>>>>>
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN And _
                    row.Cells(Me.ColBatchNo_PR.Index).Value = vBatchNo And _
                    row.Cells(Me.ColPONo_PR.Index).Value = vPONo Then
                        Jml = Jml + CDec(row.Cells(Me.ColQtyMulti_PR.Index).Value) ''-->>>>>>>
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function GetJmlDGVPR1(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryGRN As Decimal, _
    'ByVal vSecQty As Decimal) As Decimal
    '    Try
    '        Dim row As DataGridViewRow
    '        Dim Jml As Decimal = 0

    '        For Each row In Me.dgvPR.Rows
    '            If IsNumeric(row.Cells(Me.ColSecQty_PR.Index).Value) = True Then
    '                If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
    '                row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN And _
    '                row.Cells(Me.ColSecQty_PR.Index).Value = vSecQty Then
    '                    Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
    '                End If
    '            Else
    '                If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
    '                row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN Then
    '                    Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
    '                End If
    '            End If
    '        Next

    '        Return Jml
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Private Sub Hitung_Amount1_Default(ByVal row As DataGridViewRow)
    '    Dim DefaultPrecisionRow As Integer
    '    Try
    '        DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
    '        'hitung nett price
    '        'row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value * _
    '        '(100 - row.Cells(Me.ColDiscount.Index).Value) / 100, DefaultPrecisionRow))

    '        'hitung amount
    '        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
    '            row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQty.Index).Value * _
    '            row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecisionRow))
    '        Else
    '            If CDec(row.Cells(Me.ColSecQty.Index).Value) > 0 Then
    '                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQty.Index).Value * _
    '                 row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value, DefaultPrecisionRow))
    '            Else
    '                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQty.Index).Value * _
    '                row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecisionRow))
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub Hitung_Amount1_Multi(ByVal row As DataGridViewRow)
        Dim DefaultPrecisionRow As Integer
        Try
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            'row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value * _
            '(100 - row.Cells(Me.ColDiscount.Index).Value) / 100, DefaultPrecisionRow))

            'hitung amount
            If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColQtyMulti_Item.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value), DefaultPrecisionRow))
                Else
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColTotalQty.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value), DefaultPrecisionRow))
                End If
                'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColQtyMulti_Item.Index).Value * _
                'row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value), DefaultPrecisionRow))
            Else
                If CDec(row.Cells(Me.ColSecQty.Index).Value) > 0 Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColQtyMulti_Item.Index).Value * _
                     row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value), DefaultPrecisionRow))
                Else
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColQtyMulti_Item.Index).Value * _
                        row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value), DefaultPrecisionRow))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColTotalQty.Index).Value * _
                        row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value), DefaultPrecisionRow))
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetSubTotal_TanpaReturn_Original_Cur() As Double
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Double = 1
        Dim DefaultPrecisionRow As Double = 1

        Try
            'hitung sub total
            For Each row In Me.dgvItem.Rows
                'hitung purchase return jika ada
                Dim SubTotal1 As Double = 0
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                Dim SecQty As Double
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    SecQty = row.Cells(Me.ColSecQty.Index).Value
                Else
                    SecQty = -1
                End If
                Dim JmlReturn As Double = 0

                If SecQty <= 0 Then
                    If row.Cells(Me.ColTotalQty.Index).Value = 1 Then
                        SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                         row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
                    Else
                        SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                         row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
                    End If
                Else
                    SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                    row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, DefaultPrecisionRow))
                End If

                SubTotal += Me.ObjCD.SetDecimal(SubTotal1, DefaultPrecisionRow)
            Next

            Return Me.ObjCD.SetDecimal(SubTotal, DefaultPrecisionRow)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetSubTotal_TanpaReturn_Multi() As Decimal
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0
        Dim exchangRate As Decimal = 0

        Try
            'hitung sub total
            For Each row In Me.dgvItem.Rows
                'yang diperhitungkan hanya yang tidak pake secondary unit
                If Me.cboCalculateBy.SelectedIndex = 0 Then
                    'berdasarkan sales invoice
                    exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
                ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
                    'berdasarkan delivery order no
                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
                ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
                    'berdasarkan saler order no
                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
                End If

                'hitung purchase return jika ada
                Dim DefaultPrecisionRow As Decimal = 1
                Dim SubTotal1 As Decimal = 0
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                Dim SecQty As Decimal
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    SecQty = row.Cells(Me.ColSecQty.Index).Value
                Else
                    SecQty = -1
                End If
                Dim JmlReturn As Decimal = 0

                If SecQty <= 0 Then
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                         row.Cells(Me.ColNettPrice.Index).Value) + row.Cells(Me.ColItem_Adjustment.Index).Value) * exchangRate, DefaultPrecisionRow))
                    Else
                        SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                         row.Cells(Me.ColNettPrice.Index).Value) + row.Cells(Me.ColItem_Adjustment.Index).Value) * exchangRate, DefaultPrecisionRow))
                    End If
                Else
                    SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                     row.Cells(Me.ColNettPrice.Index).Value) + row.Cells(Me.ColItem_Adjustment.Index).Value) * row.Cells(Me.ColSecQty.Index).Value * exchangRate, DefaultPrecisionRow))
                End If

                SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
            Next

            Return Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Function GetSubTotal_TanpaReturn_Default() As Decimal
    '    Dim row As DataGridViewRow
    '    Dim SubTotal As Decimal = 0
    '    Dim exchangRate As Decimal = 0

    '    Try
    '        'hitung sub total
    '        For Each row In Me.dgvItem.Rows
    '            'yang diperhitungkan hanya yang tidak pake secondary unit
    '            If Me.cboCalculateBy.SelectedIndex = 0 Then
    '                'berdasarkan sales invoice
    '                exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
    '            ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
    '                'berdasarkan delivery order no
    '                exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
    '            ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
    '                'berdasarkan saler order no
    '                exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
    '            End If

    '            'hitung purchase return jika ada
    '            Dim DefaultPrecisionRow As Decimal = 1
    '            Dim SubTotal1 As Decimal = 0
    '            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
    '            Dim SecQty As Decimal
    '            If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
    '                SecQty = row.Cells(Me.ColSecQty.Index).Value
    '            Else
    '                SecQty = -1
    '            End If
    '            Dim JmlReturn As Decimal = 0

    '            If SecQty <= 0 Then
    '                SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
    '                 row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
    '            Else
    '                SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
    '                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, DefaultPrecisionRow))
    '            End If

    '            SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
    '        Next

    '        Return Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Function GetExchangeRate(ByVal vPrimaryTransaksi As Decimal) As Decimal
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvExchageRates.Rows
                If row.Cells(ColPrimaryDokumen.Index).Value = vPrimaryTransaksi Then
                    Return row.Cells(ColExchangeRate.Index).Value
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub Hitung_Total_Default()
    '    Dim row As DataGridViewRow
    '    Dim SubTotal As Decimal = 0
    '    Dim exchangRate As Decimal = 0
    '    Dim IDR_DefPrec As Integer = 0

    '    Try
    '        IDR_DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

    '        If Me.dgvExchageRates.Rows.Count > 0 Then
    '            'hitung sub total
    '            For Each row In Me.dgvItem.Rows
    '                'yang diperhitungkan hanya yang tidak pake secondary unit
    '                If Me.cboCalculateBy.SelectedIndex = 0 Then
    '                    'berdasarkan sales invoice
    '                    exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
    '                ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
    '                    'berdasarkan delivery order no
    '                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
    '                ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
    '                    'berdasarkan saler order no
    '                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
    '                End If

    '                'hitung purchase return jika ada
    '                Dim DefaultPrecisionRow As Decimal = 1
    '                Dim SubTotal1 As Decimal = 0
    '                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
    '                Dim SecQty As Decimal
    '                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
    '                    SecQty = row.Cells(Me.ColSecQty.Index).Value
    '                Else
    '                    SecQty = -1
    '                End If
    '                Dim JmlReturn As Decimal = 0
    '                JmlReturn = Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '                row.Cells(Me.ColPrimaryGRN.Index).Value, SecQty, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)


    '                If SecQty <= 0 Then
    '                    SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
    '                     row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))
    '                Else
    '                    SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
    '                    row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, DefaultPrecisionRow))
    '                End If

    '                SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
    '            Next
    '            Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
    '            'hitung DP
    '            Me.txtDownPayment.Text = "0"
    '            For Each row In Me.DGVdp.Rows
    '                Me.txtDownPayment.Text = CDec(Me.txtDownPayment.Text) + CDec(row.Cells(ColDP_DPAMount.Index).Value)
    '            Next
    '            Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtDownPayment.Text, DefaultPrecision))
    '            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), DefaultPrecision))
    '            'Me.txtSubTotalAfterDP.Text = Me.txtSubTotalBeforeDP.Text

    '            'hitung pajak
    '            If Me.chkTax.Checked = True Then
    '                If Me.cboPPNDipungut.SelectedIndex = 0 Then
    '                    'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
    '                    If Me.txtPrimaryCurrency.Text = "1" Then
    '                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
    '                        CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), Me.ObjCurrency.GetDefaultPrecission(1))
    '                    Else
    '                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
    '                        CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), Me.ObjCurrency.GetDefaultPrecission(1))
    '                    End If
    '                Else
    '                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
    '                End If
    '            Else
    '                Me.txtTax.Text = "0.00"
    '                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
    '            End If

    '            'hitung total = sub total + taxj 
    '            If Me.txtPrimaryCurrency.Text = "1" Then
    '                'kalo IDR, duty stamp diperhitungkan
    '                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text) + _
    '                CDec(Me.txtTaxAmount.Text) + CDec(Me.TxtDutyStamp.Text), DefaultPrecision))
    '            Else
    '                'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
    '                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalAfterDP.Text), _
    '                DefaultPrecision))
    '            End If
    '        End If
    '        'hitung ulang unit price
    '        Me.Hitung_AtCost_Default()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Sub SetDefaultPrecision(Optional ByVal vHitungTotal As Boolean = True)
        Try
            If Me.txtPrimaryCurrency.Text = "1" Then
                Me.txtExhangeRateToIDR.Text = "1"
                Me.txtExhangeRateToIDR.ReadOnly = True
                Me.txtExhangeRateToIDR.BackColor = Me.txtReadOnly.BackColor
            Else
                Me.txtExhangeRateToIDR.ReadOnly = False
                Me.txtExhangeRateToIDR.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
            End If

            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
            Me.FormatDGV()
            If vHitungTotal = True Then
                Me.Hitung_Total_Multi()
            End If
            Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotalBeforeDP.Text, DefaultPrecision))
            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotalAfterDP.Text, DefaultPrecision))
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxAmount.Text, Me.ObjCurrency.GetDefaultPrecission(1)))
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim Dummy As String = ""
        Try
            If Me.chkTax.Checked = True Then
                Me.txtTax.ReadOnly = False
                Me.txtTax.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                Me.txtTax.TabStop = True
                Me.txtTaxExchangeRate.ReadOnly = False
                Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                Me.txtTaxExchangeRate.TabStop = True
                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                'If CDec(Me.txtTax.Text) = 0 Then
                '    Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
                'End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2)
            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.BackColor = Me.txtReadOnly.BackColor
                Me.txtTax.TabStop = False
                Me.txtTaxExchangeRate.ReadOnly = True
                Me.txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Me.txtTaxExchangeRate.TabStop = False
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", 2)
            End If

            If Me.TxtRemark.Enabled = True Then
                Me.SetTaxRate()
                Me.Hitung_Total_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetTaxRate(Optional ByVal vSetNewTaxRate As Boolean = True)
        Dim Rate As Decimal = 0

        Try
            If Me.chkTax.Checked = True Then
                If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                    Me.txtTaxExchangeRate.Enabled = False
                    Me.txtTaxExchangeRate.ReadOnly = True
                    Me.txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    Me.txtTaxExchangeRate.TabStop = False
                    If vSetNewTaxRate = True Then
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Me.ObjCurrency.GetDefaultPrecission(1))
                    End If
                Else
                    Me.txtTaxExchangeRate.Enabled = True
                    Me.txtTaxExchangeRate.ReadOnly = False
                    Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    Me.txtTaxExchangeRate.TabStop = True
                    If vSetNewTaxRate = True Then
                        Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                        If Rate = 0 Then
                            Rate = Me.ObjCurrency.GetDefaultExhangeRate(Me.txtPrimaryCurrency.Text)
                            If Rate = 0 Then
                                Rate = 1
                            End If
                        End If
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Me.ObjTaxRate.DefaultPrec)
                    End If
                End If
            Else
                If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                    Me.txtTaxExchangeRate.Enabled = False
                    Me.txtTaxExchangeRate.ReadOnly = True
                    Me.txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    Me.txtTaxExchangeRate.TabStop = False
                    If vSetNewTaxRate = True Then
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Me.ObjCurrency.GetDefaultPrecission(1))
                    End If
                Else
                    Me.txtTaxExchangeRate.Enabled = True
                    Me.txtTaxExchangeRate.ReadOnly = False
                    Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    Me.txtTaxExchangeRate.TabStop = True
                    If vSetNewTaxRate = True Then
                        Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                        If Rate = 0 Then
                            Rate = Me.ObjCurrency.GetDefaultExhangeRate(Me.txtPrimaryCurrency.Text)
                            If Rate = 0 Then
                                Rate = 1
                            End If
                        End If
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Me.ObjTaxRate.DefaultPrec)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        If Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Dim ApakahLokal As Boolean = False
        Dim IndexStr, IndexStr1 As Integer
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""

        Dim Reset As Boolean = False
        Try
            If e.KeyValue = 13 Then
                If Me.Objsupplier.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                    Me.txtPrimarySupplier.Text = Me.Objsupplier.GetPrimary(Me.txtSupplierCode.Text)
                    Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                    Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                    '                    Me.txtNoSeriFakturPajak.Text = Me.Objsupplier.GetTaxCode(Me.txtPrimarySupplier.Text)
                    'Me.txtNoNPWP.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.SelectedIndex = 0
                    Me.IsiCboDeliveryOrder(Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.Focus()
                Else
                    'ada 2 Supplier dengan kode yang sama atau tidak ada Supplier yang dimaksud
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.Objsupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                        '                        Me.txtNoSeriFakturPajak.Text = Me.Objsupplier.GetTaxCode(Me.txtPrimarySupplier.Text)
                        '                        Me.txtNoNPWP.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                        Me.cboSupplierAddress.SelectedIndex = 0
                        Me.IsiCboDeliveryOrder(Me.txtPrimarySupplier.Text)
                        Me.cboSupplierAddress.Focus()

                        'ganti no transaksi
                        If Me.baru = True Then
                            Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                        Else
                            ApakahLokal = Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                            If ApakahLokal = True Then
                                IndexStr = Me.TxtAOP.Text.IndexOf("L")
                                IndexStr1 = Me.TxtAOP.Text.IndexOf("I", 12)
                            Else
                                IndexStr = Me.TxtAOP.Text.IndexOf("I", 12)
                                IndexStr1 = Me.TxtAOP.Text.IndexOf("L")
                            End If
                            If IndexStr = -1 Then
                                Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                            End If
                        End If
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.cboSupplierAddress.Items.Clear()
                        Me.txtSupplierAddIdentifier.Text = ""
                        Me.txtPrimarySupplier.Text = "-1"
                        '     Me.txtNoNPWP.Clear()
                        '         Me.txtNoSeriFakturPajak.Clear()
                        Me.cboDO.DataSource = Nothing
                        Me.txtSupplierCode.Focus()
                    End If
                End If

                'cek supplier
                If Me.txtSupplierCode.Text <> "" Then
                    If Me.Lokal = True Then
                        If Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text) = False Then
                            Reset = True
                        End If
                    Else
                        If Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text) = True Then
                            Reset = True
                        End If
                    End If

                    If Reset = True Then
                        If Me.Lokal = True Then
                            MsgBox("Please choose local supplier.", MsgBoxStyle.Information, "Information")
                        Else
                            MsgBox("Please choose foreign supplier.", MsgBoxStyle.Information, "Information")
                        End If

                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.cboSupplierAddress.Items.Clear()
                        Me.txtSupplierAddIdentifier.Text = ""
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.cboDO.DataSource = Nothing
                        Me.txtSupplierCode.Focus()
                    End If
                End If
            End If
            Me.TxtNPWPNo.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)

            '2025 Core Tax Tahap 2.1
            If Me.dtpDate.Value >= Me.ObjSetNPWP.Tglberlaku.Date Then
                Me.TxtNPWPNo.Text = Me.ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        If Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Dim ApakahLokal As Boolean = False
        Dim IndexStr, IndexStr1 As Integer
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""
        Dim reset As Boolean = False

        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.Objsupplier.GetJml(Me.txtSupplierCode.Text) = 1 Then
                    Me.txtPrimarySupplier.Text = Me.Objsupplier.GetPrimary(Me.txtSupplierCode.Text)
                    Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                    Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                    '                    Me.txtNoSeriFakturPajak.Text = Me.Objsupplier.GetTaxCode(Me.txtPrimarySupplier.Text)
                    '                    Me.txtNoNPWP.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.SelectedIndex = 0
                    Me.IsiCboDeliveryOrder(Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.Focus()

                    'ganti no transaksi
                    If Me.baru = True Then
                        Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                    Else
                        ApakahLokal = Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                        If ApakahLokal = True Then
                            IndexStr = Me.TxtAOP.Text.IndexOf("L")
                            IndexStr1 = Me.TxtAOP.Text.IndexOf("I", 12)
                        Else
                            IndexStr = Me.TxtAOP.Text.IndexOf("I", 12)
                            IndexStr1 = Me.TxtAOP.Text.IndexOf("L")
                        End If
                        If IndexStr = -1 Then
                            Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                        End If
                    End If
                Else
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.Objsupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                        '         Me.txtNoSeriFakturPajak.Text = Me.Objsupplier.GetTaxCode(Me.txtPrimarySupplier.Text)
                        '         Me.txtNoNPWP.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                        Me.cboSupplierAddress.SelectedIndex = 0
                        Me.IsiCboDeliveryOrder(Me.txtPrimarySupplier.Text)
                        Me.cboSupplierAddress.Focus()

                        'ganti no transaksi
                        If Me.baru = True Then
                            Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                        Else
                            ApakahLokal = Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                            If ApakahLokal = True Then
                                IndexStr = Me.TxtAOP.Text.IndexOf("L")
                                IndexStr1 = Me.TxtAOP.Text.IndexOf("I", 12)
                            Else
                                IndexStr = Me.TxtAOP.Text.IndexOf("I", 12)
                                IndexStr1 = Me.TxtAOP.Text.IndexOf("L")
                            End If
                            If IndexStr = -1 Then
                                Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                            End If
                        End If
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.cboSupplierAddress.Items.Clear()
                        Me.txtSupplierAddIdentifier.Text = ""
                        Me.txtPrimarySupplier.Text = "-1"
                        '                        Me.txtNoNPWP.Clear()
                        '                        Me.txtNoSeriFakturPajak.Clear()
                        Me.cboDO.DataSource = Nothing
                        Me.txtSupplierCode.Focus()
                    End If
                End If
            Else
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.cboSupplierAddress.Items.Clear()
                Me.txtSupplierAddIdentifier.Text = ""
                Me.txtPrimarySupplier.Text = "-1"
                '                Me.txtNoNPWP.Clear()
                Me.cboDO.DataSource = Nothing
                '                Me.txtNoSeriFakturPajak.Clear()
            End If

            'cek supplier
            If Me.txtSupplierCode.Text <> "" Then
                If Me.Lokal = True Then
                    If Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text) = False Then
                        reset = True
                    End If
                Else
                    If Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text) = True Then
                        reset = True
                    End If
                End If

                If reset = True Then
                    If Me.Lokal = True Then
                        MsgBox("Please choose local supplier.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Please choose foreign supplier.", MsgBoxStyle.Information, "Information")
                    End If

                    Me.txtSupplierCode.Clear()
                    Me.TxtSupplierName.Clear()
                    Me.cboSupplierAddress.Items.Clear()
                    Me.txtSupplierAddIdentifier.Text = ""
                    Me.txtPrimarySupplier.Text = "-1"
                    Me.cboDO.DataSource = Nothing
                    Me.txtSupplierCode.Focus()
                End If
            End If
            If Me.txtPrimarySupplier.Text <> Me.buffer Then
                Me.TxtNPWPNo.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)

                '2025 Core Tax Tahap 2.1
                If Me.dtpDate.Value >= Me.ObjSetNPWP.Tglberlaku.Date Then
                    Me.TxtNPWPNo.Text = Me.ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Dim ApakahLokal As Boolean = False
        Dim IndexStr, IndexStr1 As Integer
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""
        Dim Reset As Boolean = False

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.txtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = i.PrimarySupplier
                Me.txtSupplierCode.Text = Me.Objsupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.cboSupplierAddress.SelectedIndex = 0
                '      Me.txtNoSeriFakturPajak.Text = Me.Objsupplier.GetTaxCode(Me.txtPrimarySupplier.Text)
                '      Me.txtNoNPWP.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                Me.IsiCboDeliveryOrder(Me.txtPrimarySupplier.Text)
                Me.cboSupplierAddress.Focus()

                'ganti no transaksi
                If Me.baru = True Then
                    Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                Else
                    ApakahLokal = Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text)
                    If ApakahLokal = True Then
                        IndexStr = Me.TxtAOP.Text.IndexOf("L")
                        IndexStr1 = Me.TxtAOP.Text.IndexOf("I", 12)
                    Else
                        IndexStr = Me.TxtAOP.Text.IndexOf("I", 12)
                        IndexStr1 = Me.TxtAOP.Text.IndexOf("L")
                    End If
                    If IndexStr = -1 Then
                        Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                    End If
                End If
            Else
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.cboSupplierAddress.Items.Clear()
                Me.txtSupplierAddIdentifier.Text = ""
                Me.txtPrimarySupplier.Text = "-1"
                Me.cboDO.DataSource = Nothing
                Me.txtSupplierCode.Focus()
            End If

            'cek supplier
            If Me.txtSupplierCode.Text <> "" Then
                If Me.Lokal = True Then
                    If Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text) = False Then
                        Reset = True
                    End If
                Else
                    If Me.Objsupplier.ApakahLokal(Me.txtPrimarySupplier.Text) = True Then
                        Reset = True
                    End If
                End If

                If Reset = True Then
                    If Me.Lokal = True Then
                        MsgBox("Please choose local supplier.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Please choose foreign supplier.", MsgBoxStyle.Information, "Information")
                    End If

                    Me.txtSupplierCode.Clear()
                    Me.TxtSupplierName.Clear()
                    Me.cboSupplierAddress.Items.Clear()
                    Me.txtSupplierAddIdentifier.Text = ""
                    Me.txtPrimarySupplier.Text = "-1"
                    Me.cboDO.DataSource = Nothing
                    Me.txtSupplierCode.Focus()
                End If
            End If
            Me.TxtNPWPNo.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)

            '2025 Core Tax Tahap 2.1
            If Me.dtpDate.Value >= Me.ObjSetNPWP.Tglberlaku.Date Then
                Me.TxtNPWPNo.Text = Me.ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub cboSupplierAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSupplierAddress.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.cboSupplierAddress.SelectedIndex <> -1 Then
                    Me.txtSupplierAddIdentifier.Text = Me.Objsupplier.GetIdentifierAdd(Me.txtPrimarySupplier.Text, _
                    Me.cboSupplierAddress.Text)
                End If
            End If
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
                If CDec(Me.txtTax.Text) > 100 Then
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

    Private Sub TxtSupplierName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierName.TextChanged
        Try
            If Me.txtPrimarySupplier.Text <> "-1" And Me.txtPrimarySupplier.Text <> "" Then
                If Me.TxtSupplierName.Text <> "" Then
                    If Me.txtSupplierCode.Text <> "" Then
                        Me.Text = "AOP: " + Me.txtSupplierCode.Text + " - " + Me.TxtSupplierName.Text
                    Else
                        Me.Text = "AOP"
                    End If
                Else
                    Me.Text = "AOP"
                End If
            Else
                Me.Text = "AOP"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
            If Me.TxtRemark.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryTransaksiAOP.Text) AndAlso CDec(Me.TxtPrimaryTransaksiAOP.Text)) Then
                If Me.txtPrimarySupplier.Text = "" Or Me.txtPrimarySupplier.Text = "-1" Then
                    If Me.TxtRemark.Enabled = True Then
                        MsgBox("Please specify Supplier first.", MsgBoxStyle.Information, "Information")
                        Me.txtSupplierCode.Focus()
                    End If
                    Exit Sub
                End If

                Dim ij As New Core_Module.FrmDetailAddressSupplier
                ij.vAddIdentifier = Me.txtSupplierAddIdentifier.Text
                ij.PrimarySup = Me.txtPrimarySupplier.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                If Me.TxtRemark.Enabled = False Then
                    ij.cboCustAddCode.Enabled = False
                End If
                ij.ShowDialog()

                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.Text = Me.Objsupplier.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiCboDeliveryOrder(ByVal vPrmarySupplier As Decimal) As Boolean
        Dim row2 As DseTransaksiAOP.DataTable9Row
        Dim row1 As DseTransaksiGRN.DataTable8Row
        Dim PrimaryMinGRN As Double = Double.MaxValue
        Dim Filter As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data yang akan digunakan sebagai acuan
            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()
            Me.dgvItem.Rows.Clear()

            'ambil delivery order yang belum ada invoice
            Me.objGRN.IsiTBLGRNBelumAdaAOP(Me.txtPrimarySupplier.Text, Me.TBL)
            'tambahkan dengan delivery order yang sudah disimpan dalam object lama
            Me.ObjAOPItem.IsiTBLGRN(Me.TxtPrimaryTransaksiAOP.Text, Me.TBL2)
            While Me.TBL2.Rows.Count > 0
                row2 = Me.TBL2.Rows(0)
                Me.TBL.Rows.Add(row2.Primary_GRN, row2.GRN_No)
                Me.TBL2.Rows.Remove(row2)
            End While
            'kurangkan dengan yang ada didalam dgv
            Me.IsiTBLDO(Me.TBL2)
            Dim Ketemu As Boolean = False
            While Me.TBL2.Rows.Count > 0
                Ketemu = False
                row2 = Me.TBL2.Rows(0)
                For Each row1 In Me.TBL.Rows
                    If row1.PrimaryGRN = row2.Primary_GRN And row1.GRN_No = row2.GRN_No Then
                        Me.TBL.Rows.Remove(row1)
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    'tidak mungkin
                    MsgBox("Logic error, contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Exit Function
                End If
                Me.TBL2.Rows.Remove(row2)
            End While

            If Me.cboFillBy.SelectedIndex = 0 Then
                'by supplier do number
                For Each row1 In Me.TBL.Rows
                    If row1.PrimaryGRN < PrimaryMinGRN Then
                        PrimaryMinGRN = row1.PrimaryGRN
                    End If
                    row1.Supplier_DO_Number = ""
                Next

                If PrimaryMinGRN <> Double.MaxValue Then
                    Me.DTAGRN.Fill(Me.TBLGRN, Me.txtPrimarySupplier.Text, PrimaryMinGRN)
                    For Each row1 In Me.TBL.Rows
                        Filter = "PrimaryGRN = " + row1.PrimaryGRN.ToString
                        Me.TBLGRN.DefaultView.RowFilter = Filter

                        If Me.TBLGRN.DefaultView.Count <> 0 Then
                            Me.rowGRN = Me.TBLGRN.DefaultView(0).Row
                            If Me.rowGRN._Supplier_s_DO_Number <> "" Then
                                row1.Supplier_DO_Number = Me.rowGRN._Supplier_s_DO_Number
                            Else
                                row1.Supplier_DO_Number = row1.GRN_No
                            End If
                        End If
                    Next
                End If

                Me.cboDO.DataSource = Me.TBL.DefaultView
                Me.TBL.DefaultView.Sort = "[Supplier DO Number] ASC"
                Me.cboDO.DisplayMember = "Supplier DO Number"
            Else
                'by grn number
                Me.cboDO.DataSource = Me.TBL.DefaultView
                Me.TBL.DefaultView.Sort = "[Stock In No] ASC"
                Me.cboDO.DisplayMember = "Stock In No"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function IsiCboDeliveryOrder2(ByVal vPrimarySupplier As Decimal) As Boolean
        Dim row2 As DseTransaksiAOP.DataTable9Row
        Dim row1 As DseTransaksiGRN.DataTable8Row
        Dim Filter As String = ""
        Dim PrimaryMinGRN As Double = -1

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data yang akan digunakan sebagai acuan
            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()

            'ambil delivery order yang belum ada invoice
            Me.objGRN.IsiTBLGRNBelumAdaAOP(Me.txtPrimarySupplier.Text, Me.TBL)
            'tambahkan dengan delivery order yang sudah disimpan dalam object lama
            Me.ObjAOPItem.IsiTBLGRN(CDec(Me.TxtPrimaryTransaksiAOP.Text), Me.TBL2)
            While Me.TBL2.Rows.Count > 0
                row2 = Me.TBL2.Rows(0)

                row1 = TBL.NewDataTable8Row
                row1.PrimaryGRN = row2.Primary_GRN
                row1.GRN_No = row2.GRN_No
                row1.Supplier_DO_Number = row2._Supplier_s_DO_Number
                row1.Stock_In_No = row2.Stock_In_No
                Me.TBL.Rows.Add(row1)

                Me.TBL2.Rows.Remove(row2)
            End While

            'kurangkan dengan yang ada didalam dgv
            Me.IsiTBLDO(Me.TBL2)
            Dim Ketemu As Boolean = False
            While Me.TBL2.Rows.Count > 0
                Ketemu = False
                row2 = Me.TBL2.Rows(0)
                For Each row1 In Me.TBL.Rows
                    If row1.PrimaryGRN = row2.Primary_GRN And row1.GRN_No = row2.GRN_No Then
                        Me.TBL.Rows.Remove(row1)
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    'tidak mungkin
                    MsgBox("Logic error, contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Exit Function
                End If
                Me.TBL2.Rows.Remove(row2)
            End While

            For Each row1 In Me.TBL.Rows
                If row1.Supplier_DO_Number.Trim = "" Then
                    row1.Supplier_DO_Number = row1.Stock_In_No
                End If
            Next
            If Me.cboFillBy.SelectedIndex = 0 Then
                'by supplier do number
                'For Each row1 In Me.TBL.Rows
                '    If row1.PrimaryGRN < PrimaryMinGRN Then
                '        PrimaryMinGRN = row1.PrimaryGRN
                '    End If
                '    row1.Supplier_DO_Number = ""
                'Next

                'If PrimaryMinGRN <> Double.MaxValue Then
                '    Me.DTAGRN.Fill(Me.TBLGRN, Me.txtPrimarySupplier.Text, PrimaryMinGRN)
                '    For Each row1 In Me.TBL.Rows
                '        Filter = "PrimaryGRN = " + row1.PrimaryGRN.ToString
                '        Me.TBLGRN.DefaultView.RowFilter = Filter

                '        If Me.TBLGRN.DefaultView.Count <> 0 Then
                '            Me.rowGRN = Me.TBLGRN.DefaultView(0).Row
                '            If Me.rowGRN._Supplier_s_DO_Number <> "" Then
                '                row1.Supplier_DO_Number = Me.rowGRN._Supplier_s_DO_Number
                '            Else
                '                row1.Supplier_DO_Number = row1.GRN_No
                '            End If
                '        End If
                '    Next
                'End If

                Me.cboDO.DataSource = Me.TBL.DefaultView
                Me.TBL.DefaultView.Sort = "[Supplier DO Number] ASC"
                Me.cboDO.DisplayMember = "Supplier DO Number"
            Else
                'by grn number
                Me.cboDO.DataSource = Me.TBL.DefaultView
                Me.TBL.DefaultView.Sort = "[Stock In No] ASC"
                Me.cboDO.DisplayMember = "Stock In No"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function


    Private Sub IsiTBLDO(ByVal TBL As DseTransaksiAOP.DataTable9DataTable)
        Dim row As DataGridViewRow
        Dim PrimaryDummy As Decimal = -1
        Dim PrimaryStockIn As Double = -1

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryGRN.Index).Value <> PrimaryDummy Then
                    PrimaryDummy = row.Cells(Me.ColPrimaryGRN.Index).Value
                    PrimaryStockIn = Me.objGRN.GetPrimaryStockIn(row.Cells(Me.ColPrimaryGRN.Index).Value)
                    TBL.Rows.Add(row.Cells(Me.ColPrimaryGRN.Index).Value, _
                    Me.objGRN.GetNoGRN(row.Cells(Me.ColPrimaryGRN.Index).Value), _
                    Me.ObjStockIn.GetSupplierDONumber(PrimaryStockIn), _
                    Me.ObjStockIn.GetNoTransaksi(PrimaryStockIn))
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub InputProcedure()
        If Me.cboDO.Items.Count > 0 Then
            Dim TBL1 As New DseTransaksiGRN.tbltrans_grnitemview2_MultiDataTable
            Dim DTA1 As New DseTransaksiGRNTableAdapters.tbltrans_grnitemview2_MultiTableAdapter
            Dim rowItemGRN As DseTransaksiGRN.tbltrans_grnitemview2_MultiRow
            Dim rowGRNNo As DseTransaksiGRN.DataTable8Row
            Dim Dataview As DataRowView
            'Dim NetPriceExclude As Decimal = 0
            Dim Amount As Decimal = 0
            Dim PrimaryPO, PrimaryGRN, PrimarySupplier, PrimaryCur, DefaultPrec As Decimal
            Dim PONumber, StockInNumber, CurCode As String
            Dim PrimaryCurrency As Double = -1
            Dim RowDGV As DataGridViewRow

            Try
                If Me.txtPrimaryCurrency.Text = "-1" Or Me.txtPrimaryCurrency.Text = "" Then
                    MsgBox("Please specify the currency for the approval of purchase.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCurrencyCode.Focus()
                    Exit Sub
                End If

                Dataview = Me.cboDO.SelectedItem
                rowGRNNo = Dataview.Row
                'cek tanggal GRN harus berada pada bulan yang sama dengan tanggal AOP
                If Format(Me.objGRN.GetTanggalTransaksi(rowGRNNo.PrimaryGRN), "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                    MsgBox("You can only enter stock in purchase from " + Format(Me.dtpDate.Value.Date, "MMM yyyy"), MsgBoxStyle.Information, _
                    "Information")
                    Exit Sub
                End If

                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                'isikan tabel acuan
                DTA1.FillBy_PrimaryGRN(TBL1, rowGRNNo.PrimaryGRN)
                Me.Isikan_UnitPrice_Discount_NettPrice(TBL1)
                'tambahkan batasan supaya currency sama pada setiap item grn
                If TBL1.Rows.Count > 0 And Me.dgvItem.Rows.Count > 0 Then
                    rowItemGRN = TBL1.Rows(0)
                    RowDGV = Me.dgvItem.Rows(0)
                    PrimaryCurrency = RowDGV.Cells(Me.ColPrimaryCurrency.Index).Value
                    If rowItemGRN.Primary_Currency <> PrimaryCurrency Then
                        MsgBox("You can only enter stock in with po using currency :" + _
                        Me.ObjCurrency.GetCurrencyCode(PrimaryCurrency) + ".", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If

                'set currency code jika, masih kosong dgvnya
                If Me.dgvItem.RowCount = 0 Then
                    rowItemGRN = TBL1.Rows(0)
                    Me.txtPrimaryCurrency.Text = Me.ObjPO.GetPrimaryCurrency(rowItemGRN.Primary_PO)
                    Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetTaxRate()
                    Me.SetDefaultPrecision()
                End If

                'untuk setiap baris dalam tabel
                PrimaryPO = -1
                PrimaryGRN = -1
                PrimarySupplier = -1
                PrimaryCur = -1
                DefaultPrec = 0
                PONumber = ""
                StockInNumber = ""
                CurCode = ""
                For Each rowItemGRN In TBL1.Rows
                    'hitung ulang exclude price jika perlu
                    With rowItemGRN
                        If .Primary_PO <> PrimaryPO Then
                            PrimaryPO = .Primary_PO
                            PONumber = Me.ObjPO.GetNoTransaksiPO(.Primary_PO)
                            PrimarySupplier = Me.ObjPO.GetPrimarySupplier(.Primary_PO)
                            PrimaryCur = Me.ObjPO.GetPrimaryCurrency(.Primary_PO)
                            CurCode = Me.ObjCurrency.GetCurrencyCode(PrimaryCur)
                            DefaultPrec = Me.ObjCurrency.GetDefaultPrecission(PrimaryCur)
                        End If
                        If .PrimaryGRN <> PrimaryGRN Then
                            PrimaryGRN = .PrimaryGRN
                            StockInNumber = Me.ObjStockIn.GetNoTransaksi(Me.objGRN.GetPrimaryStockIn(.PrimaryGRN))
                        End If
                        .No_PO = PONumber
                        .Primary_Supplier = PrimarySupplier
                        .Currency_Code = CurCode
                        .Default_Precision = DefaultPrec
                        .Primary_Currency = PrimaryCur
                        .No_StockIn = StockInNumber

                        'NetPriceExclude = Me.ObjCD.SetDecimal(.Unit_Price * (100 - .Discount) / 100, .Default_Precision)

                        ''@rosadinnn
                        ''@MUaT
                        If .Secondary_Qty <= 0 Then
                            If .price_by_multi_unit = 1 Then
                                Amount = .Nett_Price * .Total_Qty_Multi
                            Else
                                Amount = .Nett_Price * .TotalQty
                            End If
                        Else
                            Amount = .Nett_Price * .Total_Qty_Multi * .Secondary_Qty
                        End If

                        'tambahkan ke dalam dgv
                        Dim AtPrice As Decimal = 0

                        ''@rosadinnn
                        ''@MUaT
                        Me.dgvItem.Rows.Add()
                        RowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        RowDGV.Cells(Me.ColItem_No.Index).Value = ""
                        RowDGV.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        RowDGV.Cells(Me.ColDescription.Index).Value = .Description
                        RowDGV.Cells(Me.ColSuppliersCodeNumber.Index).Value = ""
                        RowDGV.Cells(Me.ColSuppliersProductDesc.Index).Value = ""
                        RowDGV.Cells(Me.ColPONO.Index).Value = .No_PO
                        RowDGV.Cells(Me.ColGRNNO.Index).Value = .GRN_No
                        RowDGV.Cells(Me.ColStockOutNo.Index).Value = .No_StockIn
                        RowDGV.Cells(Me.ColSuppDONo.Index).Value = ._Supplier_s_DO_Number
                        RowDGV.Cells(Me.ColQtyMaster.Index).Value = .Qty_Master
                        RowDGV.Cells(Me.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        RowDGV.Cells(Me.ColMasterUnits.Index).Value = ""
                        RowDGV.Cells(Me.ColQty.Index).Value = .Qty
                        RowDGV.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(.TotalQty, .Decimal_Product)
                        RowDGV.Cells(Me.ColUnit.Index).Value = .Units
                        RowDGV.Cells(Me.ColBatchNo.Index).Value = .Batch_No
                        RowDGV.Cells(Me.ColSecQty.Index).Value = .Secondary_Qty
                        RowDGV.Cells(Me.ColSecUnit.Index).Value = .SecUnit
                        RowDGV.Cells(Me.ColCurrency.Index).Value = .Currency_Code
                        RowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(.Unit_Price, .Default_Precision)
                        RowDGV.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(.Discount, 2)
                        RowDGV.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(.DiscAmount, .Default_Precision)
                        RowDGV.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(.Nett_Price, .Default_Precision)
                        RowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, .Default_Precision)
                        RowDGV.Cells(Me.ColATCost.Index).Value = AtPrice
                        RowDGV.Cells(Me.ColShowDetail.Index).Value = "Show"
                        RowDGV.Cells(Me.ColRemarks.Index).Value = .Remarks
                        RowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDGV.Cells(Me.ColPrimaryPO.Index).Value = .Primary_PO
                        RowDGV.Cells(Me.ColPrimaryGRN.Index).Value = .PrimaryGRN
                        RowDGV.Cells(Me.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Units
                        RowDGV.Cells(Me.ColPrimaryUnit.Index).Value = .Primary_Units
                        RowDGV.Cells(Me.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        RowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        RowDGV.Cells(Me.Col_AtAmount.Index).Value = .DiscFromPerc
                        RowDGV.Cells(Me.ColDiscFromPerct.Index).Value = .DiscFromPerc
                        RowDGV.Cells(Me.ColQtyMulti_Item.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi))
                        RowDGV.Cells(Me.ColUnitsCodeMulti_Item.Index).Value = .Unit_Code_Multi
                        RowDGV.Cells(Me.ColRatioReference_Item.Index).Value = .Ratio_Reference
                        RowDGV.Cells(Me.ColRatioDefault_Item.Index).Value = .Ratio_Default
                        RowDGV.Cells(Me.ColPrimaryUnitsMulti_Item.Index).Value = .Primary_Unit_Multi
                        RowDGV.Cells(Me.ColItem_Adjustment.Index).Value = 0
                        RowDGV.Cells(Me.ColItem_AdjustmentReason.Index).Value = ""
                        RowDGV.Cells(Me.ColItem_AdjustUser.Index).Value = ""
                        RowDGV.Cells(Me.ColItemAdjustmentTimestamp.Index).Value = "-"

                        If .Primary_Units <> .Primary_Unit_Multi Then
                            RowDGV.Cells(Me.ColRatioDesc_Item.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, _
                                    ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi)) _
                                    & " " & .Unit_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, _
                                    ObjProduct.GetAngkaDecimal(.Primary_Product)) & " " & .Units
                        Else
                            RowDGV.Cells(Me.ColRatioDesc_Item.Index).Value = "-"
                        End If

                        RowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            RowDGV.Cells(Me.ColPriceByMultiUnit_str.Index).Value = "/" + .Unit_Code_Multi
                        Else
                            RowDGV.Cells(Me.ColPriceByMultiUnit_str.Index).Value = "/" + .Units
                        End If
                    End With
                Next

                'tambahkan exchange rate data
                Me.SetCalculateBy()

                'penomoran
                Me.ObjCD.SetNumber(Me.dgvItem)
                Me.ObjCD.SetNumber(Me.dgvExchageRates)
                Me.FormatDGV()
                'hilangkan do dari item
                Me.TBL.Rows.Remove(rowGRNNo)
                'tidak bisa ganti kode supplier
                Me.btnFindSupplier.Enabled = False
                Me.txtSupplierCode.BackColor = Me.TxtAOP.BackColor
                Me.txtSupplierCode.ReadOnly = True
                'tampilan exchange rate
                Me.SetDGVExchangeRate()
                'hitung ulang total
                If Me.dgvItem.RowCount > 0 Then
                    Me.IsiDGVReturn()
                    Me.Add_DP_Procedure()
                End If
                Me.Hitung_Total_Multi()

                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                DTA1.Dispose()
                TBL1.Dispose()
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End Try
        End If
    End Sub

    Private Sub BtnInput_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Try
            Me.InputProcedure()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    ''@rosadinnn
    ''@MUaT
    Private Sub Isikan_UnitPrice_Discount_NettPrice(ByVal vTBL As DseTransaksiGRN.tbltrans_grnitemview2_MultiDataTable)
        Dim rowPO As DseTransaksiGRN.tbltrans_po_item_view5Row
        Dim rowGRN As DseTransaksiGRN.tbltrans_grnitemview2_MultiRow
        Dim ListPrimaryPO As New List(Of Double)
        Dim Filter As String = ""
        Dim PrimaryPO As Decimal

        Try
            'kumpulkan primary po
            For Each rowGRN In vTBL.Rows
                If ListPrimaryPO.Contains(rowGRN.Primary_PO) = False Then
                    ListPrimaryPO.Add(rowGRN.Primary_PO)
                End If
            Next
            For Each PrimaryPO In ListPrimaryPO
                Me.DTAPO.FillBy_PrimaryPO(Me.TBLPO, PrimaryPO)
                For Each rowGRN In vTBL.Rows
                    If rowGRN.Primary_PO = PrimaryPO Then
                        Filter = "[Primary Transaksi] =" + rowGRN.Primary_PO.ToString + _
                        " AND [Primary Product] = " + rowGRN.Primary_Product.ToString
                        TBLPO.DefaultView.RowFilter = Filter
                        If TBLPO.DefaultView.Count > 0 Then
                            rowPO = TBLPO.DefaultView(0).Row
                            rowGRN.Unit_Price = rowPO.Unit_Price
                            rowGRN.Discount = rowPO.Discount
                            rowGRN.Nett_Price = rowPO.Nett_Price
                            rowGRN.Amount = rowGRN.Nett_Price * rowGRN.Total_Qty_Multi
                            rowGRN.Primary_Currency = rowPO.Primary_Currency
                            rowGRN.DiscAmount = rowPO.Disc_Amount
                            rowGRN.DiscFromPerc = rowPO.Disc_From_Perct
                            rowGRN.price_by_multi_unit = rowPO.price_by_multi_unit
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryPO = Nothing
        End Try
    End Sub

    Public Sub SetDGVExchangeRate(Optional ByVal vHitungTotal As Boolean = False)
        Dim rowDGV As DataGridViewRow
        Dim PrimaryCurrency1, PrimaryCurrency2 As Decimal
        Dim GaSama As Boolean = False

        Try
            If Me.dgvItem.RowCount > 0 Then
                PrimaryCurrency1 = Me.txtPrimaryCurrency.Text
                rowDGV = Me.dgvItem.Rows(0)
                PrimaryCurrency2 = rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value
                For Each rowDGV In Me.dgvItem.Rows
                    If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> PrimaryCurrency1 Then
                        GaSama = True
                    End If
                Next

                If GaSama = False Then
                    If Me.Lokal = True Then
                        Me.SplitContainer1.Panel1Collapsed = True
                    Else
                        Me.SplitContainer2.Panel2Collapsed = True
                    End If
                    For Each rowDGV In Me.dgvExchageRates.Rows
                        rowDGV.Cells(Me.ColExchangeRate.Index).Value = 1
                    Next
                    If vHitungTotal = True Then

                        ''@rosadinnn
                        ''@MUaT
                        Me.Hitung_Total_Multi()

                    End If
                Else
                    If Me.Lokal = True Then
                        Me.SplitContainer1.Panel1Collapsed = False
                    Else
                        Me.SplitContainer2.Panel2Collapsed = False
                    End If
                End If

                If Me.SplitContainer2.Panel2Collapsed = False Then
                    For Each rowDGV In Me.dgvCost.Rows
                        If rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value = PrimaryCurrency1 Then
                            'rowDGV.Cells(Me.ColExcRateCost.Index).Value = 1
                            'Me.Hitung_Cost_Amount(rowDGV)
                        End If
                    Next

                    'hitung ulang AT-cost

                    ''@rosadinnn
                    ''@MUaT
                    Me.Hitung_AtCost_Multi()

                End If
            End If

            If Me.txtCurrencyCode.Text = "" Then
                Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate"
                'Me.dgvCost.Columns(Me.ColExcRateCost.Index).HeaderText = "Exchange Rate"
            Else
                Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate to " _
                + Me.txtCurrencyCode.Text
                'Me.dgvCost.Columns(Me.ColExcRateCost.Index).HeaderText = "Exchange Rate to " _
                '+ Me.txtCurrencyCode.Text
            End If

            Me.SplitContainer1.Panel1Collapsed = False
            Me.SplitContainer1.Panel2Collapsed = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowCustCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowCustCode.CheckedChanged
        Try
            If Me.dgvItem.Columns.Count > 0 Then
                If Me.chkShowCustCode.Checked = True Then
                    Me.dgvItem.Columns(3).Visible = True
                    Me.dgvItem.Columns(4).Visible = True
                Else
                    Me.dgvItem.Columns(3).Visible = False
                    Me.dgvItem.Columns(4).Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Me.buffer = e.Row.Cells(Me.ColPrimaryGRN.Index).Value
    End Sub

    Private Sub SetCalculateBy()
        Try
            If Me.TxtRemark.Enabled = True Then
                'reset data item exchange rate
                Me.dgvExchageRates.Rows.Clear()
                If Me.cboCalculateBy.SelectedIndex = 0 Then
                    'hitung berdasarkan sales invoice
                    If Me.dgvItem.RowCount = 0 Then
                        'kalo ga ada anak, maka exc
                        Me.dgvExchageRates.Rows.Add("0", Me.TxtAOP.Text, Me.txtCurrencyCode.Text, _
                        1, "", Me.TxtPrimaryTransaksiAOP.Text, Me.txtPrimaryCurrency.Text)
                    Else
                        Dim rowDGV As DataGridViewRow
                        rowDGV = Me.dgvItem.Rows(0)
                        Me.dgvExchageRates.Rows.Add("0", Me.TxtAOP.Text, rowDGV.Cells(Me.ColCurrency.Index).Value, _
                        1, "", Me.TxtPrimaryTransaksiAOP.Text, rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                    End If
                ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
                    'hitung berdasarkan delivery order
                    'hitung berdasarkan sales order
                    Dim Ketemu As Boolean = False
                    Dim rowER, rowItem As DataGridViewRow

                    For Each rowItem In Me.dgvItem.Rows
                        Ketemu = False
                        For Each rowER In Me.dgvExchageRates.Rows
                            If rowER.Cells(ColPrimaryDokumen.Index).Value = rowItem.Cells(Me.ColPrimaryGRN.Index).Value Then
                                Ketemu = True
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan exchange rate baru 
                            Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(Me.ColGRNNO.Index).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
                            1, "", rowItem.Cells(Me.ColPrimaryGRN.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
                        End If
                    Next
                ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
                    'hitung berdasarkan sales order
                    Dim Ketemu As Boolean = False
                    Dim rowER, rowItem As DataGridViewRow

                    For Each rowItem In Me.dgvItem.Rows
                        Ketemu = False
                        For Each rowER In Me.dgvExchageRates.Rows
                            If rowER.Cells(ColPrimaryDokumen.Index).Value = rowItem.Cells(Me.ColPrimaryPO.Index).Value Then
                                Ketemu = True
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan exchange rate baru 
                            Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(Me.ColPONO.Index).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
                            1, "", rowItem.Cells(Me.ColPrimaryPO.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
                        End If
                    Next
                End If
                'tampilan exchange rate
                Me.SetDGVExchangeRate()
                'hitung ulang
                Me.Hitung_Total_Multi()
                'set number
                Me.ObjCD.SetNumber(Me.dgvExchageRates)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCalculateBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCalculateBy.SelectedIndexChanged
        Try
            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutyStamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDutyStamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutyStamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDutyStamp.Enabled = True Then
                    If Me.TxtDutyStamp.ReadOnly = False Then
                        Me.TxtDutyStamp.BackColor = Me.BackColorAll
                    Else
                        Me.TxtDutyStamp.BackColor = Me.txtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDutyStamp.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtDutyStamp, sender, e)
    End Sub

    Private Sub txtDutyStamp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDutyStamp.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtDutyStamp, sender, e)
    End Sub

    Private Sub txtTaxExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Enter
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = False And Me.txtTaxExchangeRate.Enabled = True Then
                    Me.ObjCD.Number_Enter(Me.txtTaxExchangeRate, sender, e)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = True Then
                    txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtTaxExchangeRate.BackColor = BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtTaxExchangeRate, sender, e)
    End Sub

    Private Sub txtTaxExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtTaxExchangeRate, sender, e)
    End Sub

    Private Sub dgvExchageRates_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvExchageRates.CellBeginEdit
        Try
            'yang bisa diedit hanya:
            '1. exchange rate
            '2. remarks
            If e.ColumnIndex = 3 Or e.ColumnIndex = 4 Then
                Me.buffer = Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExchageRates_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvExchageRates.CellEndEdit
        Try
            'pastikan diisi
            If IsNothing(Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = 3 Then
                    'pastikan nilainya positif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit grouping
                    Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCD.DigitGrouping(Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value)
                    'hitung kembali total invoice
                    Me.Hitung_Total_Multi()
                    'Me.Hitung_AtCost()
                End If
            Else
                If e.ColumnIndex = 4 Then
                    Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = ""
                Else
                    Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Add_DP_Procedure()
        Dim ij As New List(Of Double)
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.Rows.Count <> 0 Then
                If Me.txtCurrencyCode.Text = "" Then
                    MsgBox("Please specify currency code.", MsgBoxStyle.Information, "Information")
                    Me.txtCurrencyCode.Focus()
                    Exit Sub
                End If


                'tambahkan item DP
                Me.DGVdp.Rows.Clear()
                For Each row In Me.dgvItem.Rows
                    If ij.Contains(Me.dgvItem.Item(Me.ColPrimaryPO.Index, row.Index).Value) = False Then
                        ij.Add(Me.dgvItem.Item(Me.ColPrimaryPO.Index, row.Index).Value)
                    End If
                Next
                Me.ObjAOP.IsiDgvDP(ij, Me.DGVdp, Me.txtCurrencyCode.Text, Me.ObjCD, Me.TxtPrimaryTransaksiAOP.Text)
                Me.ObjCD.SetNumber(Me.DGVdp)
                Me.FormatDGV()
                Me.Hitung_Total_Multi()
            Else
                MsgBox("Please insert AOP item first", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddDP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddDP.LinkClicked
        Try
            Add_DP_Procedure()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVdp.CellBeginEdit
        Try
            'yang bisa diedit hanya remarks
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex <> 4 Then
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDGVReturn()
        Dim ListPrimary As New List(Of Double)
        Dim row As DataGridViewRow
        Dim rowpR As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_view1Row
        Dim dummy As String = ""

        Try
            'kosongkan baris yang diberikan
            Me.DGVitemPR.Rows.Clear()
            Me.dgvPR.Rows.Clear()
            Me.TBLPR.Rows.Clear()

            'cari data SR, berdasarkan data do yang dituliskan
            For Each row In Me.dgvItem.Rows
                If ListPrimary.Contains(row.Cells(Me.ColPrimaryGRN.Index).Value) = False Then
                    ListPrimary.Add(row.Cells(Me.ColPrimaryGRN.Index).Value)
                End If
            Next
            If IsNumeric(Me.TxtPrimaryTransaksiAOP.Text) = False Then
                Me.TxtPrimaryTransaksiAOP.Text = "-1"
            End If
            Me.ObjPR.IsiTBLPRByPrimaryGRN(ListPrimary, TBLPR, Me.TxtPrimaryTransaksiAOP.Text)
            For Each rowpR In Me.TBLPR.Rows
                'cek, hanya purchase return yang sudah stock out yang boleh dimasukkan ke dalam AOP
                If Me.ObjPR.GetNoTransaksiStockOut(rowpR.PrimaryPR) <> "" Then
                    Me.DGVitemPR.Rows.Add("", rowpR.PR_No, rowpR.GRN_No, "", rowpR.PrimaryPR)
                Else
                    If dummy = "" Then
                        dummy = rowpR.PR_No
                    Else
                        dummy = dummy + ", " + rowpR.PR_No
                    End If
                End If
            Next

            If dummy <> "" Then
                MsgBox("Please check purchase return: " + dummy + _
                ". The purchase return don't have any stock out refering to them. Please create stock out purchase return.", _
                MsgBoxStyle.Information, "Information")
            End If

            'set number untuk dgv
            Me.ObjCD.SetNumber(Me.DGVitemPR)
            'isikan item return 
            Me.IsiDGVReturnItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Function ApakahAdaDalamDGVReturn(ByVal vPR_Number As String) As Boolean
        Dim rowdgv As DataGridViewRow

        Try
            For Each rowdgv In Me.DGVitemPR.Rows
                If rowdgv.Cells(CopPRNo.Index).Value = vPR_Number Then
                    Return True
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiDGVReturnItem(Optional ByVal vHitungTotal As Boolean = True)
        Dim ListPrimary As New List(Of Double)
        Dim row As DataGridViewRow

        Try
            'untuk setiap data PR yang ada, masukkan ke dalam dgv item
            For Each row In Me.DGVitemPR.Rows
                If ListPrimary.Contains(row.Cells(Me.ColPrimaryPR.Index).Value) = False Then
                    ListPrimary.Add(row.Cells(Me.ColPrimaryPR.Index).Value)
                End If
            Next
            Me.ObjPRItem.IsiDGVByPrimaryPR_Multi(ListPrimary, Me.dgvPR)
            Me.FormatDGVPR()

            'hitung ulang sub total dengan memperhitungkan SR
            For Each row In Me.dgvItem.Rows
                Me.Hitung_Amount1_Multi(row)
            Next
            If vHitungTotal = True Then
                Me.Hitung_Total_Multi()
            End If
            Me.ObjCD.SetNumber(Me.DGVitemPR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub llbAddSR_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddSR.LinkClicked
        Try
            If Me.dgvItem.RowCount > 0 Then
                Me.IsiDGVReturn()
            Else
                MsgBox("Please insert AOP item first", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGVPR()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvPR.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiryDate_PR.Index).Value) = False Then
                    row.Cells(Me.ColExpiryDate_PR.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiryDate_PR.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiryDate_PR.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiryDate_PR.Index).Value = _
                        Format(CDate(row.Cells(Me.ColExpiryDate_PR.Index).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(Me.ColSecQty_PR.Index).Value) = True Then
                    If row.Cells(Me.ColSecQty_PR.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.ColSecQty_PR.Index).Value = "-1" Then
                    row.Cells(Me.ColSecQty_PR.Index).Value = "-"
                    row.Cells(Me.ColSecUnit_PR.Index).Value = "-"
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatchNo_PR.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvPR.Columns(Me.ColSecQty_PR.Index).Visible = False
                Me.dgvPR.Columns(Me.ColSecUnit_PR.Index).Visible = False
            Else
                Me.dgvPR.Columns(Me.ColSecQty_PR.Index).Visible = True
                Me.dgvPR.Columns(Me.ColSecUnit_PR.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvPR.Columns(Me.ColBatchNo_PR.Index).Visible = False
            Else
                Me.dgvPR.Columns(Me.ColBatchNo_PR.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvPR.Columns(Me.ColExpiryDate_PR.Index).Visible = False
            Else
                Me.dgvPR.Columns(Me.ColExpiryDate_PR.Index).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvItemSR_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvPR.UserDeletedRow
        Try
            Me.IsiDGVReturnItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Dim DeffPrec As Decimal = 0
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, DeffPrec)
                Me.Hitung_Total_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If CDec(Me.txtTaxExchangeRate.Text) = 0 Then
                    Me.txtTaxExchangeRate.Text = 1
                End If
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.txtTaxExchangeRate.Text, 2)
                Me.Hitung_Total_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = True Then
                    txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtTaxExchangeRate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCost.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. currency
            '2. cost nominal
            '3. exchange rate
            '4. remarks
            '5. description
            '6. Account --> hanya jika sudah lewat convertion date
            If e.ColumnIndex = Me.ColCurrencyCost.Index Or _
            e.ColumnIndex = Me.ColRemarksCost.Index Or _
            e.ColumnIndex = Me.ColDescCost.Index Or e.ColumnIndex = Me.ColCostNominal.Index Then
                Me.buffer = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColExcRateCost.Index Then
                If Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value = 1 Then
                    'jika sudah IDR, maka exchange rate untuk cost tidak bisa diubah lagi
                    e.Cancel = True
                End If
                If e.ColumnIndex = Me.ColCostNominal.Index Or e.ColumnIndex = Me.ColExcRateCost.Index Then
                    'jika kolom yg diedit adalah exchange rate dan cost nominal maka buang decimal
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = _
                    Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            ElseIf e.ColumnIndex = Me.ColAccounrCode.Index Then
                If Me.ObjAOP.GetConvertionDate.Date <= Me.dtpDate.Value.Date Then
                    Me.buffer = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    'yang dibuat sebelum convertion date, tidak boleh masukkin kode akun
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCost.CellEndEdit
        Dim DefPrec As Integer = 0
        Dim PrimaryCurrencyAkun As Double = -1

        Try
            If IsNothing(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                'jika tidak ada perubahan maka tidak perlu lakukan pengecekan
                If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                    Exit Sub
                End If

                If e.ColumnIndex = Me.ColCurrencyCost.Index Then
                    'end edit currency
                    If Me.ObjCurrency.ApakahSudahAda(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Dim i As New FrmCariMasterCurrency
                        If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                            i.txtCari.Text = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
                        End If

                        i.ShowDialog()
                        If i.PrimaryCurrency <> -1 Then
                            Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.ObjCurrency.GetCurrencyCode(i.PrimaryCurrency)).ToUpper
                            Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value = i.PrimaryCurrency
                            'isikan exchange rate
                            Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjCurrency.GetDefaultExhangeRate(Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value), _
                            Core_Module.DefPrecExchangeRate)
                            If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvCost.Item(Me.ColAccounrCode.Index, e.RowIndex).Value) = True Then
                                If Me.ObjAkun.GetTipeAkun2(Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value) = ClsMasterAkun_Purchase.EnumTipeAkun2.Account_Payable Then
                                    PrimaryCurrencyAkun = Me.ObjAkun.GetPrimaryCurrency(Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value)
                                    If PrimaryCurrencyAkun <> Core_Module.PrimaryBaseCurrency Then
                                        Dim DummyTax As Decimal = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, _
                                        PrimaryCurrencyAkun)
                                        If DummyTax <> 0 Then
                                            Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(DummyTax, Core_Module.DefPrecExchangeRate)
                                        End If
                                    End If
                                End If
                            End If
                            'format ulang cost
                            DefPrec = Me.ObjCurrency.GetDefaultPrecission(Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value)
                            Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value, _
                            DefPrec)
                            Me.Hitung_Cost_Amount(Me.dgvCost.Rows(e.RowIndex))
                            Me.Hitung_AtCost_Multi()
                        Else
                            Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        i = Nothing
                    Else
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                        Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value = Me.ObjCurrency.GetPrimary(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                        'isikan exchange rate
                        Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjCurrency.GetDefaultExhangeRate(Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value), _
                        Core_Module.DefPrecExchangeRate)
                        If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvCost.Item(Me.ColAccounrCode.Index, e.RowIndex).Value) = True Then
                            If Me.ObjAkun.GetTipeAkun2(Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value) = ClsMasterAkun_Purchase.EnumTipeAkun2.Account_Payable Then
                                PrimaryCurrencyAkun = Me.ObjAkun.GetPrimaryCurrency(Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value)
                                If PrimaryCurrencyAkun <> Core_Module.PrimaryBaseCurrency Then
                                    Dim DummyTax As Decimal = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, _
                                    PrimaryCurrencyAkun)
                                    If DummyTax <> 0 Then
                                        Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(DummyTax, Core_Module.DefPrecExchangeRate)
                                    End If
                                End If
                            End If
                        End If
                        'format ulang cost
                        DefPrec = Me.ObjCurrency.GetDefaultPrecission(Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value)
                        Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvCost.Item(Me.ColCostNominal.Index, e.RowIndex).Value, _
                        DefPrec)
                        Me.Hitung_Cost_Amount(Me.dgvCost.Rows(e.RowIndex))
                        Me.Hitung_AtCost_Multi()
                    End If
                ElseIf e.ColumnIndex = Me.ColDescCost.Index Then
                    'deskripsi
                    'tidak boleh kosong
                    If IsNothing(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = True Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString = "" Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColExcRateCost.Index Then
                    'perhitungan cost amount
                    'harus positif nilainnya
                    If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekBilanganPositif(Me.dgvCost.Item(e.ColumnIndex, _
                    e.RowIndex).Value) = False Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    Else
                        'format tampilan
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value, Core_Module.DefPrecExchangeRate)
                        'hitung ulang AT-cost
                        Me.Hitung_Cost_Amount(Me.dgvCost.Rows(e.RowIndex))
                        Me.Hitung_AtCost_Multi()
                    End If
                ElseIf e.ColumnIndex = Me.ColCostNominal.Index Then
                    If Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekBilanganPositif(Me.dgvCost.Item(e.ColumnIndex, _
                                        e.RowIndex).Value) = False Then
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Else
                        'format tampilan
                        Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value, _
                        Me.ObjCurrency.GetDefaultPrecission(Me.dgvCost.Item(Me.ColPrimaryCurrencyCost.Index, e.RowIndex).Value)))
                        'hitung ulang AT-cost
                        Me.Hitung_Cost_Amount(Me.dgvCost.Rows(e.RowIndex))
                        Me.Hitung_AtCost_Multi()
                    End If
                ElseIf e.ColumnIndex = Me.ColAccounrCode.Index Then
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        Me.dgvCost.Item(Me.ColAccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New FrmMasterAkun_Cari
                        ij.BackColorAll = Me.BackColorAll
                        ij.txtCari.Text = Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                            Me.dgvCost.Item(Me.ColAccountDesc.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value = ij.mPrimaryAkun
                        ElseIf ij.mPrimaryAkun = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        ij = Nothing
                    End If
                End If
            Else
                If e.ColumnIndex = Me.ColAccounrCode.Index Then
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Me.dgvCost.Item(Me.ColAccountDesc.Index, e.RowIndex).Value = ""
                    Me.dgvCost.Item(Me.ColPrimaryAccount.Index, e.RowIndex).Value = -1
                Else
                    Me.dgvCost.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCost.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvCost)
            Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowDetailCalculation(ByVal CodeNo As String, ByVal rowIndex As Integer)
        Dim DetailCalculation As String = ""
        Dim row As DataGridViewRow
        Dim TotalOtherCost As Decimal = 0
        Dim TotalFOBPrice As Decimal = 0
        Dim SubTotal, NettPrice As Decimal
        Dim Currency As String = Me.dgvItem.Item(Me.ColCurrency.Index, rowIndex).Value.ToString
        Dim i As New FrmTransaksiAOPDetailCalc
        Dim AcuanRow As Integer

        Try
            'lakukan perhitungan FOB price
            DetailCalculation = "FOB Price(1) = Sub Total Before DP" + vbCrLf
            TotalFOBPrice = (Me.GetSubTotal_TanpaReturn_Multi * CDec(Me.txtTaxExchangeRate.Text))
            DetailCalculation = DetailCalculation + "                     = IDR " + Me.ObjCD.DigitGrouping(TotalFOBPrice) + vbCrLf + vbCrLf

            'hitung other cost
            DetailCalculation = DetailCalculation + "Other Cost(2) = "
            If Me.TxtRemark.Enabled = False Then
                AcuanRow = Me.dgvCost.Rows.Count
            Else
                AcuanRow = Me.dgvCost.Rows.Count - 1
            End If
            If (Me.TxtRemark.Enabled = True And Me.dgvCost.Rows.Count > 1) Or _
            (Me.TxtRemark.Enabled = False And Me.dgvCost.RowCount > 0) Then
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        If row.Index = 0 Then
                            DetailCalculation = DetailCalculation + row.Cells(Me.ColDescCost.Index).Value.ToString
                        Else
                            DetailCalculation = DetailCalculation + " + " + row.Cells(Me.ColDescCost.Index).Value.ToString
                        End If
                    End If
                Next
                DetailCalculation = DetailCalculation + vbCrLf
                DetailCalculation = DetailCalculation + "                      = "
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        If row.Index = 0 Then
                            DetailCalculation = DetailCalculation + "IDR (" + row.Cells(Me.ColCostNominal.Index).Value.ToString + _
                            " * " + row.Cells(Me.ColExcRateCost.Index).Value.ToString + ")"
                        Else
                            DetailCalculation = DetailCalculation + " + IDR (" + row.Cells(Me.ColCostNominal.Index).Value.ToString + _
                            " * " + row.Cells(Me.ColExcRateCost.Index).Value.ToString + ")"
                        End If
                    End If
                Next
                DetailCalculation = DetailCalculation + vbCrLf
                DetailCalculation = DetailCalculation + "                      = "
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        If row.Index = 0 Then
                            DetailCalculation = DetailCalculation + "IDR " + row.Cells(Me.ColCostAmount.Index).Value.ToString
                        Else
                            DetailCalculation = DetailCalculation + " + IDR " + row.Cells(Me.ColCostAmount.Index).Value.ToString
                        End If
                    End If
                Next
                For Each row In Me.dgvCost.Rows
                    If row.Index <> AcuanRow Then
                        TotalOtherCost += row.Cells(Me.ColCostAmount.Index).Value
                    End If
                Next
                DetailCalculation = DetailCalculation + vbCrLf
                DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalOtherCost))
                DetailCalculation = DetailCalculation + vbCrLf + vbCrLf
            Else
                DetailCalculation = DetailCalculation + " IDR 0" + vbCrLf + vbCrLf
            End If

            'hitung total cost
            DetailCalculation = DetailCalculation + "Total Cost(3) = FOB Price(1) + Other Cost(2)" + vbCrLf
            DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalFOBPrice)) _
            + " + IDR " + Me.ObjCD.DigitGrouping(CStr(TotalOtherCost)) + vbCrLf
            DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalFOBPrice + TotalOtherCost)) + vbCrLf + vbCrLf

            'lakukan perhitungan unit price
            DetailCalculation = DetailCalculation + "Unit price = Nett Price / SubTotal * Total Cost(3) " + vbCrLf
            For Each row In Me.dgvItem.Rows
                SubTotal += CDec(row.Cells(Me.ColAmount.Index).Value)
            Next

            ''@rosadinnn
            ''@MUaT
            If Me.dgvItem.Item(Me.ColQtyMulti_Item.Index, rowIndex).Value - Me.GetJmlDGVPR_WithBatch1_Multi1(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1, Me.dgvItem.Item(Me.ColBatchNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPONO.Index, rowIndex).Value) <> 0 Then
                NettPrice = Me.dgvItem.Item(Me.ColNettPrice.Index, rowIndex).Value
            Else
                NettPrice = 0
            End If

            DetailCalculation = DetailCalculation + "                = " + Currency + " (" + Me.dgvItem.Item(Me.ColAmount.Index, rowIndex).Value.ToString + " / " + _
            Currency + Me.ObjCD.DigitGrouping(SubTotal.ToString) + " * IDR " + Me.ObjCD.DigitGrouping((TotalFOBPrice + TotalOtherCost).ToString) + ")/ " + _
            Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(Me.ColQtyMulti_Item.Index, rowIndex).Value - Me.GetJmlDGVPR_WithBatch1_Multi1(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1, Me.dgvItem.Item(Me.ColBatchNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPONO.Index, rowIndex).Value)) + " " + Me.dgvItem.Item(Me.ColUnitsCodeMulti_Item.Index, rowIndex).Value.ToString + vbCrLf
            DetailCalculation = DetailCalculation + "                = IDR " + Me.dgvItem.Item(Me.ColATCost.Index, rowIndex).Value.ToString

            ''@rosadinnn
            ''@MUaT
            If Me.dgvItem.Item(Me.ColQtyMulti_Item.Index, rowIndex).Value - Me.GetJmlDGVPR_WithBatch1_Multi1(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1, Me.dgvItem.Item(Me.ColBatchNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPONO.Index, rowIndex).Value) = 0 Then
                DetailCalculation = DetailCalculation + vbCrLf + "(All returned, AT Cost =0)"
            End If

            i.TxtDetailCalculation.Text = DetailCalculation

            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub Hitung_Cost_Amount(ByVal row As DataGridViewRow)
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.txtPrimaryCurrency.Text <> "" And Me.txtPrimaryCurrency.Text <> "-1" Then
                row.Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColCostNominal.Index).Value * _
                row.Cells(Me.ColExcRateCost.Index).Value, DefPrec), DefPrec)
            Else
                row.Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping("1", DefPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Cost_Amount()
        Dim row As DataGridViewRow
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            For Each row In Me.dgvCost.Rows
                If Me.txtPrimaryCurrency.Text <> "" And Me.txtPrimaryCurrency.Text <> "-1" Then
                    row.Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColCostNominal.Index).Value * _
                    row.Cells(Me.ColExcRateCost.Index).Value, DefPrec), DefPrec)
                Else
                    row.Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping("1", DefPrec)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub Hitung_AtCost_Default()
    '    Dim TotalOngkosLain As Decimal = 0
    '    Dim TotalCost As Decimal = 0
    '    Dim SubTotal As Decimal
    '    Dim row As DataGridViewRow
    '    Dim Acuan As Integer
    '    Dim DefPrec As Integer = 0
    '    Dim JmlReturn As Decimal = 0
    '    Dim SubTotal1 As Decimal = 0
    '    Dim exchangRate As Decimal = 0
    '    Dim DummyCurrency As Double = -1
    '    Dim DefaultPrecisionRow As Integer = 0
    '    Dim SubTotalOriginal As Decimal = 0
    '    Dim JmlReturn1 As Decimal = 0
    '    Dim Jml As Decimal = 0

    '    Try
    '        DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

    '        If Me.dgvItem.Rows.Count Then
    '            row = Me.dgvItem.Rows(0)
    '            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
    '        End If
    '        For Each row In Me.dgvItem.Rows
    '            JmlReturn1 = Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '            row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)
    '            Jml = row.Cells(Me.ColTotalQty.Index).Value
    '            SubTotalOriginal += Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
    '            row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecisionRow))
    '        Next

    '        If Me.TxtRemark.Enabled = False Then
    '            Acuan = Me.dgvCost.Rows.Count
    '        Else
    '            Acuan = Me.dgvCost.Rows.Count - 1
    '        End If

    '        'hitung total ongkos lain - lain
    '        For Each row In Me.dgvCost.Rows
    '            If row.Index <> Acuan Then
    '                TotalOngkosLain += row.Cells(Me.ColCostAmount.Index).Value
    '            End If
    '        Next

    '        'For Each row In Me.dgvItem.Rows
    '        'SubTotal += row.Cells(Me.ColAmount.Index).Value
    '        'Next
    '        SubTotal = CDec(Me.txtSubTotalBeforeDP.Text)

    '        'hitung total harga FOB
    '        'jika rupiah maka langsung perhitungkan rupiah, jika tidak maka gunakan kurs konversi ke
    '        'rupiah
    '        TotalCost = Math.Round((CDec(Me.txtSubTotalBeforeDP.Text) * CDec(Me.txtTaxExchangeRate.Text)) + _
    '        TotalOngkosLain, DefPrec, MidpointRounding.AwayFromZero) ' --> duty stamp dan tax amount tidak diperhitungkan untuk At Cost
    '        'CDec(Me.TxtDutyStamp.Text) + Me.txtTaxAmount.Text + 

    '        'hitung unit price untuk masing - masing item

    '        For Each row In Me.dgvItem.Rows
    '            'yang diperhitungkan hanya yang tidak pake secondary unit
    '            If Me.cboCalculateBy.SelectedIndex = 0 Then
    '                'berdasarkan sales invoice
    '                exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
    '            ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
    '                'berdasarkan delivery order no
    '                exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
    '            ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
    '                'berdasarkan saler order no
    '                exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
    '            End If
    '            If DummyCurrency <> row.Cells(Me.ColPrimaryCurrency.Index).Value Then
    '                DummyCurrency = row.Cells(Me.ColPrimaryCurrency.Index).Value
    '                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
    '            End If

    '            If SubTotal = 0 Or TotalCost = 0 Then
    '                row.Cells(Me.ColATCost.Index).Value = 0
    '            Else
    '                If row.Cells(Me.ColTotalQty.Index).Value - Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryGRN.Index).Value, _
    '                -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value) <> 0 Then
    '                    row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
    '                    SubTotalOriginal, Me.ObjCurrency.GetDecimal_DefaultCurrency, MidpointRounding.AwayFromZero), DefPrec)

    '                    If row.Index = Me.dgvItem.RowCount - 1 Then
    '                        row.Cells(Me.Col_AtAmount.Index).Value = TotalCost - Me.GetTotalAtAmount_KecualiBarisAkhir
    '                    Else
    '                        'hitung amount stlh return
    '                        JmlReturn = Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '                        row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, _
    '                        row.Cells(Me.ColPONO.Index).Value)
    '                        SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
    '                        row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow))

    '                        row.Cells(Me.Col_AtAmount.Index).Value = Math.Round((SubTotal1 * TotalCost) / _
    '                        SubTotal, DefPrec, MidpointRounding.AwayFromZero)
    '                        row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDec(row.Cells(Me.Col_AtAmount.Index).Value) / CDec(row.Cells(Me.ColTotalQty.Index).Value - JmlReturn), _
    '                        DefPrec, MidpointRounding.AwayFromZero), DefPrec)
    '                    End If
    '                Else
    '                    row.Cells(Me.ColATCost.Index).Value = 0
    '                    row.Cells(Me.Col_AtAmount.Index).Value = 0
    '                End If
    '            End If
    '        Next

    '        Me.Col_AtAmount.Visible = False
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Function ApakahAtCostKonsisten() As Boolean
        Dim TotalOngkosLain As Decimal = 0
        Dim TotalCost As Decimal = 0
        Dim DefPrec As Integer
        Dim row As DataGridViewRow
        Dim TotalAtCost As Decimal = 0
        Dim Acuan As Integer

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            If Me.TxtRemark.Enabled = False Then
                Acuan = Me.dgvCost.Rows.Count
            Else
                Acuan = Me.dgvCost.Rows.Count - 1
            End If

            For Each row In Me.dgvCost.Rows
                If row.Index <> Acuan Then
                    TotalOngkosLain += row.Cells(Me.ColCostAmount.Index).Value
                End If
            Next

            TotalCost = Math.Round((CDec(Me.txtSubTotalBeforeDP.Text) * CDec(Me.txtTaxExchangeRate.Text)) + _
            TotalOngkosLain, DefPrec, MidpointRounding.AwayFromZero)

            For Each row In Me.dgvItem.Rows
                'cek minus
                If CDec(row.Cells(Me.Col_AtAmount.Index).Value()) < 0 Then
                    MsgBox("Inconsistent @ cost detected: negative value data found. Please contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                TotalAtCost += CDec(row.Cells(Me.Col_AtAmount.Index).Value())
            Next

            'cek total
            If TotalAtCost <> TotalCost Then
                MsgBox("Inconsistent @ cost detected. Please contact your vendor if you see this message", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Sub Hitung_AtCost_Multi()
        Dim TotalOngkosLain As Decimal = 0
        Dim TotalCost As Decimal = 0
        Dim SubTotal As Decimal
        Dim row As DataGridViewRow
        Dim Acuan As Integer
        Dim DefPrec As Integer = 0
        Dim JmlReturn As Decimal = 0
        Dim SubTotal1 As Decimal = 0
        Dim exchangRate As Decimal = 0
        Dim DummyCurrency As Double = -1
        Dim DefaultPrecisionRow As Integer = 0
        Dim SubTotalOriginal As Decimal = 0
        Dim JmlReturn1 As Decimal = 0
        Dim Jml As Decimal = 0

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.dgvItem.Rows.Count Then
                row = Me.dgvItem.Rows(0)
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            End If
            For Each row In Me.dgvItem.Rows
                Jml = row.Cells(Me.ColQtyMulti_Item.Index).Value

                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    JmlReturn1 = Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)
                    SubTotalOriginal += Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn1) * _
                    row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecisionRow))
                Else
                    JmlReturn1 = Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)
                    SubTotalOriginal += Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn1) * _
                    row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecisionRow))
                End If
            Next

            If Me.TxtRemark.Enabled = False Then
                Acuan = Me.dgvCost.Rows.Count
            Else
                Acuan = Me.dgvCost.Rows.Count - 1
            End If

            'hitung total ongkos lain - lain
            For Each row In Me.dgvCost.Rows
                If row.Index <> Acuan Then
                    TotalOngkosLain += row.Cells(Me.ColCostAmount.Index).Value
                End If
            Next

            'For Each row In Me.dgvItem.Rows
            'SubTotal += row.Cells(Me.ColAmount.Index).Value
            'Next
            SubTotal = CDec(Me.txtSubTotalBeforeDP.Text)

            'hitung total harga FOB
            'jika rupiah maka langsung perhitungkan rupiah, jika tidak maka gunakan kurs konversi ke
            'rupiah
            TotalCost = Math.Round((CDec(Me.txtSubTotalBeforeDP.Text) * CDec(Me.txtTaxExchangeRate.Text)) + _
            TotalOngkosLain, DefPrec, MidpointRounding.AwayFromZero) ' --> duty stamp dan tax amount tidak diperhitungkan untuk At Cost
            'CDec(Me.TxtDutyStamp.Text) + Me.txtTaxAmount.Text + 

            'hitung unit price untuk masing - masing item
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'reset data at cost dan at amount, supaya ketika get amount kecuali baris terakhir, nilainya selalu tepat
            '2025Creative
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColATCost.Index).Value = 0
                row.Cells(Me.Col_AtAmount.Index).Value = 0
            Next

            For Each row In Me.dgvItem.Rows
                'yang diperhitungkan hanya yang tidak pake secondary unit
                If Me.cboCalculateBy.SelectedIndex = 0 Then
                    'berdasarkan sales invoice
                    exchangRate = Me.dgvExchageRates.Rows(0).Cells(Me.ColExchangeRate.Index).Value
                ElseIf Me.cboCalculateBy.SelectedIndex = 1 Then
                    'berdasarkan delivery order no
                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryGRN.Index).Value)
                ElseIf Me.cboCalculateBy.SelectedIndex = 2 Then
                    'berdasarkan saler order no
                    exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryPO.Index).Value)
                End If
                If DummyCurrency <> row.Cells(Me.ColPrimaryCurrency.Index).Value Then
                    DummyCurrency = row.Cells(Me.ColPrimaryCurrency.Index).Value
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                End If

                row.Cells(Me.ColATCost_Unit.Index).Value = "/" + row.Cells(Me.ColUnitsCodeMulti_Item.Index).Value

                If SubTotal = 0 Or TotalCost = 0 Then
                    row.Cells(Me.ColATCost.Index).Value = 0
                Else
                    If row.Cells(Me.ColQtyMulti_Item.Index).Value - Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryGRN.Index).Value, _
                    -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value) <> 0 Then
                        'row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
                        '    SubTotalOriginal, Me.ObjCurrency.GetDecimal_DefaultCurrency, MidpointRounding.AwayFromZero), DefPrec)


                        If row.Index = Me.dgvItem.RowCount - 1 Or IsLastNonZeroQtyRow(row.Index) Then '2025Creative: baris terakhir yg qty <> 0, dianggap baris terakhir
                            row.Cells(Me.Col_AtAmount.Index).Value = TotalCost - Me.GetTotalAtAmount_KecualiBarisAkhir
                            JmlReturn = Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)
                            row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDec(row.Cells(Me.Col_AtAmount.Index).Value) / CDec(row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn), _
                                DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                        Else
                            'hitung amount stlh return
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                JmlReturn = Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)

                                SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(((row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn) * _
                                    row.Cells(Me.ColNettPrice.Index).Value + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value)) * exchangRate, DefaultPrecisionRow))
                            Else
                                JmlReturn = Me.GetJmlDGVPR_WithBatch1_Default1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                    row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)

                                SubTotal1 = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                                    row.Cells(Me.ColNettPrice.Index).Value + CDec(row.Cells(Me.ColItem_Adjustment.Index).Value)) * exchangRate, DefaultPrecisionRow))
                            End If

                            row.Cells(Me.Col_AtAmount.Index).Value = Math.Round((SubTotal1 * TotalCost) / _
                                SubTotal, DefPrec, MidpointRounding.AwayFromZero)
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                                'nilai @cost, per multi unit. jadi jmlr return selalu dalam unit multi
                                JmlReturn = Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                row.Cells(Me.ColPrimaryGRN.Index).Value, -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value)
                            End If
                            row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDec(row.Cells(Me.Col_AtAmount.Index).Value) / CDec(row.Cells(Me.ColQtyMulti_Item.Index).Value - JmlReturn), _
                                DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                        End If
                    Else
                        row.Cells(Me.ColATCost.Index).Value = 0
                        row.Cells(Me.Col_AtAmount.Index).Value = 0
                    End If
                End If
            Next

            Me.Col_AtAmount.Visible = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025Creative
    Private Function IsLastNonZeroQtyRow(ByVal rowIndex As Integer) As Boolean
        Dim row As DataGridViewRow
        Dim nextNonZeroQtyExist As Boolean = False

        Try
            row = Me.dgvItem.Rows(rowIndex)
            If row.Cells(Me.ColQtyMulti_Item.Index).Value - Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryGRN.Index).Value, _
                -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value) = 0 Then
                Return False 'return false jika baris tersebut qty = 0
            End If

            rowIndex = rowIndex + 1 'cek baris selanjutnya

            While rowIndex < Me.dgvItem.Rows.Count
                row = Me.dgvItem.Rows(rowIndex)

                If row.Cells(Me.ColQtyMulti_Item.Index).Value - Me.GetJmlDGVPR_WithBatch1_Multi1(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryGRN.Index).Value, _
                    -1, row.Cells(Me.ColBatchNo.Index).Value, row.Cells(Me.ColPONO.Index).Value) <> 0 Then
                    Return False 'return false jika ada baris selanjutnya yg qty > 0
                End If

                rowIndex = rowIndex + 1 'cek baris selanjutnya
            End While

            'return true bila:
            '(i) qty pada baris tsb <> 0
            '(ii) ada qty pada baris selanjutnya yg <> 0
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function GetTotalAtAmount_KecualiBarisAkhir() As Decimal
        Dim d As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                    d += CDec(rowDGV.Cells(Me.Col_AtAmount.Index).Value)
                End If
            Next

            Return d
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DGVdp_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DGVdp.UserDeletedRow
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.SetNumber(Me.DGVdp)
                Me.FormatDGV()
                Me.Hitung_Total_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutyStamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDutyStamp.ReadOnly = False Then
                    Me.TxtDutyStamp.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                Else
                    Me.TxtDutyStamp.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FilterDP(ByVal vCurrencyCode As String)
        Dim rowDGV As DataGridViewRow
        Dim AdaYangDihapus As Boolean = False
        Dim counter As Decimal = 0
        Try
            While counter < Me.DGVdp.RowCount
                rowDGV = Me.DGVdp.Rows(counter)
                If rowDGV.Cells(Me.ColDP_CurrencyDP.Index).Value.ToString <> Me.txtCurrencyCode.Text Then
                    AdaYangDihapus = True
                    Me.DGVdp.Rows.Remove(rowDGV)
                Else
                    counter += 1
                End If
            End While

            If AdaYangDihapus = True Then
                MsgBox("There several downpayment transaction became invalid due to currency code change. Proceeding to re-calculate amount.", _
                MsgBoxStyle.Information, "Information")
                Me.Hitung_Total_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' dgv pruchase return tidak bisa diedit
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DGVitemPR_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVitemPR.CellBeginEdit

        Try
            'yang bisa diedit hanya remarks
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex <> Me.ColRemarks_PurchaseReturn.Index Then
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
            'e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub DGVitemPR_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DGVitemPR.UserDeletedRow
        Try
            'load ulang data item purchase return
            Me.IsiDGVReturnItem()
            'hitung total
            Me.Hitung_Total_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentsTerms.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentsTerms.ReadOnly = True Then
                    Me.TxtPaymentsTerms.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPaymentsTerms.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentsTerms_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentsTerms.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPaymentsTerms, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentsTerms_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentsTerms.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPaymentsTerms, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentsTerms.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentsTerms.ReadOnly = True Then
                    Me.TxtPaymentsTerms.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPaymentsTerms.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True And Me.txtPrimarySupplier.Text <> "" And Me.txtPrimarySupplier.Text <> "-1" Then
                If Me.baru = False Then
                    If Me.ObjAOP.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjAOP.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtAOP.Text = Me.ObjAOP.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                    End If
                Else
                    Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, Me.txtPrimarySupplier.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCost.RowsAdded
        Dim DefPrecBase As Integer = 0

        Try
            DefPrecBase = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.dgvCost.AllowUserToAddRows = True Then
                With Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                    .Cells(Me.ColDescCost.Index).Value = "-"
                    .Cells(Me.ColCurrencyCost.Index).Value = Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
                    .Cells(Me.ColPrimaryCurrencyCost.Index).Value = Core_Module.PrimaryBaseCurrency
                    .Cells(Me.ColCostNominal.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                    .Cells(Me.ColRemarksCost.Index).Value = ""
                    .Cells(Me.ColExcRateCost.Index).Value = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                    .Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                    .Cells(Me.ColAccounrCode.Index).Value = ""
                    .Cells(Me.ColAccountDesc.Index).Value = ""
                    .Cells(Me.ColPrimaryAccount.Index).Value = -1
                End With
            End If
            ObjCD.SetNumber(Me.dgvCost)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCost_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvCost.UserAddedRow
        Dim DefPrecBase As Integer = 0

        Try
            DefPrecBase = Me.ObjCurrency.GetDecimal_DefaultCurrency
            With Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 1)
                .Cells(Me.ColDescCost.Index).Value = "-"
                .Cells(Me.ColCurrencyCost.Index).Value = Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
                .Cells(Me.ColPrimaryCurrencyCost.Index).Value = Core_Module.PrimaryBaseCurrency
                .Cells(Me.ColCostNominal.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                .Cells(Me.ColRemarksCost.Index).Value = ""
                .Cells(Me.ColExcRateCost.Index).Value = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)
                .Cells(Me.ColCostAmount.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecBase)
                .Cells(Me.ColAccounrCode.Index).Value = ""
                .Cells(Me.ColAccountDesc.Index).Value = ""
                .Cells(Me.ColPrimaryAccount.Index).Value = -1
            End With
            ObjCD.SetNumber(Me.dgvCost)
            Me.Hitung_Cost_Amount(Me.dgvCost.Rows(Me.dgvCost.Rows.Count - 2))
            Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPRNumber() As String
        Dim ListPrimary As New List(Of Double)
        Dim row As DataGridViewRow
        Dim rowpR As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_view1Row
        Dim Retval As String = ""

        Try
            'kosongkan baris yang diberikan
            Me.TBLPR.Rows.Clear()

            'cari data SR, berdasarkan data do yang dituliskan
            For Each row In Me.dgvItem.Rows
                If ListPrimary.Contains(row.Cells(Me.ColPrimaryGRN.Index).Value) = False Then
                    ListPrimary.Add(row.Cells(Me.ColPrimaryGRN.Index).Value)
                End If
            Next
            If IsNumeric(Me.TxtPrimaryTransaksiAOP.Text) = False Then
                Me.TxtPrimaryTransaksiAOP.Text = "-1"
            End If
            Me.ObjPR.IsiTBLPRByPrimaryGRN(ListPrimary, TBLPR, Me.TxtPrimaryTransaksiAOP.Text)
            For Each rowpR In Me.TBLPR.Rows
                'cek bila masih ada purchase return yang belum dimasukkan ke dalam AOP
                If Me.ApakahAdaDalamDGVReturn(rowpR.PR_No) = False Then
                    If Retval = "" Then
                        Retval = rowpR.PR_No
                    Else
                        Retval = Retval + ", " + rowpR.PR_No
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListPrimary = Nothing
        End Try
    End Function

    Private Sub btnPrint_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim obj As New ClsTransaksiAOP
        Dim ij As New FrmCetakTransaksiAOP
        Dim ObjCust As New CustomSMESS.ClsTransaksiAOP

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.TxtPrimaryTransaksiAOP.Text = "" Or Me.TxtPrimaryTransaksiAOP.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryTransaksiAOP.Text <> "" And Me.TxtPrimaryTransaksiAOP.Text <> "-1" Then
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) <> 1 Then
                    obj.CetakAop(ij.crv1, Me.TxtPrimaryTransaksiAOP.Text)
                    ij.ShowDialog()
                Else
                    If ObjCust.CetakCustomAOP(ij.crv1, Me.TxtPrimaryTransaksiAOP.Text, Me.txtPrimaryCurrency.Text) = True Then
                        ij.ShowDialog()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
            ij.Dispose()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

                ''@MUaT
                Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFillBy_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFillBy.SelectedValueChanged
        Try
            If Me.cboFillBy.SelectedIndex = -1 Then
                Me.cboDO.DataSource = Nothing
            Else
                If Me.txtPrimarySupplier.Text <> "" And Me.txtPrimarySupplier.Text <> "-1" Then
                    If Me.TxtRemark.Enabled = True Then
                        Me.IsiCboDeliveryOrder2(Me.txtPrimarySupplier.Text)
                    Else
                        Me.IsiCboDeliveryOrder2(Me.txtPrimarySupplier.Text)
                    End If
                Else
                    Me.cboDO.DataSource = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtNPWPNo.Enabled = True Then
                    If Me.TxtNPWPNo.ReadOnly = True Then
                        Me.TxtNPWPNo.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtNPWPNo.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtNPWPNo.Enabled = True Then
                    If Me.TxtNPWPNo.ReadOnly = True Then
                        Me.TxtNPWPNo.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtNPWPNo.BackColor = Me.txtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoiceNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoiceNumber.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtTaxInvoiceNumber.Enabled = True Then
                    If Me.TxtTaxInvoiceNumber.ReadOnly = True Then
                        Me.TxtTaxInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtTaxInvoiceNumber.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoiceNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoiceNumber.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtTaxInvoiceNumber.Enabled = True Then
                    If Me.TxtTaxInvoiceNumber.ReadOnly = True Then
                        Me.TxtTaxInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.TxtTaxInvoiceNumber.BackColor = Me.txtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.ObjAOP.CariDetail(vPrimary, _
                 Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me) _
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


    'Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged

    '    Try
    '        If Me.txtPrimaryCurrTax.Text = -1 Then
    '            If Me.txtPrimaryCurrency.Text <= 1 Then
    '                Me.cboCurrTax.Enabled = False
    '            Else
    '                Me.cboCurrTax.Enabled = True
    '            End If

    '            Me.IsiCBOCurrencyTax()
    '            Me.IsiCBOArAkunByTipePerkiraan1()
    '        End If


    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub


    Private Sub btnStopProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStopProcess.Click
        Try
            If MsgBox("Stop AOP creation process?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.LanjutCreateAOP = False
                Me.LangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAdjustDPP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAdjustDPP.LinkClicked
        Dim ij As New FrmTransaksiAOP_AdjustDPP
        Dim rowDGV, rowDGV1 As DataGridViewRow

        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Exit Sub
            End If

            ij.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'isikan data item
            For Each rowDGV In Me.dgvItem.Rows
                ij.dgvItem.Rows.Add()
                rowDGV1 = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                For wk As Integer = 0 To Me.dgvItem.Columns.Count - 1
                    rowDGV1.Cells(wk).Value = rowDGV.Cells(wk).Value
                Next
            Next
            'set kolom visible
            For wk As Integer = 0 To Me.dgvItem.Columns.Count - 1
                ij.dgvItem.Columns(wk).Visible = Me.dgvItem.Columns(wk).Visible
            Next
            ij.dgvItem.Columns(ij.ColATCost.Index).Visible = False
            ij.dgvItem.Columns(ij.Col_AtAmount.Index).Visible = False
            ij.dgvItem.Columns(ij.ColShowDetail.Index).Visible = False
            ij.dgvItem.Columns(ij.ColItem_Adjustment.Index).Visible = True
            ij.dgvItem.Columns(ij.ColItem_AdjustmentReason.Index).Visible = True
            ij.dgvItem.Columns(ij.ColItem_AdjustUser.Index).Visible = True
            ij.dgvItem.Columns(ij.ColItemAdjustmentTimestamp.Index).Visible = True

            ij.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            ij.BackColorAll = Me.BackColorAll
            ij.UserCode = Me.UserCode
            If Me.TxtRemark.Enabled = False Then
                ij.IsReadOnly = True
            Else
                ij.IsReadOnly = False
            End If

            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                If ij.IsReadOnly = False Then
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.dgvItem.Rows.Clear()

                    For Each rowDGV In ij.dgvItem.Rows
                        Me.dgvItem.Rows.Add()
                        rowDGV1 = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        For wk As Integer = 0 To Me.dgvItem.Columns.Count - 1
                            rowDGV1.Cells(wk).Value = rowDGV.Cells(wk).Value
                        Next
                    Next

                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    Me.Hitung_Total_Multi()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Public Sub FromCutOff_AOP_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal vTblPrimaryPO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryGrn As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryAOPDP As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryPR As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core)
        Dim TblTempPO, TblTempGrn, TblTempAOPDP, TblTempPR As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.ObjAOP.ResetObject(Me.dgvItem, Me.DGVdp, Me.dgvExchageRates, Me.DGVitemPR, Me.dgvCost)
            Me.ObjAOP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP
            Me.ObjAOP.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me)

            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                'Update Primary GRN
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryGRN.Index).Value) = False And rowD.Cells(Me.ColPrimaryGRN.Index).Value <> "-1" Then
                    TblTempGrn.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryGRN.Index).Value)
                    If TblTempGrn.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryGRN.Index).Value), vTblPrimaryGrn)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryGRN.Index).Value), PrimaryUpdate, TblTempGrn)
                        rowD.Cells(Me.ColPrimaryGRN.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryGRN.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryGRN.Index).Value), TblTempGrn)
                    End If
                End If

                'Update Primary PO
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryPO.Index).Value) = False And rowD.Cells(Me.ColPrimaryPO.Index).Value <> "-1" Then
                    TblTempPO.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value)
                    If TblTempPO.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), vTblPrimaryPO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), PrimaryUpdate, TblTempPO)
                        rowD.Cells(Me.ColPrimaryPO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryPO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), TblTempPO)
                    End If
                End If
            Next

            If Me.DGVdp.RowCount > 0 Then
                For Each rowD_DP As DataGridViewRow In Me.DGVdp.Rows
                    'Update Primary AOP DP
                    If String.IsNullOrEmpty(rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value) = False And rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value <> "-1" Then
                        TblTempAOPDP.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value)
                        If TblTempAOPDP.DefaultView.Count = 0 Then
                            PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value), vTblPrimaryAOPDP)
                            ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value), PrimaryUpdate, TblTempAOPDP)
                            rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value = PrimaryUpdate
                        Else
                            rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD_DP.Cells(Me.ColDP_PrimaryAOPDP.Index).Value), TblTempAOPDP)
                        End If
                    End If
                Next
            End If

            If Me.DGVitemPR.RowCount > 0 Then
                For Each rowD_PR As DataGridViewRow In Me.DGVitemPR.Rows
                    'Update Primary Purchase Return
                    If String.IsNullOrEmpty(rowD_PR.Cells(Me.ColPrimaryPR.Index).Value) = False And rowD_PR.Cells(Me.ColPrimaryPR.Index).Value <> "-1" Then
                        TblTempPR.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD_PR.Cells(Me.ColPrimaryPR.Index).Value)
                        If TblTempPR.DefaultView.Count = 0 Then
                            PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD_PR.Cells(Me.ColPrimaryPR.Index).Value), vTblPrimaryPR)
                            ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD_PR.Cells(Me.ColPrimaryPR.Index).Value), PrimaryUpdate, TblTempPR)
                            rowD_PR.Cells(Me.ColPrimaryPR.Index).Value = PrimaryUpdate
                        Else
                            rowD_PR.Cells(Me.ColPrimaryPR.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD_PR.Cells(Me.ColPrimaryPR.Index).Value), TblTempPR)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_AOP_Save() As Double

        Try
            Me.baru = True
            Me.ObjAOP.SimpanBaru1(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me, False, True)
            Return Me.ObjAOP.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    '2025 Core Tax Tahap 2.2
    Private Sub txtDPPNilaiLain_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.Enter
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtDPPNilaiLain, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtDPPNilaiLain.Enabled = True Then
                    If Me.txtDPPNilaiLain.ReadOnly = True Then
                        Me.txtDPPNilaiLain.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.txtDPPNilaiLain.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDPPNilaiLain.KeyDown
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.Number_KeyDown(Me.txtDPPNilaiLain, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDPPNilaiLain.KeyUp
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.Number_KeyUp(Me.txtDPPNilaiLain, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(Me.txtDPPNilaiLain.Text, Me.DefaultPrecision)
                'Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtDPPNilaiLain.Enabled = True Then
                    If Me.txtDPPNilaiLain.ReadOnly = True Then
                        Me.txtDPPNilaiLain.BackColor = Me.txtReadOnly.BackColor
                    Else
                        Me.txtDPPNilaiLain.BackColor = Me.txtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.2
    Private Sub UbahDPPNilaiLain()
        Try
            'DPP Nilai Lain untuk 04, sementara 11/12, tapi bisa diedit
            If Me.chkTax.Checked Then
                If (String.IsNullOrEmpty(Me.TxtTaxInvoiceNumber.Text) = False And Me.TxtTaxInvoiceNumber.Text.Length >= 2) _
                AndAlso Me.TxtTaxInvoiceNumber.Text.Substring(0, 2) = "04" Then
                    Me.txtDPPNilaiLain.ReadOnly = True
                    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotalAfterDP.Text) * 11 / 12, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                ElseIf (String.IsNullOrEmpty(Me.TxtTaxInvoiceNumber.Text) = False And Me.TxtTaxInvoiceNumber.Text.Length >= 2) _
                AndAlso Me.TxtTaxInvoiceNumber.Text.Substring(0, 2) = "05" Then
                    'DPP Nilai Lain untuk 05, sementara 10% dari total DPP, bisa diedit
                    Me.txtDPPNilaiLain.ReadOnly = True
                    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotalAfterDP.Text) * 100 / 100, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                Else
                    Me.txtDPPNilaiLain.ReadOnly = True
                    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotalAfterDP.Text), Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoiceNumber_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoiceNumber.TextChanged
        Try
            If Me.chkTax.Checked And String.IsNullOrEmpty(Me.TxtTaxInvoiceNumber.Text) = False Then
                If CDec(Me.txtSubTotalAfterDP.Text) > 0 Then
                    Me.UbahDPPNilaiLain()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub
    '2025 Core Tax Tahap 2.2 end
End Class