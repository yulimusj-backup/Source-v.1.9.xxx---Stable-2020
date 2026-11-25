Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiAOP_Quick2

    Dim tblCboUnit As New DseTransaksiPO.CboUnitDataTable
    Dim rowCboUnit As DseTransaksiPO.CboUnitRow
    Dim tblMultiUnit As New DseTransaksiPO.tblmaster_product_item_multiunitDataTable
    Dim rowMultiUnit As DseTransaksiPO.tblmaster_product_item_multiunitRow
    Dim objUnit As New Core_Module.ClsMasterUnit
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjAOP As New ClsTransaksiAOP
    Public baru As Boolean
    Public UpdateAtAmount As Boolean = False
    Public UpdatePrg As ProgressBar

    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    'Dim ObjSupplierProductCode As New clsTransaksiSettingSupplierProductCodeNo
    Dim objGRN As New ClsTransaksiGRN
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim Objsupplier As New clsMasterSupplier
    Dim ObjAOPItem As New ClsTransaksiAOPItem
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjPOItem As New Inventory_Module.ClsTransaksiPOItem
    Dim ObjPOItem2 As New ClsTransaksiPOItem
    Dim ObjPR As New ClsTransaksiPurchaseReturn
    Dim ObjPRItem As New ClsTransaksiPurchaseReturnItem
    Dim ObjCL As New ClsTransaksiMonthlyClosing
    Dim ObjAkun As New ClsMasterAkun_Purchase
    Dim ObjLicense As New ClsLicense
    Dim ObjItemInOut As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjItemBundling As New Inventory_Module.clsTransaksiStockBundlingItem
    Dim ObjStockInItem As New Inventory_Module.clsTransaksiStockInOutItem
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjNumbering As New ClsMasterTransactionNumbering

    Dim DefaultPrecision As Integer = 0
    Dim buffer As String = ""
    Dim buffer2 As String = ""
    Dim BufferPONumber As String = ""
    'Dim SedangRevisi As Boolean = False
    Dim TBL As New DseTransaksiGRN.DataTable8DataTable
    Dim TBL2 As New DseTransaksiAOP.DataTable9DataTable
    Dim TBLPR As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_view1DataTable
    Dim DTAPO As New DseTransaksiGRNTableAdapters.tbltrans_po_item_view5TableAdapter
    Dim TBLPO As New DseTransaksiGRN.tbltrans_po_item_view5DataTable
    Dim ObjTaxRate As New Sales_Module.ClsTransaksiSalesInvoiceTaxRate

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
    Dim visibleCBO As Boolean = False

    Dim IndexAsal As Integer = -1
    Dim IndexTujuan As Integer = -1


    Dim GetHistoricalPrice As Boolean = True

    '2025 Core Tax Tahap 2.1
    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo

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

    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged

        Try
            If IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso CDec(Me.txtPrimaryCurrency.Text) > -1 Then
                If Me.TxtRemark.Enabled = True Then
                    Me.IsiCBOArAkunByTipePerkiraan1()
                    Me.IsiCBOCurrencyTax()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        Dim row1 As AccModule.DseMultiAccount.tblisi_cbo_ar_akunRow
        Try
            'Me.txtPrimaryAccount.Visible = True
            If Me.cboAccount.SelectedIndex = -1 Then
                Me.txtPrimaryAccount.Text = "-1"
            Else
                row1 = Me.tblAccount.DefaultView(Me.cboAccount.SelectedIndex).Row
                Me.txtPrimaryAccount.Text = row1.Primary
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub txtPrimaryAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryAccount.TextChanged
    '    Dim PrimaryARAkun As Double = -1

    '    Try
    '        If IsNumeric(Me.txtPrimaryAccount.Text) Then
    '            PrimaryARAkun = CDbl(Me.txtPrimaryAccount.Text)
    '        End If

    '        If Me.tblAccount.Rows.Count > 0 Then
    '            For IntA As Integer = 0 To Me.tblAccount.Rows.Count - 1
    '                Me.rowAccount = Me.tblAccount.Rows(IntA)
    '                With Me.rowAccount
    '                    If .Primary = PrimaryARAkun Then
    '                        Me.cboPrimaryAccount.SelectedIndex = IntA
    '                        Me.cboAccount.SelectedIndex = IntA
    '                    End If
    '                End With
    '            Next
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim UseTaxSetting, TaxPercentSetting As String

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
              MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TabControl1.SelectedIndex = 0
            Me.dtpDate.Focus()
            Me.ObjAOP.ResetObject(Me.dgvItem, Me.DGVdp, Me.dgvExchageRates, Me.DGVitemPR, Me.dgvCost)
            'pertama kali dianggap supplier lokal dulu
            Me.TxtAOP.Text = Me.ObjAOP.GetNoBaru(Me.dtpDate.Value, -1)
            'Me.txtNoFakturPajak.Text = Me.ObjAOP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
            Me.ObjAOP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP

            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.AOP, ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.txtTax.Text = "10.00"
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

    Public Sub AktifkanDetail()
        Dim TidakAdaPO As Boolean = True
        Dim row As DataGridViewRow

        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.KendalikanControl(Me.SplitContainer7.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer7.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)

            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)

            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.DGVdp.AllowUserToDeleteRows = True
            Me.DGVitemPR.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.dgvExchageRates.ReadOnly = False
            'set dgv cost, boleh edit, bisa tambahkan baris baru
            Me.dgvCost.ReadOnly = False
            Me.dgvCost.AllowUserToAddRows = True
            Me.dgvCost.AllowUserToDeleteRows = True

            Me.llbAddress.Enabled = True
            Me.llbAddDP.Enabled = True
            Me.llbAddSR.Enabled = True

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    TidakAdaPO = False
                    Exit For
                End If
            Next
            If TidakAdaPO = False Then
                Me.btnFindSupplier.Enabled = False
                Me.txtSupplierCode.Enabled = False
                Me.txtSupplierCode.TabStop = False
                Me.btnFindSupplier.TabStop = False
            Else
                Me.btnFindSupplier.Enabled = True
                Me.txtSupplierCode.Enabled = True
                Me.txtSupplierCode.TabStop = True
                Me.btnFindSupplier.TabStop = True
            End If

            Me.cboAccount.Enabled = True
            'Me.cboPrimaryAccount.Enabled = True
            Me.cboCurrTax.Enabled = True
            Me.cboPrimaryCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
            End If

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            Me.chkPONumber.Checked = False
            Me.cboPurchaseOrder.Enabled = False
            Me.SetTaxRate(False)

            Me.LinkLabel1.Enabled = True
            Atur1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.ObjAOP
                Dim defPrec As Integer = 0
                Me.cboCalculateBy.SelectedIndex = .CalculatesBy
                Me.TxtPrimaryTransaksiAOP.Text = .Primary
                Me.TxtAOP.Text = .NoTransaksi
                Me.txtDownPayment.Text = .DownPayment
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
                Me.TxtPaymentsTerms.Text = Me.ObjCD.DigitGrouping(.PaymentTerms)
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubtotalBeforeDP, Me.DefaultPrecision))
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxExchangeRate, Me.DefaultPrecision))
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.DownPayment, Me.DefaultPrecision))

                Me.txtExhangeRateToIDR.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExchangRateToIDR, Me.DefaultPrecision))
                Me.txtExchangeRateToUSD.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExhangeRateToUSD, Me.DefaultPrecision))
                If .TaxExclude = ClsTransaksiAOP.EnumTaxExclude.Exclude Then
                    Me.chkTax.Checked = True
                Else
                    Me.chkTax.Checked = False
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(.Tax, 2)
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(.TaxExchangeRate, 2)
                Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = .EFak_KodeJenisTransaksiPajak
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                Me.IsiDGVReturnItem(False)

                Me.IsiCBOArAkunByTipePerkiraan1()
                If .PrimaryAccount > -1 Then
                    Me.txtPrimaryAccount.Text = .PrimaryAccount
                    'Me.cboPrimaryAccount.Text = .PrimaryAccount
                    Me.cboAccount.SelectedIndex = Me.GetIndexAkun(Me.txtPrimaryAccount.Text)
                Else
                    Me.cboAccount.DataSource = Nothing
                    'Me.cboPrimaryAccount.DataSource = Nothing
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

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                If Me.ObjAOP.VATDipungut = True Then
                    Me.cboPPNDipungut.SelectedIndex = 0
                Else
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If
                Me.TxtNPWPNo.Text = .NoNPWPSupplier
                Me.TxtTaxInvoiceNumber.Text = .NoFakturPajak

                '2025 Core Tax Tahap 2.2
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(.DPPNilaiLain, Me.DefaultPrecision)
            End With


            Me.FormatDGV()
            Me.FormatDGV2()
            'tampilan exchange rate
            Me.SetDGVExchangeRate()
            Me.Hitung_Total_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexAkun(ByVal vPrimaryAkun As Double) As Integer
        Dim rowAccount1 As AccModule.DseMultiAccount.tblisi_cbo_ar_akunRow

        Try
            For ij As Integer = 0 To Me.tblAccount.DefaultView.Count - 1
                rowAccount1 = Me.tblAccount.DefaultView(ij).Row
                If rowAccount1.Primary = vPrimaryAkun Then
                    Return ij
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjAOP
                .Primary = CDbl(Me.TxtPrimaryTransaksiAOP.Text)
                .NoTransaksi = Me.TxtAOP.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimarySupplier = CDbl(Me.txtPrimarySupplier.Text)
                .SupplierInvoiceNo = Me.txtSupplierInvoiceNo.Text
                .SupplierInvoiceDate = Me.DtpInvDate.Value
                .Lokal = ClsTransaksiAOP.EnumLokal.Lokal
                .SupplierInvoiceNo = Me.txtSupplierInvoiceNo.Text
                .SupplierInvoiceDate = Me.DtpInvDate.Value
                .PrimaryCurrency = CDbl(Me.txtPrimaryCurrency.Text)
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
                .EFak_KodeJenisTransaksiPajak = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
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
                .NoFakturPajak = Me.TxtTaxInvoiceNumber.Text
                .NoNPWPSupplier = Me.TxtNPWPNo.Text
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

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.KendalikanControl(Me.SplitContainer7.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer7.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)

            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)

            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit

            Me.DGVdp.AllowUserToDeleteRows = False
            Me.DGVitemPR.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvExchageRates.ReadOnly = True
            Me.dgvCost.ReadOnly = True
            Me.dgvCost.AllowUserToDeleteRows = False
            Me.dgvCost.AllowUserToAddRows = False
            Me.dgvItem.ReadOnly = True

            Me.llbAddress.Enabled = True
            Me.llbAddDP.Enabled = False
            Me.llbAddSR.Enabled = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

            Me.chkPONumber.Checked = False
            Me.cboCodeNo.DataSource = Nothing
            Me.cboCodeNo.Items.Clear()
            Me.cboPurchaseOrder.DataSource = Nothing
            Me.cboPurchaseOrder.Items.Clear()
            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False
            Me.cboPrimaryCurrTax.Enabled = False
            'Me.cboPrimaryAccount.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KosongkanControl(Me.SplitContainer7.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer7.Panel2.Controls, True, True, True)

            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage4.Controls, True, True, True)

            Me.TxtPrimaryTransaksiAOP.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimarySupplier.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtPrimaryAccount.Text = "-1"
            Me.txtSubTotalBeforeDP.Text = "0"
            Me.txtSubTotalAfterDP.Text = "0"
            Me.txtTax.Text = "0"
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", 2)
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.txtDownPayment.Text = "0"
            Me.TxtDutyStamp.Text = "0"
            Me.TxtPaymentsTerms.Text = "0"
            Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = "-1"
            Me.BufferPONumber = ""
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
            'Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.cboCurrTax.DataSource = Nothing

            Me.cboSupplierAddress.Text = ""
            'Me.cboPrimaryAccount.Text = ""
            Me.cboAccount.Text = ""
            Me.cboCurrTax.Text = ""

            Me.cboSupplierAddress.Items.Clear()
            'Me.cboPrimaryAccount.Items.Clear()
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.cboAccount.Items.Clear()
            Me.cboCurrTax.Items.Clear()

            Me.chkPONumber.Checked = False
            Me.cboCodeNo.DataSource = Nothing
            Me.cboCodeNo.Items.Clear()
            Me.cboPurchaseOrder.DataSource = Nothing
            Me.cboPurchaseOrder.Items.Clear()

            Me.TblCustomDummy.Rows.Clear()
            Me.TblItemCustomInOut.Rows.Clear()

            '2025 Core Tax Tahap 2.2 
            Me.txtDPPNilaiLain.Text = "0"
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
        Dim PrimaryAccountAP1, PrimaryAccountAP2, PrimaryCurrency As Double
        Dim ListPrimaryProduct As New List(Of Double)
        Dim ObjAOPCustom As New CustomSMESS.ClsTransaksiAOP

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '~ab~
            'Remarks : tambahan batasan 
            '25/02/2020
            Dim DTA_Convt As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_accounting_closingTableAdapter
            If DTA_Convt.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA_Convt.ScalarQuery_GetConvertionDate.Value.Date > Me.dtpDate.Value Then
                    MsgBox("Please specify transaction date after convertion date", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox("Please specify transaction date after convertion date", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '

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
                        Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("AOP number")
            KoleksiTextBox.Add(Me.TxtAOP)
            KoleksiNamaData.Add("Supplier code")
            KoleksiTextBox.Add(Me.txtSupplierCode)
            KoleksiNamaData.Add("Currency")
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

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

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
                If MsgBox("The purchase return: " + Dummy + " refers to product in this AOP, continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            'jika currency yang digunakan adalah IDR, maka tax rate harus bernilai 1
            If Me.txtPrimaryCurrency.Text = "1" Then
                If CDbl(Me.txtTaxExchangeRate.Text) <> 1 Then
                    MsgBox("For invoice with " + Me.ObjCurrency.GetCurrencyCode(1) + " currency, the tax exchange rate value must be 1.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'jika bukan IDR dan tax ratenya bernilai 1 maka berikan peringatan
                If CDbl(Me.txtTaxExchangeRate.Text) = 1 Then
                    MsgBox("This invoice use non-" + Me.ObjCurrency.GetCurrencyCode(1) + " currency and the tax exchange rate value = 1. This is wrong data input.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
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
                        If CDbl(row_dgv.Cells(Me.ColExchangeRate.Index).Value) = 1 Then
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
                        If CDbl(row_dgv.Cells(Me.ColExchangeRate.Index).Value) <> 1 Then
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
                    Me.TabControl1.SelectedIndex = 0
                    Me.dgvCost.ClearSelection()
                    rowDGV.Selected = True
                    Me.dgvCost.Focus()
                    Return False
                End If
                If rowDGV.Cells(Me.ColDescCost.Index).Value.ToString = "" And rowDGV.Index <> Me.dgvCost.RowCount - 1 Then
                    MsgBox("Please specify cost description data on row " + (rowDGV.Index + 1).ToString + ".", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
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
                If row_dgv.Cells(Me.ColCurrencyDP.Index).Value.ToString <> Me.txtCurrencyCode.Text Then
                    If Dummy = "" Then
                        Dummy = row_dgv.Cells(Me.ColAOP_DP_No.Index).Value.ToString
                    Else
                        Dummy = Dummy + ", " + row_dgv.Cells(Me.ColAOP_DP_No.Index).Value.ToString
                    End If
                End If
            Next
            If Dummy <> "" Then
                MsgBox("The following downpayment used different currency with this AOP: " _
                + Dummy + ". We cannot enter the data.", MsgBoxStyle.Information, "Information")
            End If

            'update data payment terms, tidak dilakukan karena tidak merujuk pada po
            'Me.UpdatePaymentTerm()

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

                        TipeAkun2 = Me.ObjAkun.GetTipeAkun2(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value)
                        If TipeAkun2 = ClsMasterAkun_Purchase.EnumTipeAkun2.Account_Payable Then
                            'jika tipenya account payable, maka primary akun harus primary akun ap untuk curency AOP atau primary akun ap dr IDR
                            If IsNumeric(Me.txtPrimaryAccount.Text) Then
                                PrimaryAccountAP1 = CDbl(Me.txtPrimaryAccount.Text)
                            End If
                            'PrimaryAccountAP1 = Me.ObjCurrency.GetPrimaryAkunAP(Me.txtPrimaryCurrency.Text)
                            PrimaryAccountAP2 = Me.ObjCurrency.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)


                            If PrimaryAccountAP1 = PrimaryAccountAP2 Then
                                If CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP1 And _
                                                                 CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP2 Then
                                    MsgBox("Please only use account payable: '" & Me.ObjAkun.GetDescription2(PrimaryAccountAP1) & _
                                    "' for cost data on row " & (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                                    Me.TabControl1.SelectedIndex = 0
                                    Me.dgvCost.ClearSelection()
                                    Me.dgvCost.Rows(rowDGV.Index).Selected = True
                                    Return False
                                End If
                            Else
                                If CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP1 And _
                                   CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP2 Then
                                    MsgBox("Please use account payable: '" + Me.ObjAkun.GetDescription2(PrimaryAccountAP1) + _
                                    "' for cost data on row " + _
                                    (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")

                                    Me.TabControl1.SelectedIndex = 0
                                    Me.dgvCost.ClearSelection()
                                    Me.dgvCost.Rows(rowDGV.Index).Selected = True

                                    Return False
                                End If
                            End If


                            'currency harus sama
                            PrimaryCurrency = Me.ObjAkun.GetPrimaryCurrency(CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value))
                            If PrimaryCurrency <> CDbl(rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value) Then
                                MsgBox("Please use currency: " + Me.ObjCurrency.GetCurrencyCode(PrimaryCurrency) + _
                                " for cost data on row " + (rowDGV.Index + 1).ToString + ".", MsgBoxStyle.Information, "Information")
                                Me.TabControl1.SelectedIndex = 0
                                Me.dgvCost.ClearSelection()
                                Me.dgvCost.Rows(rowDGV.Index).Selected = True

                                Return False
                            Else
                                'jika currency sama, jika akun AP yang digunakan sebagai akun ongkos, maka pastikan 
                                'akun ap yang digunakan sama dengan akun AP di header transaksi
                                If CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) <> PrimaryAccountAP1 Then
                                    MsgBox("Please use account payable: '" + Me.ObjAkun.GetDescription2(PrimaryAccountAP1) + _
                                     "' for cost data on row " + _
                                     (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")

                                    Me.TabControl1.SelectedIndex = 0
                                    Me.dgvCost.ClearSelection()
                                    Me.dgvCost.Rows(rowDGV.Index).Selected = True

                                    Return False
                                End If
                            End If

                            'jika account payablenya non idr/sama currency-nya dengan currency AOP, maka exchange ratenya harus 
                            'sama dengan tax exchange rate, jika tidak maka balance menjadi salah hitung
                            'jika kemudian hari ditemukan kasus seperti ini, maka yang harus dimodifikasi adalah clstransaksi aop pada bagian
                            'fungsi 'Transfer Jurnal'-nya
                            If CDbl(rowDGV.Cells(Me.ColPrimaryAccount.Index).Value) = PrimaryAccountAP1 Then
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
                    MsgBox("The tax invoice serial number must starts with one of the following alternatives: 01, 02, 03, 04, 06, 07, 08, 09 or 10.", _
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
                '2025 Core Tax Tahap 2.2 End
            End If

            'cek akun duty stamp
            If CDbl(Me.TxtDutyStamp.Text) > 0 Then
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

            '@at_amount tidak boleh terlalu kecil sehingga unit price per satuan terkecil barang tersebut menjadi bernilai 0
            If Me.CekATAmount = False Then
                Return False
            End If
            'cek konsistensi @ cost
            If Me.ApakahAtCostKonsisten = False Then
                Return False
            End If

            '\TR
            Dim ObjCust As New CustomSMESS.ClsTransaksiStockInOut
            Dim rd As DataGridViewRow
            Dim TblDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
            TblDummy = Me.TblItemCustomInOut.Copy
            'If TblDummy.Rows.Count > 0 Then
            For ij As Integer = 0 To Me.dgvItem.RowCount - 1
                rd = Me.dgvItem.Rows(ij)
                With rd
                    Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    Dim PrimaryPOSO As Double = .Cells(Me.ColPrimaryPO.Index).Value
                    Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                    Dim Batch As String = .Cells(Me.ColLotNumber.Index).Value
                    Dim QtyDef As Decimal = .Cells(Me.ColTotalQty.Index).Value
                    Dim QtyMulti As Decimal = .Cells(Me.ColTotalQtyMulti.Index).Value
                    Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                    If ObjCust.CompareQtySMESS_Custom(PrimaryProduct, PrimaryPOSO, PrimaryLocation, Batch, QtyDef, QtyMulti, TblDummy) = False Then
                        If String.IsNullOrEmpty(Batch) = False Then
                            ProdCode &= " batch: " & Batch
                        End If
                        MsgBox("Inconsisten data, please cek product: " & ProdCode & "between main form and custom form.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        rd.Selected = True
                        Return False
                    End If
                End With
            Next
            'End If

            'cek batasan custom
            If ObjAOPCustom.ApakahDataBenar(Me.dgvItem, Me.ColPrimaryCodeNo.Index, Me.ColPriceByMultiUnit.Index, Me.ColPrimaryUnitsMulti.Index, _
            Me.ColPrimaryUnit.Index) = False Then
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
            ListPrimaryProduct = Nothing
            ObjAOPCustom = Nothing
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
                            If (QtyRef * CDec(rowDGV.Cells(Me.ColAtAmount.Index).Value)) / CDec(rowDGV.Cells(Me.ColTotalQty.Index).Value) < (1 / 10 ^ DefPrecIDR) Then
                                Me.TabControl1.SelectedIndex = 1
                                Me.dgvItem.ClearSelection()
                                rowDGV.Selected = True

                                MsgBox("Invalid at cost value on row: " + (rowDGV.Index + 1).ToString + vbCrLf + "Minimum at cost value is IDR " + _
                                Me.ObjCD.DigitGrouping(((1 / 10 ^ DefPrecIDR) * CDec(rowDGV.Cells(Me.ColTotalQty.Index).Value)) / QtyRef, DefPrecIDR), MsgBoxStyle.Information, "Information")

                                Return False
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

    ''@rosadinnn
    ''@MUaT
    Private Sub Hitung_Total_Multi()
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Decimal = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Total_Multi_MiddlePoint()
                Exit Sub
            End If
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
                    Dim DefaultPrecisionRow As Double = 1
                    Dim SubTotal1 As Decimal = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    'Dim JmlReturn As Decimal = Me.GetJmlDGVPR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'row.Cells(Me.ColPrimaryGRN.Index).Value, -1) -> 2019-06-10: pada quick jml return = 0

                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        SubTotal1 = ObjCD.DigitGrouping(Math.Round((CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value)) * _
                        CDec(row.Cells(Me.ColNettPrice.Index).Value), CInt(DefaultPrecisionRow), MidpointRounding.AwayFromZero) * exchangRate, DefaultPrecisionRow)
                    Else
                        SubTotal1 = ObjCD.DigitGrouping(Math.Round((CDec(row.Cells(Me.ColTotalQty.Index).Value)) * _
                        CDec(row.Cells(Me.ColNettPrice.Index).Value), CInt(DefaultPrecisionRow), MidpointRounding.AwayFromZero) * exchangRate, DefaultPrecisionRow)
                    End If

                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision) + CDec(row.Cells(Me.ColItemAdjustment.Index).Value)
                Next
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(SubTotal, DefaultPrecision)
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.DGVdp.Rows
                    Me.txtDownPayment.Text = CDbl(Me.txtDownPayment.Text) + CDbl(row.Cells(3).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.txtDownPayment.Text, DefaultPrecision)
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(SubTotal - CDbl(Me.txtDownPayment.Text), DefaultPrecision)
                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.txtPrimaryCurrency.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                            CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
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
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                        End If
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
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
        Dim SubTotal As Double = 0
        Dim exchangRate As Decimal = 0

        Try
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
                    Dim DefaultPrecisionRow As Double = 1
                    Dim SubTotal1 As Decimal = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    'Dim JmlReturn As Decimal = Me.GetJmlDGVPR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'row.Cells(Me.ColPrimaryGRN.Index).Value, -1) -> 2019-06-10: pada quick jml return = 0

                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        SubTotal1 = ObjCD.DigitGrouping(Math.Round((CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value)) * _
                        CDec(row.Cells(Me.ColNettPrice.Index).Value), CInt(DefaultPrecisionRow), MidpointRounding.AwayFromZero) * exchangRate, DefaultPrecisionRow)
                    Else
                        SubTotal1 = ObjCD.DigitGrouping(Math.Round((CDec(row.Cells(Me.ColTotalQty.Index).Value)) * _
                        CDec(row.Cells(Me.ColNettPrice.Index).Value), CInt(DefaultPrecisionRow), MidpointRounding.AwayFromZero) * exchangRate, DefaultPrecisionRow)
                    End If

                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision) + CDec(row.Cells(Me.ColItemAdjustment.Index).Value)
                Next
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(SubTotal, DefaultPrecision)
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.DGVdp.Rows
                    Me.txtDownPayment.Text = CDbl(Me.txtDownPayment.Text) + CDbl(row.Cells(3).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.txtDownPayment.Text, DefaultPrecision)
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(SubTotal - CDbl(Me.txtDownPayment.Text), DefaultPrecision)
                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.txtPrimaryCurrency.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                            CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
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
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                        End If
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                    End If
                End If

                '2025 Core Tax Tahap 2.2
                Me.UbahDPPNilaiLain()

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
    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================

    'Ingat jika GRN sudah masuk maka semua item dalam GRN tersebut harus di-AOP
    Private Sub dtpDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.GotFocus
        Me.dtpDate.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Dim TBL_PO As New Inventory_Module.DseTransaksiPO.orderbalance_report_view0DataTable
    Dim TBL_PO2 As New DseTransaksiAOP.tbltrans_aop_item_view1DataTable
    Dim TBL_PO3 As New Inventory_Module.DseTransaksiPO.orderbalance_report_view0DataTable
    Dim TBL_PO4 As New DseTransaksiAOP.tbltrans_aop_item_view1DataTable

    Private Sub dtpDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.LostFocus
        Me.dtpDate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
    End Sub

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2, ClsCekData.EnumTipeRounding.Middlepoint)
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

    'Private Sub FrmTransaksiAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '    Try
    '        Me.baru = True
    '        Me.NonAktifkanDetail()

    '        Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '        Me.dgvExchageRates.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '        Me.DGVdp.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '        Me.dgvDPDummy.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '        Me.dgvCost.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '        Me.DGVitemPR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '        Me.dgvPR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

    '        'Me.CboFakturPajak.SelectedIndex = 0
    '        Me.cboCalculateBy.SelectedIndex = 0
    '        Me.ColATCost.Visible = True
    '        Me.ColShowDetail.Visible = True
    '        Me.SplitContainer4.Panel2Collapsed = True
    '        Me.txtPrimaryPO.Text = -1
    '        Me.cboPPNDipungut.SelectedIndex = 0
    '        Me.ColAtAmount.Visible = False
    '        ' Me.SetTampilan()

    '        Me.cboPurchaseOrder.Width = Me.SplitContainer5.Panel1.Width - 5 - Me.cboPurchaseOrder.Left

    '        ''@rosadinnn
    '        ''@MUaT
    '        'Me.TxtCodeNo.Width = Me.cboPurchaseOrder.Width - Me.BtnFindCodeNo.Width - 5

    '        Me.SplitContainer6.Width = Me.cboPurchaseOrder.Width

    '        ''@rosadinnn
    '        ''@MUaT
    '        'Me.BtnFindCodeNo.Left = Me.SplitContainer5.Panel1.Width - Me.BtnFindCodeNo.Width - 5

    '        Me.cboCodeNo.Width = Me.SplitContainer6.Panel1.Width - 5 - Me.cboCodeNo.Left
    '        Me.TxtOrderBalance.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.TxtOrderBalance.Left

    '        ''@rosadinnn
    '        ''@MUaT
    '        'Me.txtDesc.Width = Me.cboPurchaseOrder.Width

    '        Me.SplitContainer3.TabStop = False
    '        Me.SplitContainer5.TabStop = False
    '        Me.SplitContainer3.Width = Me.SplitContainer5.Panel1.Width
    '        Me.btnFindLocation.Left = Me.BtnFindCodeNo.Left
    '        Me.TxtLocation.Width = Me.TxtCodeNo.Width
    '        Me.txtBatchNo.Width = Me.SplitContainer3.Panel1.Width - 5 - Me.txtBatchNo.Left
    '        Me.DtpExpired.Width = Me.SplitContainer3.Panel2.Width - 5 - Me.DtpExpired.Left

    '        Me.txtAmount.Width = Me.SplitContainer5.Panel2.Width - 5 - Me.txtAmount.Left
    '        Me.btnFindCurrency2.Left = Me.SplitContainer5.Panel2.Width - Me.btnFindCurrency2.Width - 5
    '        Me.TxtCurrencyCode1.Width = Me.txtAmount.Width - 5 - Me.btnFindCurrency2.Width
    '        Me.TxtDiscount.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.TxtDiscount.Width
    '        Me.Label35.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.TxtDiscount.Width - Me.Label35.Width
    '        Me.btnInput.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.btnInput.Width
    '        Me.TxtRemarkItem.Width = Me.btnInput.Left - 5 - Me.TxtRemarkItem.Left
    '        Me.TxtUnitPrice.Width = Me.Label35.Left - 5 - Me.TxtUnitPrice.Left
    '        Me.LblUnitsMulti.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.LblUnitsMulti.Width
    '        Me.TxtQtyMulti.Left = Me.TxtDiscount.Left
    '        Me.Label37.Left = Me.Label35.Left
    '        Me.txtNettPrice.Left = Me.TxtUnitPrice.Left
    '        Me.txtNettPrice.Width = Me.TxtUnitPrice.Width
    '        Me.BtnFindCodeNo.Left = Me.SplitContainer5.Panel1.Width - 5 - Me.BtnFindCodeNo.Width
    '        Me.TxtCodeNo.Width = Me.BtnFindCodeNo.Left - 5 - Me.TxtCodeNo.Left
    '        Me.txtDesc.Width = Me.SplitContainer5.Panel1.Width - 5 - Me.txtDesc.Left
    '        Me.btnFindLocation.Left = Me.BtnFindCodeNo.Left
    '        Me.btnFindCurrency2.Left = Me.BtnFindCodeNo.Left
    '        Me.btnFindCurrency2.Width = Me.BtnFindCodeNo.Width
    '        Me.TxtCurrencyCode1.Width = Me.TxtCodeNo.Width
    '        Me.CboUnitMulti.Left = Me.btnInput.Left + Me.btnInput.Width - Me.CboUnitMulti.Width

    '        Me.dtpDate.Left = Me.SplitContainer7.Panel1.Width - Me.dtpDate.Width
    '        Me.Label3.Left = Me.dtpDate.Left - Me.Label3.Width - 5
    '        Me.TxtAOP.Width = Me.dtpDate.Left - Me.TxtAOP.Left - 5
    '        Me.TxtSupplierName.Width = Me.dtpDate.Left + Me.dtpDate.Width - Me.TxtSupplierName.Left
    '        Me.llbAddress.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.llbAddress.Width
    '        Me.cboSupplierAddress.Width = Me.llbAddress.Left - 5 - Me.cboSupplierAddress.Left
    '        Me.txtSupplierInvoiceNo.Width = Me.dtpDate.Left + Me.dtpDate.Width - Me.txtSupplierInvoiceNo.Left
    '        Me.TxtNPWPNo.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.TxtNPWPNo.Width
    '        Me.Label13.Left = Me.TxtNPWPNo.Left - 5 - Me.Label13.Width
    '        Me.TxtTaxInvoiceNumber.Width = Me.Label13.Left - 5 - Me.TxtTaxInvoiceNumber.Left
    '        Me.cboCalculateBy.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.cboCalculateBy.Width
    '        Me.Label20.Left = Me.cboCalculateBy.Left - 5 - Me.Label20.Width
    '        Me.txtCurrencyCode.Anchor = Me.TxtTaxInvoiceNumber.Anchor
    '        Me.btnFindCurrency.Anchor = Me.cboCalculateBy.Anchor
    '        Me.btnFindCurrency.Left = Me.TxtTaxInvoiceNumber.Left + Me.TxtTaxInvoiceNumber.Width - Me.btnFindCurrency.Width
    '        Me.txtCurrencyCode.Width = Me.btnFindCurrency.Left - 5 - Me.txtCurrencyCode.Left
    '        Me.cboAccount.Left = Me.txtCurrencyCode.Left
    '        Me.cboPrimaryAccount.Visible = False
    '        Me.cboAccount.Width = Me.dtpDate.Left + Me.dtpDate.Width - Me.cboAccount.Left

    '        Me.txtSubTotalBeforeDP.Width = Me.SplitContainer7.Panel2.Width - Me.txtSubTotalBeforeDP.Left - 5
    '        Me.txtSubTotalAfterDP.Width = Me.txtSubTotalBeforeDP.Width
    '        Me.txtTaxExchangeRate.Width = 108
    '        Me.txtTaxExchangeRate.Left = Me.txtSubTotalBeforeDP.Left + Me.txtSubTotalBeforeDP.Width - Me.txtTaxExchangeRate.Width
    '        Me.txtTaxAmount.Width = 209
    '        Me.TxtDutyStamp.Width = Me.txtSubTotalBeforeDP.Left + Me.txtSubTotalBeforeDP.Width - Me.TxtDutyStamp.Left
    '        Me.TxtTotal.Width = Me.TxtDutyStamp.Width
    '        Me.Label7.Left = Me.txtSubTotalBeforeDP.Left + Me.txtSubTotalBeforeDP.Width - Me.Label7.Width
    '        Me.TxtPaymentsTerms.Width = Me.Label7.Left - 5 - Me.TxtPaymentsTerms.Left
    '        Me.TxtRemark.Width = Me.TxtTotal.Width
    '        Me.btnFindLocation.Left = Me.BtnFindCodeNo.Left
    '        Me.btnFindLocation.Width = Me.BtnFindCodeNo.Width
    '        Me.TxtLocation.Width = Me.TxtCodeNo.Width
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
    Private Sub FrmTransaksiAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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

            Me.LinkLabel1.Parent = Me.TabPage1
            Me.LinkLabel1.BringToFront()
            Me.ObjAOP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP
            Me.ColATCost.Visible = True
            Me.ColShowDetail.Visible = True
            Me.SplitContainer1.Panel2Collapsed = True

            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.SplitContainer3.TabStop = False
            Me.SplitContainer4.TabStop = False
            Me.SplitContainer5.TabStop = False
            Me.SplitContainer6.TabStop = False
            Me.SplitContainer7.TabStop = False

            Me.SplitContainer3.Parent = Me.SplitContainer5.Panel1
            Me.btnFindCurrency2.Anchor = Me.BtnFindCodeNo.Anchor
            Me.btnFindCurrency2.Left = Me.BtnFindCodeNo.Left
            Me.btnFindCurrency2.Width = Me.BtnFindCodeNo.Width
            Me.btnFindLocation.Anchor = Me.BtnFindCodeNo.Anchor
            Me.btnFindLocation.Left = Me.BtnFindCodeNo.Left
            Me.btnFindLocation.Width = Me.BtnFindCodeNo.Width
            Me.txtPrimaryAccount.Visible = False

            'Me.CboFakturPajak.SelectedIndex = 0
            Me.cboCalculateBy.SelectedIndex = 0
            Me.txtPrimaryPO.Text = -1
            Me.cboPPNDipungut.SelectedIndex = 0
            Me.ColAtAmount.Visible = False
            'Me.SetTampilan()
            Me.SetLocation()
            Me.txtPrimaryAccount.Visible = False
            Me.txtPrimaryAccount.BringToFront()

            Me.KosongkanField()
            Me.SplitContainer3.Width = Me.SplitContainer5.Panel1.Width - Me.SplitContainer3.Left

            Me.DtpInvDate.Left = Me.DtpInvDate.Parent.Width - 5 - Me.DtpInvDate.Width
            Me.Label14.Left = Me.DtpInvDate.Left - 5 - Me.Label14.Width
            Me.txtSupplierInvoiceNo.Width = Me.Label14.Left - 5 - Me.txtSupplierInvoiceNo.Left
            Me.ShowIcon = False
            Me.btnUP.BackgroundImage = My.Resources.arrow_up
            Me.btnDown.BackgroundImage = My.Resources.arrow_down
            Me.Atur1()

            'Me.txtPriceByMultiUnit.Visible = True
            'Me.ColPriceByMultiUnit_St.Visible = True
            Me.ColPriceByMultiUnit.Visible = False

            Me.btnInput.TabStop = True

            Me.dgvItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvItem.ColumnHeadersHeight = 25
            Me.dgvCost.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvCost.ColumnHeadersHeight = 25
            Me.dgvExchageRates.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvExchageRates.ColumnHeadersHeight = 25

            '###TR
            Me.llbAdjustDPP.Visible = False ' disuruh pa iwan terkait error saat klik +/-


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

    '===================================================================================================
    'Code Erna
    '===================================================================================================

    Private Sub IsiCBOArAkunByTipePerkiraan1()
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
            Me.tblAccount.Rows.Clear()

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccount, Me.txtPrimaryCurrency.Text)
            Me.tblAccount.DefaultView.RowFilter = Nothing
            Me.cboAccount.DataSource = Me.tblAccount.DefaultView
            Me.cboAccount.DisplayMember = "Account Code"
            Me.cboAccount.SelectedIndex = -1

            If Me.cboAccount.Items.Count > 0 Then
                Me.cboAccount.SelectedIndex = Me.cboAccount.Items.Count - 1
                Me.cboAccount.SelectedIndex = 0
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOCurrencyTax()
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

            Me.cboCurrTax.SelectedIndex = -1

            If Me.cboCurrTax.Items.Count > 0 Then
                Me.cboCurrTax.SelectedIndex = Me.cboCurrTax.Items.Count - 1
                Me.cboCurrTax.SelectedIndex = 0

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '===================================================================================================



    Private Sub SetTaxRate(Optional ByVal vSetNewTaxRate As Boolean = True)
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

    Private Sub UpdatePaymentTerm()
        Dim ListPrimaryPO As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim PrimaryPO As Double = -1
        Dim PaymentDue As Double = -1
        Dim Dummy As Double = 0

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

            If CDbl(Me.TxtPaymentsTerms.Text) <> PaymentDue Then
                If MsgBox("Update payment terms?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                    Me.TxtPaymentsTerms.Text = PaymentDue
                    Me.TabControl1.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryPO = Nothing
        End Try
    End Sub


    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim dummyPrimaryCurrency As Double = -1
        Dim DefaulPres As Double = 0

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'format dp
            For Each row In Me.DGVdp.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
            Next

            'yang diformat dgv-nya berkaitan sec unit:
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColSecQty.Index).Value = "-"
                row.Cells(Me.ColSecUnit.Index).Value = "-"
            Next
            Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False
            Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False

            'untuk exchange rate 
            For Each row In Me.dgvExchageRates.Rows
                'digit grouping exchange rate menurut mata uangnya
                row.Cells(Me.ColExchangeRate.Index).Value = _
                Me.ObjCD.DigitGrouping(row.Cells(Me.ColExchangeRate.Index).Value)
                'digit grouping cost amount menurut rupiah
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Function ApakahPerluBuatPOBaru() As Boolean
        Dim NoPOLama As String = ""

        Try
            If Me.baru = True Then
                For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColPONO.Index).Value.ToString) = True Then
                        Return True
                    End If
                Next

                Return False
            Else
                NoPOLama = Me.ObjPO.GetNoTransaksiPO_Quick(Me.TxtPrimaryTransaksiAOP.Text)
                For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColPONO.Index).Value.ToString) = True Or _
                    rowDGV.Cells(Me.ColPONO.Index).Value.ToString = NoPOLama Then
                        Return True
                    End If
                Next

                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim FrmPO As New FrmTransaksiAOP_Quick_DetailNumber
        Dim Lanjut As Boolean = True

        Try
            FrmPO.PrimaryAOPQuickAcuan = Me.TxtPrimaryTransaksiAOP.Text
            FrmPO.BackColorAll = Me.BackColorAll
            If Me.baru = False Then
                FrmPO.NoPO_Lama = Me.ObjPO.GetNoTransaksiPO_Quick(Me.TxtPrimaryTransaksiAOP.Text)
                FrmPO.TxtPONumber.Text = FrmPO.NoPO_Lama
            Else
                FrmPO.NoPO_Lama = Me.BufferPONumber
            End If

            If Me.ApakahDataBenar() = True Then
                If Me.ApakahPerluBuatPOBaru = True Then
                    If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) = False Then
                        If FrmPO.ShowDialog = Windows.Forms.DialogResult.OK Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    Else
                        Lanjut = True
                    End If
                Else
                    Lanjut = True
                End If

                If Lanjut = True Then
                    Me.BufferPONumber = FrmPO.TxtPONumber.Text
                    Me.KosongkanItem()
                    Me.SalinKeObjek()

                    If baru = True Then
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor

                            If Me.ObjAOP.SimpanBaru(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                                                      Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me, FrmPO.TxtPONumber.Text, Me.TblCustomDummy, Me.TblItemCustomInOut) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            End If
                        End If
                        'End If
                    Else
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor

                            If Me.ObjAOP.SimpanPerubahan(Me.dgvItem, Me.DGVdp, _
                                                        Me.DGVitemPR, Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me, FrmPO.TxtPONumber.Text, Me.TblCustomDummy, Me.TblItemCustomInOut) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            FrmPO = Nothing
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 Then
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
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.ObjAOP.CariDetail(Me.TxtPrimaryTransaksiAOP.Text, Me.dgvItem, _
                        Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut)
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

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiAOP

        Try
            ij.BackColorAll = Me.BackColorAll
            ij.TipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP
            ij.ForeignSupplier = 0
            ij.QuickPurchase = True

            ij.ShowDialog()

            If Not ij.PrimaryTransaksi = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjAOP.CariDetail(ij.PrimaryTransaksi, _
                        Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut) _
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
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim mPrimaryCurrency As Double = -1
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtAOP.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOP.Text <> "" Then
                If Me.ObjAOP.ApakahBolehEdit(Me.ObjAOP.Primary, Me.UserCode, True) = True Then
                    If Me.ObjAOP.ApakahQuickPurchase(Me.TxtPrimaryTransaksiAOP.Text) = False Then
                        MsgBox("This aop is not quick purchase. Use approval of purchase form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.baru = False
                    ' Me.SedangRevisi = False
                    Me.ObjAOP.CariDetail(Me.ObjAOP.Primary, Me.dgvItem, Me.DGVdp, _
                      Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                    'primary currency item, ambil dari bari pertama
                    rowDGV = Me.dgvItem.Rows(0)
                    mPrimaryCurrency = rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value
                    Me.TxtPrimaryCurrency1.Text = mPrimaryCurrency
                    Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim PrimaryStockIn As Double = -1

        Try
            If Me.TxtAOP.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOP.Text <> "" Then
                If Me.ObjAOP.ApakahBolehHapus(Me.ObjAOP.Primary, True, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    If Me.ObjAOP.ApakahQuickPurchase(Me.TxtPrimaryTransaksiAOP.Text) = False Then
                        MsgBox("This aop is not quick purchase. Use approval of purchase form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    'PrimaryStockIn = Me.objGRN.GetPrimaryStockIn(Me.dgvItem.Rows(0).Cells(Me.ColPrimaryGRN.Index).Value)
                    'If Me.ObjStockInItem.ApakahHapusMengakibatkanMinus(PrimaryStockIn, _
                    'Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, True) = True Then
                    'Exit Sub
                    'End If

                    Dim NoInvLama As String = Me.TxtAOP.Text.Substring(0, 19)
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjAOP.Hapus(Me.TxtPrimaryTransaksiAOP.Text, Me.UserCode, True, Me) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            'jika masih ada invoce yang revisi selanjutnya maka tampilkan
                            Dim PrimaryInvBaru As Double = Me.ObjAOP.GetPrimaryByNoTransaksiLIKE(NoInvLama)
                            If PrimaryInvBaru <> -1 Then
                                If MsgBox("Load previous revision?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Me.ObjAOP.CariDetail(PrimaryInvBaru, Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                                    Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut)
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
        Dim FrmPO As New FrmTransaksiAOP_Quick_DetailNumber

        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    FrmPO.PrimaryAOPQuickAcuan = Me.TxtPrimaryTransaksiAOP.Text
                    FrmPO.NoPO_Lama = Me.ObjPO.GetNoTransaksiPO_Quick(Me.TxtPrimaryTransaksiAOP.Text)
                    FrmPO.BackColorAll = Me.BackColorAll

                    If FrmPO.ShowDialog = Windows.Forms.DialogResult.OK Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinKeObjek()
                            If baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.ObjAOP.SimpanBaru(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                                Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me, FrmPO.TxtPONumber.Text, Me.TblCustomDummy, Me.TblItemCustomInOut) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjAOP.SimpanPerubahan(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, _
                            Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me, FrmPO.TxtPONumber.Text, Me.TblCustomDummy, Me.TblItemCustomInOut) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmPO.Dispose()
        End Try
    End Sub

    'Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
    '    Try
    '        If e.ColumnIndex = Me.ColShowDetail.Index Then
    '            If e.RowIndex <> -1 Then
    '                Me.ShowDetailCalculation(Me.dgvItem.Item(Me.ColCodeNo.Index, _
    '                e.RowIndex).Value, e.RowIndex)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
        Dim row As DataGridViewRow

        Try
            If e.ColumnIndex = Me.ColShowDetail.Index Then
                If e.RowIndex <> -1 Then
                    Me.ShowDetailCalculation(Me.dgvItem.Item(Me.ColCodeNo.Index, _
                        e.RowIndex).Value, e.RowIndex)
                End If
            ElseIf e.ColumnIndex = Me.ColRatioDesc.Index Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOP_Quick2_AddByCustomProduct") = True Then
                    Exit Sub
                End If

                If e.RowIndex <> -1 Then
                    If Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value Or _
                            Me.TxtRemark.Enabled = False Then
                        Exit Sub
                    End If

                    Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    Dim BufferRatioDesc As String = Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value

                    frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                    frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                    frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    frmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value
                    frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                        row = Me.dgvItem.Rows(e.RowIndex)
                        Me.SetRatioAndQtyDefault(row, True)
                        Me.Hitung_Amount_Multi(row)
                        Me.Hitung_AtCost_Multi()

                        If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) = 0 Then
                            MsgBox("Ratio value restored to previous value, because ratio you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                            Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = BufferRatioDesc
                            Me.SetQtyDefaultOnDgv(e.RowIndex)
                        End If

                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColPriceByMultiUnit_Str.Index And Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOP_Quick2_AddByCustomProduct") = True Then
                        Exit Sub
                    End If

                    Dim FrmRef As New Core_Module.FrmTrans_PriceByMultiUnit
                    FrmRef.UnitMulti = Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value
                    FrmRef.UnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    FrmRef.Buffer_PriceByMultiUnit = Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value

                    If FrmRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = FrmRef.cboPriceByMultiUnit.SelectedIndex
                        If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = "1" Then
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value.ToString
                        Else
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString
                        End If
                        'hitung ulang row
                        Me.Hitung_Amount_Multi(Me.dgvItem.Rows(e.RowIndex))
                        'hitung total
                        Me.Hitung_Total_Multi()
                    End If

                    FrmRef = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.IsikanCBOPO()
            Me.FormatDGV2()
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.Hitung_Total_Multi()
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

    Private Function GetJmlDGVPR(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryGRN As Double, _
    ByVal vSecQty As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvPR.Rows
                If IsNumeric(row.Cells(Me.ColSecQty_PR.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN And _
                    row.Cells(Me.ColSecQty_PR.Index).Value = vSecQty Then
                        Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo_PR.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColPrimaryGRN_PR.Index).Value = vPrimaryGRN Then
                        Jml = Jml + CDec(row.Cells(Me.ColTotalQty_PR.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Hitung_Amount_Default(ByVal row As DataGridViewRow)
        Dim DefaultPrecisionRow As Integer
        Try
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value * _
            (100 - row.Cells(Me.ColDiscount.Index).Value) / 100, DefaultPrecisionRow))

            'hitung amount
            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecisionRow)
            Else
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecisionRow)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub Hitung_Amount_Multi(ByVal row As DataGridViewRow)
        Dim DefaultPrecisionRow As Integer
        Try
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value - _
            row.Cells(Me.ColDiscAmount.Index).Value, DefaultPrecisionRow))

            'hitung amount
            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                row.Cells(Me.ColNettPrice.Index).Value + CDec(row.Cells(Me.ColItemAdjustment.Index).Value), DefaultPrecisionRow)
            Else
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value * _
                row.Cells(Me.ColNettPrice.Index).Value + CDec(row.Cells(Me.ColItemAdjustment.Index).Value), DefaultPrecisionRow)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Function GetSubTotal_TanpaReturn_Default() As Double
    '    Dim row As DataGridViewRow
    '    Dim SubTotal As Double = 0
    '    Dim exchangRate As Double = 0

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
    '            Dim DefaultPrecisionRow As Double = 1
    '            Dim SubTotal1 As Double = 0
    '            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)

    '            SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * _
    '                 row.Cells(Me.ColNettPrice.Index).Value * exchangRate, DefaultPrecisionRow)

    '            SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
    '        Next

    '        Return Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetSubTotal_TanpaReturn_Multi() As Double
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Double = 0

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
                Dim DefaultPrecisionRow As Double = 1
                Dim SubTotal1 As Double = 0
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)

                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    SubTotal1 = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                         CDec(row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItemAdjustment.Index).Value)) * exchangRate, DefaultPrecisionRow)
                Else
                    SubTotal1 = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColTotalQty.Index).Value) * _
                    CDec(row.Cells(Me.ColNettPrice.Index).Value) + CDec(row.Cells(Me.ColItemAdjustment.Index).Value)) * exchangRate, DefaultPrecisionRow)
                End If

                SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
            Next

            Return Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetExchangeRate(ByVal vPrimaryTransaksi As Double) As Double
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvExchageRates.Rows
                If row.Cells(5).Value = vPrimaryTransaksi Then
                    Return row.Cells(3).Value
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Hitung_Total_Default()
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Decimal = 0

        Try
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
                    Dim DefaultPrecisionRow As Double = 1
                    Dim SubTotal1 As Decimal = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    'Dim JmlReturn As Decimal = Me.GetJmlDGVPR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'row.Cells(Me.ColPrimaryGRN.Index).Value, -1) -> 2019-06-10: jml return pada quick = 0

                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        SubTotal1 = ObjCD.DigitGrouping(Math.Round((CDec(row.Cells(Me.ColTotalQty.Index).Value)) * _
                        CDec(row.Cells(Me.ColNettPrice.Index).Value), CInt(DefaultPrecisionRow), MidpointRounding.AwayFromZero) * exchangRate, DefaultPrecisionRow)
                    Else
                        SubTotal1 = ObjCD.DigitGrouping(Math.Round((CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value)) * _
                        CDec(row.Cells(Me.ColNettPrice.Index).Value), CInt(DefaultPrecisionRow), MidpointRounding.AwayFromZero) * exchangRate, DefaultPrecisionRow)
                    End If

                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
                Next
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(SubTotal, DefaultPrecision)
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.DGVdp.Rows
                    Me.txtDownPayment.Text = CDbl(Me.txtDownPayment.Text) + CDbl(row.Cells(3).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.txtDownPayment.Text, DefaultPrecision)
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(SubTotal - CDbl(Me.txtDownPayment.Text), DefaultPrecision)
                'hitung pajak
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.txtPrimaryCurrency.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDbl(Me.txtTax.Text) * _
                            CDbl(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefaultPrecision)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDbl(Me.txtTaxExchangeRate.Text) _
                            * CDbl(Me.txtTax.Text) * CDbl(Me.txtSubTotalAfterDP.Text) / 100, _
                            0, ClsCekData.EnumTipeRounding.Down), DefaultPrecision)
                        End If
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefaultPrecision)
                    End If
                Else
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefaultPrecision)
                End If

                'hitung total = sub total + taxj 
                If Me.txtPrimaryCurrency.Text = "1" Then
                    'kalo IDR, duty stamp diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(CDbl(Me.txtSubTotalAfterDP.Text) + _
                    CDbl(Me.txtTaxAmount.Text) + CDbl(Me.TxtDutyStamp.Text), DefaultPrecision)
                Else
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(CDbl(Me.txtSubTotalAfterDP.Text), _
                    DefaultPrecision)
                End If
            End If
            'hitung ulang unit price
            Me.Hitung_AtCost_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
                    FilterDP()
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
                        FilterDP()
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
                Me.cboPrimaryCurrTax.Enabled = False
            Else
                Me.cboCurrTax.Enabled = True
                Me.cboPrimaryCurrTax.Enabled = True
            End If

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
                    FilterDP()
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
                        FilterDP()
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
            If Me.txtPrimaryCurrency.Text <= 1 Then
                Me.cboCurrTax.Enabled = False
                Me.cboPrimaryCurrTax.Enabled = False
            Else
                Me.cboCurrTax.Enabled = True
                Me.cboPrimaryCurrTax.Enabled = True
            End If

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
                Me.SetDefaultPrecision()
                'tampilan exchange rate
                'Me.SetDGVExchangeRate(True)
                'FilterDP()
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.txtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                    'tampilan exchange rate                 
                    Me.cboCalculateBy.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetDefaultPrecision()
                End If
            End If
            ' If Me.txtPrimaryCurrTax.Text = -1 Then
            If Me.txtPrimaryCurrency.Text <= 1 Then
                Me.cboCurrTax.Enabled = False
                Me.cboPrimaryCurrTax.Enabled = False
            Else
                Me.cboCurrTax.Enabled = True
                Me.cboPrimaryCurrTax.Enabled = True
            End If

            Me.IsiCBOCurrencyTax()
            Me.IsiCBOArAkunByTipePerkiraan1()
            '  End If

            Me.SetTaxRate()
            Me.SetCalculateBy()
            Me.SetDGVExchangeRate(True)
            FilterDP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetDefaultPrecision(Optional ByVal vHitungTotal As Boolean = True)
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
            If vHitungTotal = True Then
                Me.Hitung_Total_Multi()
            End If
            'Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.txtSubTotalBeforeDP.Text, DefaultPrecision)
            'Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.txtSubTotalAfterDP.Text, DefaultPrecision)
            'Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
            'Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
            'Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.TxtTotal.Text, DefaultPrecision)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim Dummy As String = ""
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.chkTax.Checked = True Then
                    Me.txtTax.ReadOnly = False
                    Me.txtTax.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    Me.txtTax.TabStop = True
                    Me.txtTaxExchangeRate.ReadOnly = False
                    Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    Me.txtTaxExchangeRate.TabStop = True
                    Me.TxtNPWPNo.ReadOnly = False
                    Me.TxtNPWPNo.TabStop = True
                    Me.TxtNPWPNo.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    Me.TxtTaxInvoiceNumber.ReadOnly = False
                    Me.TxtTaxInvoiceNumber.TabStop = True
                    Me.TxtTaxInvoiceNumber.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    If Me.txtPrimarySupplier.Text <> "-1" And Me.txtPrimarySupplier.Text <> "" And _
                    IsNumeric(Me.txtPrimarySupplier.Text) = True And Me.TxtNPWPNo.Text = "" Then
                        Me.TxtNPWPNo.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                    End If

                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy = "" Then
                        Dummy = "0"
                    End If
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                    'If CDbl(Me.txtTax.Text) = 0 Then
                    '    Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
                    'End If
                    'Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2)
                Else
                    Me.txtTax.ReadOnly = True
                    Me.txtTax.BackColor = Me.txtReadOnly.BackColor
                    Me.txtTax.TabStop = False
                    Me.txtTaxExchangeRate.ReadOnly = True
                    Me.txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
                    Me.txtTaxExchangeRate.TabStop = False
                    Me.TxtNPWPNo.ReadOnly = True
                    Me.TxtNPWPNo.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtNPWPNo.TabStop = False
                    Me.TxtTaxInvoiceNumber.ReadOnly = True
                    Me.TxtTaxInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtTaxInvoiceNumber.TabStop = False
                    Me.TxtNPWPNo.Text = ""
                    Me.TxtTaxInvoiceNumber.Text = ""
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", 2)
                    'If Me.txtPrimaryCurrency.Text <> Core_Module.PrimaryBaseCurrency Then
                    '    Me.txtTaxExchangeRate.ReadOnly = False
                    '    Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                    '    Me.txtTaxExchangeRate.TabStop = True
                    'End If
                End If
            End If
            If Me.TxtRemark.Enabled = True Then
                Me.SetTaxRate()
                Me.Hitung_Total_Multi()
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
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""

        Dim Reset As Boolean = False
        Try
            If e.KeyValue = 13 Then
                If Me.Objsupplier.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                    Me.txtPrimarySupplier.Text = Me.Objsupplier.GetPrimary(Me.txtSupplierCode.Text)
                    Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                    Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.SelectedIndex = 0
                    Me.cboSupplierAddress.Focus()
                    Me.IsikanCBOPO()
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
                        Me.cboSupplierAddress.SelectedIndex = 0
                        Me.cboSupplierAddress.Focus()
                        Me.IsikanCBOPO()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.cboSupplierAddress.Items.Clear()
                        Me.txtSupplierAddIdentifier.Text = ""
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
                        Me.KosongkanItem()
                    End If
                End If
            End If
            Me.TxtNPWPNo.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
            Me.TxtPaymentsTerms.Text = Me.Objsupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)

            '2025 Core Tax Tahap 2.1
            If Me.dtpDate.Value >= Me.ObjSetNPWP.Tglberlaku Then
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
        Dim Dummy As String = ""
        Dim Dummy2 As String = ""
        Dim reset As Boolean = False

        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.Objsupplier.GetJml(Me.txtSupplierCode.Text) = 1 Then
                    Me.txtPrimarySupplier.Text = Me.Objsupplier.GetPrimary(Me.txtSupplierCode.Text)
                    Me.TxtSupplierName.Text = Me.Objsupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                    Me.Objsupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                    Me.cboSupplierAddress.SelectedIndex = 0
                    Me.cboSupplierAddress.Focus()
                    Me.IsikanCBOPO()
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
                        Me.cboSupplierAddress.SelectedIndex = 0
                        Me.cboSupplierAddress.Focus()
                        Me.IsikanCBOPO()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.cboSupplierAddress.Items.Clear()
                        Me.txtSupplierAddIdentifier.Text = ""
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
                        Me.KosongkanItem()
                    End If
                End If
            Else
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.cboSupplierAddress.Items.Clear()
                Me.txtSupplierAddIdentifier.Text = ""
                Me.txtPrimarySupplier.Text = "-1"
                Me.KosongkanItem()
            End If
            Me.TxtNPWPNo.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
            Me.TxtPaymentsTerms.Text = Me.Objsupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)

            '2025 Core Tax Tahap 2.1
            If Me.dtpDate.Value >= Me.ObjSetNPWP.Tglberlaku Then
                Me.TxtNPWPNo.Text = Me.ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
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
                Me.cboSupplierAddress.Focus()
                Me.IsikanCBOPO()
            Else
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.cboSupplierAddress.Items.Clear()
                Me.txtSupplierAddIdentifier.Text = ""
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Focus()
                Me.KosongkanItem()
            End If
            Me.TxtNPWPNo.Text = Me.Objsupplier.GetNPWP(Me.txtPrimarySupplier.Text)
            Me.TxtPaymentsTerms.Text = Me.Objsupplier.GetPaymentTerm(Me.txtPrimarySupplier.Text)

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
            If Me.TxtRemark.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotalBeforeDP.Text) Then
                If CDbl(Me.txtTax.Text) <= 100 Then
                    'hitung pajak
                    If Me.chkTax.Checked = True Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTax.Text * _
                        CDbl(Me.txtSubTotalBeforeDP.Text) / 100, _
                        Me.ObjCurrency.GetDefaultPrecission(1)))
                    Else
                        Me.txtTax.Text = "0"
                        Me.txtTaxAmount.Text = "0"
                    End If
                    'hitung total = sub total + tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotalBeforeDP.Text) _
                    + CDbl(Me.txtTaxAmount.Text), DefaultPrecision))
                End If
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
                        Me.Text = "Quick Purchase: " + Me.txtSupplierCode.Text + " - " + Me.TxtSupplierName.Text
                    Else
                        Me.Text = "Quick Purchase"
                    End If
                Else
                    Me.Text = "Quick Purchase"
                End If
            Else
                Me.Text = "Quick Purchase"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
            If Me.TxtRemark.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryTransaksiAOP.Text) AndAlso CDec(Me.TxtPrimaryTransaksiAOP.Text) > 0) Then
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
                Else
                    ij.cboCustAddCode.Enabled = True
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

    Private Sub SetDGVExchangeRate(Optional ByVal vHitungTotal As Boolean = False)
        Dim rowDGV As DataGridViewRow
        Dim PrimaryCurrency1, PrimaryCurrency2 As Double
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

                Me.SplitContainer1.Panel1Collapsed = False
                Me.SplitContainer2.Panel2Collapsed = False

                If Me.SplitContainer2.Panel2Collapsed = False Then
                    For Each rowDGV In Me.dgvCost.Rows
                        If rowDGV.Cells(Me.ColPrimaryCurrencyCost.Index).Value = PrimaryCurrency1 Then
                            'rowDGV.Cells(Me.ColExcRateCost.Index).Value = 1
                            'Me.Hitung_Cost_Amount(rowDGV)
                        End If
                    Next
                    'hitung ulang AT-cost
                    If Me.TxtRemark.Enabled = True Then
                        Me.Hitung_AtCost_Multi()
                    End If
                End If
            End If

            If Me.txtCurrencyCode.Text = "" Then
                Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate"
            Else
                Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate to " _
                + Me.txtCurrencyCode.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
                            If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimaryGRN.Index).Value Then
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
                            If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimaryPO.Index).Value Then
                                Ketemu = True
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan exchange rate baru 
                            Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(5).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
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
                If Me.txtTaxExchangeRate.Enabled = True And Me.txtTaxExchangeRate.ReadOnly = False Then
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

    Private Sub llbAddDP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddDP.LinkClicked
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

    Private Sub IsiDGVReturn()
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
                If rowdgv.Cells(1).Value = vPR_Number Then
                    Return True
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiDGVReturnItem(Optional ByVal vHitungTotalMulti As Boolean = True)
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
                Me.Hitung_Amount_Multi(row)
            Next
            If vHitungTotalMulti = True Then
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

    Private Sub dgvItemSR_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs)
        Try
            Me.IsiDGVReturnItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutyStamp.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Dim DeffPrecX As Decimal = 0
                DeffPrecX = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, DeffPrecX)
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
                                        Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, _
                                        PrimaryCurrencyAkun), Core_Module.DefPrecExchangeRate)
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
                                    Me.dgvCost.Item(Me.ColExcRateCost.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, _
                                    PrimaryCurrencyAkun), Core_Module.DefPrecExchangeRate)
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
            TotalFOBPrice = (Me.GetSubTotal_TanpaReturn_Multi * CDbl(Me.txtTaxExchangeRate.Text))
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
            DetailCalculation = DetailCalculation + " Total Cost(3) = FOB Price(1) + Other Cost(2)" + vbCrLf
            DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalFOBPrice)) _
            + " + IDR " + Me.ObjCD.DigitGrouping(CStr(TotalOtherCost)) + vbCrLf
            DetailCalculation = DetailCalculation + "                      = IDR " + Me.ObjCD.DigitGrouping(CStr(TotalFOBPrice + TotalOtherCost)) + vbCrLf + vbCrLf

            'lakukan perhitungan unit price
            DetailCalculation = DetailCalculation + "Unit price = Nett Price / SubTotal * Total Cost(3) " + vbCrLf
            For Each row In Me.dgvItem.Rows
                SubTotal += row.Cells(Me.ColAmount.Index).Value
            Next
            If Me.dgvItem.Item(Me.ColTotalQty.Index, rowIndex).Value - Me.GetJmlDGVPR(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1) <> 0 Then
                NettPrice = Me.dgvItem.Item(Me.ColNettPrice.Index, rowIndex).Value
            Else
                NettPrice = 0
            End If

            DetailCalculation = DetailCalculation + "                = " + Currency + " (" + Me.dgvItem.Item(Me.ColAmount.Index, rowIndex).Value.ToString + " / " + _
            Currency + Me.ObjCD.DigitGrouping(SubTotal.ToString) + " * IDR " + Me.ObjCD.DigitGrouping((TotalFOBPrice + TotalOtherCost).ToString) + ")/ " + _
            Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(Me.ColTotalQty.Index, rowIndex).Value - Me.GetJmlDGVPR(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1)) + " " + Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, rowIndex).Value.ToString + vbCrLf
            DetailCalculation = DetailCalculation + "                = IDR " + Me.dgvItem.Item(Me.ColATCost.Index, rowIndex).Value.ToString
            If Me.dgvItem.Item(Me.ColTotalQty.Index, rowIndex).Value - Me.GetJmlDGVPR(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, rowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryGRN.Index, rowIndex).Value, -1) = 0 Then
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

    '    Try
    '        DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
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
    '        SubTotal = Me.GetSubTotal_TanpaReturn_Default

    '        'hitung total harga FOB
    '        'jika rupiah maka langsung perhitungkan rupiah, jika tidak maka gunakan kurs konversi ke
    '        'rupiah
    '        TotalCost = Math.Round((Me.GetSubTotal_TanpaReturn_Default * CDec(Me.txtTaxExchangeRate.Text)) + _
    '        TotalOngkosLain, DefPrec, MidpointRounding.AwayFromZero) ' --> duty stamp dan tax amount tidak diperhitungkan untuk At Cost
    '        'CDbl(Me.TxtDutyStamp.Text) + Me.txtTaxAmount.Text + 

    '        'hitung unit price untuk masing - masing item
    '        For Each row In Me.dgvItem.Rows
    '            If SubTotal = 0 Or TotalCost = 0 Then
    '                row.Cells(Me.ColATCost.Index).Value = 0
    '            Else
    '                If row.Cells(Me.ColTotalQty.Index).Value - Me.GetJmlDGVPR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '                row.Cells(Me.ColPrimaryGRN.Index).Value, -1) <> 0 Then
    '                    If row.Index <> Me.dgvItem.RowCount - 1 Then
    '                        row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
    '                        SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
    '                        row.Cells(Me.ColAtAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColAmount.Index).Value * TotalCost) / _
    '                        SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
    '                    Else
    '                        row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
    '                        SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
    '                        row.Cells(Me.ColAtAmount.Index).Value = TotalCost - Me.GetTotalAtAmount_KecualiBarisAkhir
    '                    End If
    '                Else
    '                    row.Cells(Me.ColATCost.Index).Value = 0
    '                    row.Cells(Me.ColAtAmount.Index).Value = 0
    '                End If
    '            End If
    '        Next
    '        'Me.ColAtAmount.Visible = True
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

            TotalCost = Math.Round((GetSubTotal_TanpaReturn_Multi() * CDec(Me.txtTaxExchangeRate.Text)) + _
            TotalOngkosLain, DefPrec, MidpointRounding.AwayFromZero)

            For Each row In Me.dgvItem.Rows
                'cek minus
                If CDec(row.Cells(Me.ColAtAmount.Index).Value()) < 0 Then
                    MsgBox("Inconsistent @ cost detected: negative value data found. Please contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                TotalAtCost += CDec(row.Cells(Me.ColAtAmount.Index).Value())
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

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
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
            SubTotal = Me.GetSubTotal_TanpaReturn_Multi

            'hitung total harga FOB
            'jika rupiah maka langsung perhitungkan rupiah, jika tidak maka gunakan kurs konversi ke
            'rupiah
            TotalCost = Math.Round((Me.GetSubTotal_TanpaReturn_Multi * CDec(Me.txtTaxExchangeRate.Text)) + _
            TotalOngkosLain, DefPrec, MidpointRounding.AwayFromZero) ' --> duty stamp dan tax amount tidak diperhitungkan untuk At Cost
            'CDbl(Me.TxtDutyStamp.Text) + Me.txtTaxAmount.Text + 

            'hitung unit price untuk masing - masing item
            For Each row In Me.dgvItem.Rows
                If SubTotal = 0 Or TotalCost = 0 Then
                    row.Cells(Me.ColATCost.Index).Value = 0
                Else
                    If row.Cells(Me.ColTotalQtyMulti.Index).Value - Me.GetJmlDGVPR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    row.Cells(Me.ColPrimaryGRN.Index).Value, -1) <> 0 Then
                        If row.Index <> Me.dgvItem.RowCount - 1 Then
                            row.Cells(Me.ColAtAmount.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColAmount.Index).Value * TotalCost) / _
                            SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                            'row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
                            'SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                        Else
                            row.Cells(Me.ColAtAmount.Index).Value = TotalCost - Me.GetTotalAtAmount_KecualiBarisAkhir
                            'row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round((row.Cells(Me.ColNettPrice.Index).Value * TotalCost) / _
                            'SubTotal, DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                        End If
                        'disini agak berbeda dg form quick acop, karena tidak return revisi aop
                        row.Cells(Me.ColATCost.Index).Value = Me.ObjCD.DigitGrouping(Math.Round(CDec(row.Cells(Me.ColAtAmount.Index).Value) / _
                        CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), _
                        DefPrec, MidpointRounding.AwayFromZero), DefPrec)
                    Else
                        row.Cells(Me.ColATCost.Index).Value = 0
                        row.Cells(Me.ColAtAmount.Index).Value = 0
                    End If
                End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.ColAtAmount.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTotalAtAmount_KecualiBarisAkhir() As Decimal
        Dim d As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> Me.dgvItem.RowCount - 1 Then
                    d += CDec(rowDGV.Cells(Me.ColAtAmount.Index).Value)
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

    Private Sub FilterDP()
        Dim rowDGV As DataGridViewRow
        Dim AdaYangDihapus As Boolean = False
        Dim counter As Double = 0
        Try
            While counter < Me.DGVdp.RowCount
                rowDGV = Me.DGVdp.Rows(counter)
                If rowDGV.Cells(Me.ColCurrencyDP.Index).Value.ToString <> Me.txtCurrencyCode.Text Then
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
        e.Cancel = True
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
        Dim TanggalAOP As Date = Today

        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    TanggalAOP = Me.ObjAOP.GetTanggalAOP(Me.TxtPrimaryTransaksiAOP.Text)
                    If TanggalAOP.Month = Me.dtpDate.Value.Month And TanggalAOP.Year = Me.dtpDate.Value.Year Then
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

    Private Sub KosongkanItem()
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.RowCount = 0 Then
                If IsNumeric(Me.txtPrimaryCurrency.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = Me.txtPrimaryCurrency.Text
                Else
                    Me.TxtPrimaryCurrency1.Text = "1"
                End If
            Else
                row = Me.dgvItem.Rows(0)
                Me.TxtPrimaryCurrency1.Text = row.Cells(Me.ColPrimaryCurrency.Index).Value
            End If
            Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)

            Me.TxtUnitPrice.Text = "0"
            Me.LblUnitPrice.Text = "Unit Price"
            Me.TxtDiscount.Text = "0"
            Me.txtNettPrice.Text = "0"
            Me.TxtQtyMulti.Text = "0"
            Me.txtAmount.Text = "0"
            Me.TxtRemarkItem.Text = ""
            Me.TxtCodeNo.Text = ""
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtDesc.Text = ""
            Me.txtBatchNo.Text = ""
            Me.DtpExpired.Value = Today
            Me.txtPrimaryLocation.Text = "-1"
            Me.TxtDiscFromPerct.Text = "1"
            Me.TxtDiscAmount.Text = 0
            If Me.chkPONumber.Checked = False Then
                Me.txtPrimaryPO.Text = "-1"
            End If
            Me.TxtLocation.Text = ""

            Me.SetGboItem()

            If Me.chkPONumber.Checked = False Then
                Me.cboCodeNo.TabStop = False
                Me.cboCodeNo.DataSource = Nothing
                Me.cboCodeNo.Items.Clear()
                Me.cboCodeNo.Visible = False
                Me.cboCodeNo.TabStop = False
                Me.LblOrderBalance.Visible = False
                Me.TxtOrderBalance.Visible = False

                Me.TxtCodeNo.Visible = True
                Me.TxtCodeNo.TabStop = True
                Me.BtnFindCodeNo.Visible = True
                Me.BtnFindCodeNo.TabStop = True
            Else
                Me.cboCodeNo.Visible = True
                Me.cboCodeNo.TabStop = True
                Me.LblOrderBalance.Visible = True
                Me.TxtOrderBalance.Visible = True

                Me.TxtCodeNo.Visible = False
                Me.TxtCodeNo.TabStop = False
                Me.BtnFindCodeNo.Visible = False
                Me.BtnFindCodeNo.TabStop = False
            End If

            Me.cboPriceByMultiUnit.Visible = False
            Me.cboPriceByMultiUnit.TabStop = False
            Me.txtPriceByMultiUnit.Text = "1"
            Me.cboPriceByMultiUnit.Items.Clear()
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

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.SetGboItem()
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                            Me.TxtUnitPrice.Focus()
                        Else
                            Me.TxtCurrencyCode1.Focus()
                        End If
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGboItem()
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                Me.TxtUnitPrice.Focus()
                            Else
                                Me.TxtCurrencyCode1.Focus()
                            End If
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If

            setLocation_Product()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.SetGboItem()

                    If My.Computer.Keyboard.ShiftKeyDown = False Then
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                Me.TxtQtyMulti.Focus()
                            Else
                                Me.TxtCurrencyCode1.Focus()
                            End If
                        End If
                    Else
                        If chkPONumber.Checked = True Then
                            Me.cboPurchaseOrder.Focus()
                        Else
                            Me.btnInput.Focus()
                        End If
                    End If
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGboItem()
                        If My.Computer.Keyboard.ShiftKeyDown = False Then
                            If Me.txtBatchNo.Visible = True Then
                                Me.txtBatchNo.Focus()
                            ElseIf Me.TxtLocation.Visible = True Then
                                Me.TxtLocation.Focus()
                            Else
                                If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                    Me.TxtQtyMulti.Focus()
                                Else
                                    Me.TxtCurrencyCode1.Focus()
                                End If
                            End If
                        Else
                            If chkPONumber.Checked = True Then
                                Me.cboPurchaseOrder.Focus()
                            Else
                                Me.btnInput.Focus()
                            End If
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.KosongkanItem()

                If My.Computer.Keyboard.ShiftKeyDown = False Then
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                            Me.TxtQtyMulti.Focus()
                        Else
                            Me.TxtCurrencyCode1.Focus()
                        End If
                    End If
                Else
                    If chkPONumber.Checked = True Then
                        Me.cboPurchaseOrder.Focus()
                    Else
                        Me.btnInput.Focus()
                    End If
                End If
            End If
            setLocation_Product()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.TxtDiscFromPerct.Text = "1"
                Me.SetGboItem()
                If Me.txtBatchNo.Visible = True Then
                    Me.txtBatchNo.Focus()
                ElseIf Me.TxtLocation.Visible = True Then
                    Me.TxtLocation.Focus()
                Else
                    If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtDiscFromPerct.Text = "1"
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
            setLocation_Product()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.TxtPrimaryCodeNo.Text)

        End Try
    End Sub

    Private Sub TxtLocation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtLocation.ReadOnly = False And Me.TxtLocation.Enabled = True Then
                TxtLocation.BackColor = BackColorAll
            Else
                TxtLocation.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
            End If
        End If
    End Sub

    Private Sub TxtLocation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocation.KeyDown
        Dim i As New FrmCariMasterStorageArea
        Dim Obj As New clsMasterStorageArea

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    Me.txtPrimaryLocation.Text = Obj.GetPrimary(Me.TxtLocation.Text)
                    If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Focus()
                    End If
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocation.Text = Obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                            Me.TxtUnitPrice.Focus()
                        Else
                            Me.TxtCurrencyCode1.Focus()
                        End If
                    Else
                        Me.TxtLocation.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        Me.TxtLocation.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtLocation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.Leave
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            If Me.TxtLocation.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    Me.txtPrimaryLocation.Text = obj.GetPrimary(Me.TxtLocation.Text)
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        If Me.DtpExpired.Visible = True Then
                            Me.DtpExpired.Focus()
                        ElseIf Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        Else
                            If Me.chkPONumber.Checked = False Then
                                Me.TxtCodeNo.Focus()
                            Else
                                Me.cboCodeNo.Focus()
                            End If
                        End If
                    Else
                        If Me.TxtPrimaryCurrency1.Text = "" Or Me.TxtPrimaryCurrency1.Text = "-1" Then
                            Me.TxtCurrencyCode1.Focus()
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    End If
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)

                        If My.Computer.Keyboard.ShiftKeyDown = True Then
                            If Me.DtpExpired.Visible = True Then
                                Me.DtpExpired.Focus()
                            ElseIf Me.txtBatchNo.Visible = True Then
                                Me.txtBatchNo.Focus()
                            Else
                                If Me.chkPONumber.Checked = False Then
                                    Me.TxtCodeNo.Focus()
                                Else
                                    Me.cboCodeNo.Focus()
                                End If
                            End If
                        Else
                            If Me.TxtPrimaryCurrency1.Text = "" Or Me.TxtPrimaryCurrency1.Text = "-1" Then
                                Me.TxtCurrencyCode1.Focus()
                            Else
                                Me.TxtQtyMulti.Focus()
                            End If
                        End If
                    Else
                        Me.TxtLocation.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        Me.TxtLocation.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryLocation.Text = "-1"
                If My.Computer.Keyboard.ShiftKeyDown = True Then
                    If Me.DtpExpired.Visible = True Then
                        Me.DtpExpired.Focus()
                    ElseIf Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    Else
                        If Me.chkPONumber.Checked = False Then
                            Me.TxtCodeNo.Focus()
                        Else
                            Me.cboCodeNo.Focus()
                        End If
                    End If
                Else
                    If Me.TxtPrimaryCurrency1.Text = "" Or Me.TxtPrimaryCurrency1.Text = "-1" Then
                        Me.TxtCurrencyCode1.Focus()
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocation.Click
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocation.Text <> "" Then
                i.txtCari.Text = Me.TxtLocation.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                    Me.TxtUnitPrice.Focus()
                Else
                    Me.TxtCurrencyCode1.Focus()
                End If
            Else
                Me.TxtLocation.Clear()
                Me.TxtLocation.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.TxtPrimaryUnitsDefault)

            If Me.chkLocationControlled.Checked = False Then
                Me.txtPrimaryLocation.Text = "-1"
                Me.TxtLocation.Text = ""
                Me.TxtLocation.ReadOnly = True
                Me.lblLocation.Visible = True
                Me.TxtLocation.Visible = False
                Me.btnFindLocation.Visible = False
                Me.btnFindLocation.TabStop = False
                Me.TxtLocation.TabStop = False
            Else
                Me.TxtLocation.ReadOnly = False
                Me.TxtLocation.TabStop = True
                Me.lblLocation.Visible = False
                Me.TxtLocation.Visible = True
                Me.btnFindLocation.Visible = True
                Me.btnFindLocation.TabStop = True
            End If
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.lblExpiryDate.Visible = False
                Me.DtpExpired.Visible = True
                Me.DtpExpired.Value = Today
            Else
                Me.lblExpiryDate.Visible = True
                Me.DtpExpired.Visible = False
                Me.DtpExpired.Value = #1/1/1950#
            End If
            If Me.chkBatchNoControlled.Checked = True Then
                Me.txtBatchNo.ReadOnly = False
                Me.txtBatchNo.TabStop = True
                Me.lblBatchNo.Visible = False
                Me.txtBatchNo.Visible = True
            Else
                Me.txtBatchNo.ReadOnly = True
                Me.txtBatchNo.TabStop = False
                Me.lblBatchNo.Visible = True
                Me.txtBatchNo.Visible = False
            End If

            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.LblUnitsMulti.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
            Else
                Me.LblUnitsMulti.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.buffer = Me.TxtUnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e, 9999999999)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            If Me.TxtDiscount.Enabled = True Then
                If CDec(Me.TxtDiscount.Text) = 0 Then
                    Me.TxtDiscount.Text = 0
                End If
                Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtCodeNo.ReadOnly = False And Me.TxtCodeNo.Enabled = True Then
                TxtCodeNo.BackColor = BackColorAll
            Else
                TxtCodeNo.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
            End If
        End If
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtCodeNo.ReadOnly = False And Me.TxtCodeNo.Enabled = True Then
                TxtCodeNo.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
            End If
        End If
    End Sub

    Private Sub TxtCurrencyCode1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.GotFocus
        If Me.TxtCurrencyCode1.ReadOnly = False And Me.TxtCurrencyCode1.Enabled = True Then
            TxtCurrencyCode1.BackColor = BackColorAll
        Else
            TxtCurrencyCode1.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtCurrencyCode1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.LostFocus
        If Me.TxtCurrencyCode1.ReadOnly = False And Me.TxtCurrencyCode1.Enabled = True Then
            TxtCurrencyCode1.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
            TxtUnitPrice.BackColor = BackColorAll
        Else
            TxtUnitPrice.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
            TxtUnitPrice.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.GotFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = BackColorAll
        Else
            TxtDiscount.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.GotFocus
        If Me.TxtQtyMulti.ReadOnly = False And Me.TxtQtyMulti.Enabled = True Then
            TxtQtyMulti.BackColor = BackColorAll
        Else
            TxtQtyMulti.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then

                    ''@rosadinnn
                    ''@MUaT
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                    Me.SetQtyDefault()
                    'Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))

                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.LostFocus
        If Me.TxtQtyMulti.ReadOnly = False And Me.TxtQtyMulti.Enabled = True Then
            TxtQtyMulti.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        If Me.TxtRemarkItem.ReadOnly = False And Me.TxtRemarkItem.Enabled = True Then
            TxtRemarkItem.BackColor = BackColorAll
        Else
            TxtRemarkItem.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        If Me.TxtRemarkItem.ReadOnly = False And Me.TxtRemarkItem.Enabled = True Then
            TxtRemarkItem.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                If Me.TxtDiscFromPerct.Text = "0" Then
                    If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                        Me.TxtUnitPrice.Text = Me.buffer
                        Exit Sub
                    End If
                End If
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                Me.Hitung_Amount2()

                If My.Computer.Keyboard.ShiftKeyDown = False Then
                    If Me.TxtDiscount.Enabled = True Then
                        Me.TxtDiscount.Focus()
                    End If
                Else
                    Me.TxtQtyMulti.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscFromPerct.Text = "1"
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)

                If CDec(Me.TxtUnitPrice.Text) = 0 Then
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Amount2()

        Try
            Dim Decimal_Def As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text), Decimal_Def)

            If Me.TxtDiscFromPerct.Text = "" Then
                Me.TxtDiscFromPerct.Text = "1"
            End If
            If Me.TxtDiscFromPerct.Text = "1" Then
                'hitung disc amount
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtUnitPrice.Text) * CDec(Me.TxtDiscount.Text)) / 100, Decimal_Def)
            Else
                'hitung disc %
                If CDec(Me.TxtUnitPrice.Text) = 0 Then
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                Else
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtDiscAmount.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2)
                End If
            End If

            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount.Text), Decimal_Def)
            'Obs
            'Me.TxtNettPrice.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.TxtUnitPrice.Text) * _
            '(100 - CDec(Me.TxtDiscount.Text)) / 100, Decimal_Def))
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text) * _
                CDec(Me.TxtQtyMulti.Text), Decimal_Def)
            Else
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text) * _
                CDec(Me.TxtQtyDefault.Text), Decimal_Def)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim rowDGV As DataGridViewRow
        Dim AngkaDecimal As Integer = 0
        Dim ObjAOPCustom As New CustomSMESS.ClsTransaksiAOP

        Try
            'data salah jika:
            '1. tidak diisikan, kode product, currency
            'cek data yang harus diisi
            If Me.chkPONumber.Checked = False Then
                KoleksiNamaData.Add("Code no")
                KoleksiTextBox.Add(Me.TxtCodeNo)
            Else
                If Me.cboPurchaseOrder.Text = "" Then
                    MsgBox("Please specify purchase order number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If Me.cboCodeNo.Text = "" Then
                    MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            KoleksiNamaData.Add("Currency code")
            KoleksiTextBox.Add(Me.TxtCurrencyCode1)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. jika sudah ada item, maka kode currencynya harus sama
            If Me.dgvItem.Rows.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(0)
                If Me.TxtPrimaryCurrency1.Text <> rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value Then
                    MsgBox("Currency must be identical for each item", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '3. berkaitan dengan data angka:
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            '   a. unit price, qty harus positif
            KoleksiNamaData.Add("Unit Price")
            KoleksiTextBox.Add(Me.TxtUnitPrice)
            KoleksiNamaData.Add("Qty")
            KoleksiTextBox.Add(Me.TxtQtyMulti)
            If ModCekData.CekBilanganPositif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            '   b. discount paling besar 100% --> otomatis + non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            KoleksiNamaData.Add("discount")
            KoleksiTextBox.Add(Me.TxtDiscount)
            If ModCekData.CekNonNegatif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '4. Remarks maksimal 150 karakter
            If Me.TxtRemarkItem.Text.Length > 150 Then
                MsgBox("Maximum remarks data length is 150 character", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlled.Checked = True Then
                If Me.txtBatchNo.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtBatchNo.Focus()
                    Return False
                End If
            End If

            '6. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlled.Checked = True Then
                If Me.txtPrimaryLocation.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtLocation.Focus()
                    Return False
                End If
            End If

            '7. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlled.Checked = True And Me.chkExpiryDateControlled.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.txtBatchNo.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '7b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColLotNumber.Index).Value = Me.txtBatchNo.Text And Me.txtBatchNo.Text <> "" And row.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColExpiry.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.txtBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(10).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.DtpExpired.Focus()
                            Return False
                        End If
                    End If
                Next
                '7.c Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam transaksi lainnya: stock in adjust, stock in purchase, 
                '    stock in sales return, stock tag, transaksi bundling (item before) 
                '    (untuk stock out tidak dicek karena semua harus masuk dulu sebelum keluar, jadi dimasukka ke dalam pengecekan
                '    adalah semua transaksi yang menjadi pintu masuk inventory)
                '1. stock in (semua stock in)
                Dim JenisTransaksi As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
                        Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                For loopInOut As Integer = 0 To 4 Step 1
                    Select Case loopInOut
                        Case 0
                            '   a. stock in purchase
                            JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                            Exit Select
                        Case 1
                            '   b. stock in sales return
                            JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            Exit Select
                        Case 2
                            '   c. stock in adjust
                            JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                            Exit Select
                        Case 3
                            '   d. stock in back
                            JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                            Exit Select
                        Case 4
                            '   e. stock in proses
                            JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                            Exit Select
                    End Select
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text, _
                    JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> Me.DtpExpired.Value Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNo.Text) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.DtpExpired.Focus()
                            Return False
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNo.Text) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")
            End If

            '8. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlled.Checked = True Then
                If Me.DtpExpired.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
            End If

            If Me.chkPONumber.Checked = True Then
                '9. berkaitan dengan data order balance
                If CDec(Me.TxtOrderBalance.Text) < CDec(Me.TxtQtyMulti.Text) Then
                    MsgBox("Maximum value your can enter is " + Me.ObjCD.DigitGrouping(Me.TxtOrderBalance.Text, _
                    Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + Me.LblUnitsMulti.Text + ".", MsgBoxStyle.Information, "Information")
                    Me.TxtQtyMulti.Focus()
                    Return False
                End If

                '10. cek bila purchase order belum diotorisasi maka tidak bisa masuk
                If Me.ObjPO.GetAuthorizationLevel(Me.txtPrimaryPO.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.txtPrimaryPO.Text) Then
                    MsgBox("This purchase order require more authorization, please check authorization needed.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '11. untuk po yang masuk, harus sama currency-nya
                If Me.dgvItem.RowCount > 0 Then
                    Dim mTaxNew, mprimaryCurrency, mTaxOld, mCurrencyOld, vPrimaryPO_Old As Decimal
                    vPrimaryPO_Old = Me.GetPrimaryPOPertama

                    If vPrimaryPO_Old <> -1 Then
                        Me.ObjPO.GetPrimaryCurrency_Tax(Me.txtPrimaryPO.Text, mTaxNew, mprimaryCurrency)
                        Me.ObjPO.GetPrimaryCurrency_Tax(vPrimaryPO_Old, mTaxOld, mCurrencyOld)

                        '   a. % pajak yang digunakan sama
                        If mTaxNew <> mTaxOld Then
                            MsgBox("This purchase order item cannot be in one stock in purchase with the existing items. " + _
                            "For several purchase order item written in one stock in purchase their currency and tax must be identical", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        '   b. primary currency sama
                        If mprimaryCurrency <> mCurrencyOld Then
                            MsgBox("This purchase order item cannot be in one stock in purchase with the existing items. " + _
                            "For several purchase order item written in one stock in purchase their currency and tax must be identical", _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If


            '12. nilai qty default <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyDefault.Text) = False Then
                MsgBox("Qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyMulti.Focus()
                Return False
            End If

            'cek batasan custom
            If ObjAOPCustom.ApakahDataBenarItem(Me.TxtPrimaryCodeNo.Text, Me.txtPriceByMultiUnit.Text, _
            Me.TxtPrimaryUnitsMulti.Text, Me.TxtPrimaryUnitsDefault.Text) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            ObjAOPCustom = Nothing
        End Try
    End Function

    Private Function GetPrimaryPOPertama() As Double
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    Return rowDGV.Cells(Me.ColPrimaryPO.Index).Value
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex() As Integer
        Dim BatchNo As String = ""
        Dim PrimaryLocation As Double = -1
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                BatchNo = row.Cells(Me.ColLotNumber.Index).Value

                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And BatchNo = Me.txtBatchNo.Text And _
                CDec(Me.TxtUnitPrice.Text) = CDec(row.Cells(Me.ColUnitPrice.Index).Value) And _
                CDec(Me.TxtDiscount.Text) = CDec(row.Cells(Me.ColDiscount.Index).Value) And _
                CDec(Me.txtPrimaryLocation.Text) = row.Cells(Me.ColPrimaryLocation.Index).Value Then
                    If Me.chkPONumber.Checked = False Then
                        If row.Cells(Me.ColReferedToPO.Index).Value = 0 Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColReferedToPO.Index).Value = 1 And row.Cells(Me.ColPrimaryPO.Index).Value = Me.txtPrimaryPO.Text Then
                            Return row.Index
                        End If
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1
        Dim PrimaryLocationDummy As Double = -1
        Dim obj As New Core_Module.clsMasterStorageArea

        Try
            If Me.ApakahDataItemBenar = True Then
                If Me.txtPrimaryLocation.Text.Trim <> -1 Then
                    PrimaryLocationDummy = Me.txtPrimaryLocation.Text
                End If

                indexSama = Me.GetIndex
                If indexSama = -1 Then
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                    AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)

                    With row
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        If Me.chkPONumber.Checked = True Then
                            .Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                        Else
                            .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        End If
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColSuppliersCodeNumber.Index).Value = ""
                        .Cells(Me.ColSuppliersProductDesc.Index).Value = ""
                        .Cells(Me.ColPONO.Index).Value = Me.cboPurchaseOrder.Text
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColGRNNO.Index).Value = ""
                        .Cells(Me.ColStockOutNo.Index).Value = ""
                        .Cells(Me.ColLotNumber.Index).Value = Me.txtBatchNo.Text
                        If Me.DtpExpired.Visible = False Then
                            .Cells(Me.ColExpiry.Index).Value = #1/1/1950#
                        Else
                            .Cells(Me.ColExpiry.Index).Value = Me.DtpExpired.Value
                        End If
                        .Cells(Me.ColLocation.Index).Value = Me.TxtLocation.Text
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtDiscAmount.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscFromPerct.Index).Value = Me.TxtDiscFromPerct.Text

                        ''@rosadinnn
                        ''@MUaT
                        .Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyDefault.Text), AngkaDecimal)
                        .Cells(Me.ColUnit.Index).Value = Me.LblUnitsDefault.Text

                        .Cells(Me.ColSecQty.Index).Value = -1
                        .Cells(Me.ColSecUnit.Index).Value = ""
                        .Cells(Me.ColCurrency.Index).Value = Me.TxtCurrencyCode1.Text
                        .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscount.Index).Value = CDec(Me.TxtDiscount.Text)
                        .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtAmount.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColRemarksItem.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryGRN.Index).Value = -1
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryPO.Index).Value = Me.txtPrimaryPO.Text
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryUnit.Index).Value = Me.TxtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                        .Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency1.Text
                        .Cells(Me.ColShowDetail.Index).Value = "Show"
                        If Me.txtPrimaryPO.Text <> "" And Me.txtPrimaryPO.Text <> "-1" Then
                            .Cells(Me.ColReferedToPO.Index).Value = 1
                        Else
                            .Cells(Me.ColReferedToPO.Index).Value = 0
                        End If
                        .Cells(Me.ColItemAdjustment.Index).Value = 0
                        .Cells(Me.ColItemAdjustmentReason.Index).Value = ""
                        .Cells(Me.ColItemAdjustmentUser.Index).Value = ""
                        .Cells(Me.ColItemAdjustmentTimestamp.Index).Value = "-"

                        .Cells(Me.ColTotalQtyMulti.Index).Value = Me.TxtQtyMulti.Text
                        .Cells(Me.ColUnitsCodeMulti.Index).Value = Me.LblUnitsMulti.Text
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.TxtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                        If Me.TxtPrimaryUnitsMulti.Text <> Me.TxtPrimaryUnitsDefault.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                                    ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)) _
                                    & " " & Me.LblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                                    Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) & " " & Me.LblUnitsDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If

                        .Cells(Me.ColPriceByMultiUnit.Index).Value = Me.txtPriceByMultiUnit.Text
                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnitsCodeMulti.Index).Value
                        Else
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnit.Index).Value
                        End If
                    End With
                Else
                    row = Me.dgvItem.Rows(indexSama)
                    'tambahkan qty, 

                    ''@rosadinnn
                    ''@MUaT
                    'row.Cells(Me.ColTotalQty.Index).Value = CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(Me.TxtQtyMulti.Text)
                    If Me.TxtPrimaryUnitsMulti.Text <> row.Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                        MsgBox("There is already  " & row.Cells(Me.ColTotalQtyMulti.Index).Value & " " & row.Cells(Me.ColUnit.Index).Value & " of " & Me.TxtCodeNo.Text & " exist." & vbNewLine & "You cannot input this item with different units.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(Me.TxtQtyMulti.Text), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                    'row.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(Me.TxtQtyDefault.Text), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))

                    Me.SetQtyDefaultOnDgv(indexSama)

                    'hitung ulang amount
                    row = Me.dgvItem.Rows(indexSama)
                    Me.Hitung_Amount_Multi(row)
                End If

                PrimaryCurrenct = Me.TxtPrimaryCurrency1.Text

                Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
                Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                Me.FormatDGV2()

                Me.cboCalculateBy.SelectedIndex = -1
                Me.cboCalculateBy.SelectedIndex = 0

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                        Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                Me.Hitung_Total_Multi()

                If Me.chkPONumber.Checked = True Then
                    'perhitungan menggunakan nilai jumlah dgv yang baru
                    Dim JmlObjLama, JmlPOBBaru, JmlDGV, JmlPOB As Decimal
                    JmlDGV = Me.GetJmlDGV_(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                    JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
                    JmlObjLama = Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                    JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV

                    If JmlPOBBaru <= 0 Then
                        'remove row, pada code no ybs
                        Dim rowView As Data.DataRowView = Me.cboCodeNo.SelectedItem
                        Dim row1 As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row = rowView.Row
                        Me.TBL_PO3.Rows.Remove(row1)
                        Me.cboCodeNo_SelectedValueChanged(sender, e)
                        Me.cboCodeNo.Focus()
                        If Me.cboCodeNo.Items.Count = 0 Then
                            Me.IsikanCBOPO()
                            Me.cboPurchaseOrder.Focus()
                        End If
                    Else
                        If Me.cboCodeNo.SelectedIndex <> Me.cboCodeNo.Items.Count - 1 Then
                            Me.cboCodeNo.SelectedIndex += 1
                            Me.cboCodeNo.Focus()
                        Else
                            'focus pada cbo code no
                            Me.cboCodeNo.Focus()
                            'hitung kembali order balance dan sisa free qty
                            Me.SetOrderBalance()
                        End If
                    End If

                    Me.TxtQtyMulti.Text = "0"
                    Me.txtAmount.Text = "0"
                    Me.TxtRemarkItem.Text = ""
                Else
                    Me.KosongkanItem()
                    Me.TxtCodeNo.Focus()
                End If

                If PrimaryLocationDummy <> -1 Then
                    Me.txtPrimaryLocation.Text = PrimaryLocationDummy
                    Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub FormatDGV2()
        Dim row As DataGridViewRow
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim TidakAdaPO As Boolean = True
        Dim TidakAdaMultiUnit As Boolean = True
        Dim AdaPriceByDefaultUnit As Boolean = False

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                If CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) <> CDec(row.Cells(Me.ColTotalQty.Index).Value) Then
                    TidakAdaMultiUnit = False
                End If
                If row.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    TidakAdaPO = False
                End If
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiry.Index).Value) = False Then
                    row.Cells(Me.ColExpiry.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiry.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiry.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiry.Index).Value = Format(CDate(row.Cells(Me.ColExpiry.Index).Value), "dd MMM yyyy")
                    End If
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(Me.ColPrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.ColLocation.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColLotNumber.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    AdaPriceByDefaultUnit = True
                End If
            Next
            Me.ColPriceByMultiUnit_Str.Visible = AdaPriceByDefaultUnit

            '2. tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False

            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColLotNumber.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLotNumber.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiry.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiry.Index).Visible = True
            End If
            'po
            If TidakAdaPO = False Then
                Me.ColPONO.Visible = True
                If Me.TxtRemark.Enabled = True Then
                    Me.txtSupplierCode.Enabled = False
                    Me.txtSupplierCode.TabStop = False
                    Me.btnFindSupplier.Enabled = False
                    Me.btnFindSupplier.TabStop = False
                End If
            Else
                Me.ColPONO.Visible = False
                If Me.TxtRemark.Enabled = True Then
                    Me.txtSupplierCode.Enabled = True
                    Me.txtSupplierCode.TabStop = True
                    Me.btnFindSupplier.Enabled = True
                    Me.btnFindSupplier.TabStop = True
                End If
            End If
            'multiunit
            If TidakAdaMultiUnit = True Then
                Me.ColTotalQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColTotalQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
    '    Dim row As DataGridViewRow
    '    Dim AngkaDecimal, IndexTujuan As Integer

    '    Try
    '        row = Me.dgvItem.Rows(e.RowIndex)
    '        With row
    '            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                'bila tidak ada perubahan maka tidak perlu lakukan apa2
    '                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
    '                    Exit Sub
    '                End If

    '                'bila ada perubahan:
    '                If Me.ColTotalQty.Index = e.ColumnIndex Or Me.ColUnitPrice.Index = e.ColumnIndex Or _
    '                Me.ColDiscount.Index = e.ColumnIndex Then
    '                    If e.ColumnIndex = Me.ColDiscount.Index Then
    '                        'pastikan nilai discount non negatif
    '                        If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                            Exit Sub
    '                        End If
    '                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
    '                        Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, 2)
    '                        'pastikan nilai diskon maksimum = 100
    '                        If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) >= 100 Then
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                            Exit Sub
    '                        End If
    '                    ElseIf e.ColumnIndex = Me.ColTotalQty.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
    '                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                            Exit Sub
    '                        End If
    '                        If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                            Exit Sub
    '                        Else
    '                            'pastikan nilai angka tidak lebih dari 9 M
    '                            If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 9999999999 Then
    '                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                                Exit Sub
    '                            End If
    '                        End If

    '                        'pastikan nilai unit price dan qty positif
    '                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                            Exit Sub
    '                        Else
    '                            If e.ColumnIndex = Me.ColTotalQty.Index Then
    '                                'pastikan nilai order balance tidak negatif setelah melakukan proses edit.
    '                                Dim Expiry As Date
    '                                If IsDate(row.Cells(Me.ColExpiry.Index).Value) = True Then
    '                                    Expiry = row.Cells(Me.ColExpiry.Index).Value
    '                                Else
    '                                    Expiry = #1/1/1950#
    '                                End If

    '                                Me.EndEditQty(sender, e, -1, row.Cells(Me.ColTotalQty.Index).Value, _
    '                                row.Cells(Me.ColLotNumber.Index).Value, Expiry, row)

    '                                'jika lewat semua maka hitung amount
    '                                Me.Hitung_Amount_Multi(Me.dgvItem.Rows(e.RowIndex))
    '                                'hitung ulang total
    '                                Me.Hitung_Total_Multi()

    '                                Exit Sub
    '                            End If

    '                            If e.ColumnIndex = Me.ColTotalQty.Index Then
    '                                AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
    '                            Else
    '                                AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
    '                            End If
    '                            'digit grouping
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
    '                            Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)
    '                        End If
    '                    End If


    '                    IndexTujuan = Me.GetIndexSama(e.RowIndex)
    '                    If IndexTujuan <> -1 Then
    '                        'gabungkan baris
    '                        Me.GabungkanBaris(e.RowIndex, IndexTujuan)
    '                    Else
    '                        'jika lewat semua maka hitung amount
    '                        Me.Hitung_Amount_Multi(Me.dgvItem.Rows(e.RowIndex))
    '                    End If
    '                    'hitung ulang total
    '                    Me.Hitung_Total_Multi()
    '                ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
    '                    'jika kolom remarks maka cek, length maksimum = 150
    '                    If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Length > 150 Then
    '                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
    '                        CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Substring(0, 150)
    '                    End If
    '                End If
    '            Else
    '                If e.ColumnIndex = Me.ColRemarksItem.Index Then
    '                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
    '                Else
    '                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                End If
    '            End If
    '        End With
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim row As DataGridViewRow
        Dim AngkaDecimal, IndexTujuan As Integer

        Try
            row = Me.dgvItem.Rows(e.RowIndex)
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'bila tidak ada perubahan maka tidak perlu lakukan apa2
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If

                    'bila ada perubahan:
                    If Me.ColTotalQtyMulti.Index = e.ColumnIndex Or Me.ColUnitPrice.Index = e.ColumnIndex Or _
                            Me.ColDiscount.Index = e.ColumnIndex Or _
                            Me.ColDiscAmount.Index = e.ColumnIndex Then
                        If e.ColumnIndex = Me.ColDiscount.Index Then
                            'pastikan nilai discount non negatif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                            'pastikan nilai diskon maksimum = 100
                            If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "1"
                            Me.HitungItem(e.RowIndex, 2)
                        ElseIf e.ColumnIndex = Me.ColDiscAmount.Index Then
                            'pastikan nilai discount non negatif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)
                            'pastikan nilai diskon < Unit Price
                            If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "0"
                            Me.HitungItem(e.RowIndex, AngkaDecimal)
                        ElseIf e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            Else
                                'pastikan nilai angka tidak lebih dari 9 M
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 9999999999 Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If

                            If Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = 0 Then
                                'pastikan nilainya lebih besar dari disc amount
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If

                            'pastikan nilai unit price dan qty positif
                            If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            Else
                                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                                    'pastikan nilai order balance tidak negatif setelah melakukan proses edit.
                                    Dim Expiry As Date
                                    If IsDate(row.Cells(Me.ColExpiry.Index).Value) = True Then
                                        Expiry = row.Cells(Me.ColExpiry.Index).Value
                                    Else
                                        Expiry = #1/1/1950#
                                    End If

                                    Me.EndEditQty_Multi(sender, e, -1, row.Cells(Me.ColTotalQty.Index).Value, _
                                        row.Cells(Me.ColLotNumber.Index).Value, Expiry, row)

                                    'jika lewat semua maka hitung amount
                                    Me.Hitung_Amount_Multi(Me.dgvItem.Rows(e.RowIndex))
                                    'hitung ulang total
                                    Me.Hitung_Total_Multi()

                                    Exit Sub
                                End If

                                If e.ColumnIndex = Me.ColTotalQty.Index Then
                                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                                Else
                                    AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                End If

                                'digit grouping
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)
                            End If
                        End If

                        'jika sudah sampai sini, maka adjustment reset ke nol kembali
                        Me.dgvItem(Me.ColItemAdjustment.Index, e.RowIndex).Value = 0
                        Me.dgvItem(Me.ColItemAdjustmentReason.Index, e.RowIndex).Value = ""
                        Me.dgvItem(Me.ColItemAdjustmentTimestamp.Index, e.RowIndex).Value = ""
                        Me.dgvItem(Me.ColItemAdjustmentUser.Index, e.RowIndex).Value = ""

                        Me.HitungItem(e.RowIndex, AngkaDecimal)
                        IndexTujuan = Me.GetIndexSama(e.RowIndex)
                        If IndexTujuan <> -1 Then
                            Me.IndexAsal = e.RowIndex
                            Me.IndexTujuan = IndexTujuan
                            Me.bgw1.RunWorkerAsync()
                            Exit Sub
                        Else
                            'jika lewat semua maka hitung amount
                            Me.Hitung_Amount_Multi(Me.dgvItem.Rows(e.RowIndex))
                        End If
                        'hitung ulang total
                        Me.Hitung_Total_Multi()
                    ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
                        'jika kolom remarks maka cek, length maksimum = 150
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Length > 150 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Substring(0, 150)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarksItem.Index Then
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

    Private Sub HitungItem(ByVal vRowIndex As Integer, ByVal rounding As Integer)
        Dim rowDgv As DataGridViewRow
        Dim DefaultPrecisionRow As Integer = 0
        Dim DiscPerct As Integer = 0

        Try
            rowDgv = Me.dgvItem.Rows(vRowIndex)
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(rowDgv.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            DiscPerct = rowDgv.Cells(Me.ColDiscFromPerct.Index).Value
            If DiscPerct = 1 Then
                rowDgv.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) * CDec(rowDgv.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecisionRow))
            Else
                rowDgv.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value)), 2))
            End If
            rowDgv.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) - CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecisionRow))

            If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, vRowIndex).Value = 1 Then
                Me.dgvItem.Item(Me.ColAmount.Index, vRowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, vRowIndex).Value) * _
                CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIndex).Value), rounding))
            Else
                Me.dgvItem.Item(Me.ColAmount.Index, vRowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, vRowIndex).Value) * _
                CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIndex).Value), rounding))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)
        Dim Use_POB_Keseluruhan As Boolean = False
        Dim PackingList_Ada_StockIn As Boolean = False

        Try
            With row
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                JmlDGV = .Cells(Me.ColTotalQty.Index).Value
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, -1, _
                .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColLotNumber.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                'cek apakah jumlah purchase order balance menjadi minus setelah perubahan data
                If row.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri As Decimal
                    JmlDGV = Me.GetJmlDGV_(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPO.Index).Value)
                    JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPO.Index).Value, _
                        .Cells(Me.ColPrimaryCodeNo.Index).Value)
                    JmlObjLama = Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColPrimaryPO.Index).Value)
                    JmlSendiri = .Cells(Me.ColTotalQty.Index).Value
                    'perhitungan menggunakan nilai jumlah dgv yang baru
                    JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama

                    If JmlPOBBaru < 0 Then
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(JmlPOB - (JmlDGV - JmlSendiri) + JmlObjLama, Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value)) + _
                            " " + CStr(Me.ObjProduct.GetUnitCode(.Cells(Me.ColPrimaryCodeNo.Index).Value)) + ".", _
                            MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer

                        Exit Sub
                    End If

                    'isikan ulang data order balance
                    Me.IsikanCBOPO()
                End If

                If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) = 0 Then
                    MsgBox("Qty restored to previous value, because qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditQty_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)
        Dim Use_POB_Keseluruhan As Boolean = False
        Dim PackingList_Ada_StockIn As Boolean = False

        Try
            With row
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data

                ''@rosadinnn
                ''@MUaT
                Me.SetQtyDefaultOnDgv(e.RowIndex)

                JmlDGV = .Cells(Me.ColTotalQty.Index).Value
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColLotNumber.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                ''digit grouping
                'Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                '    Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                    Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value))

                'cek apakah jumlah purchase order balance menjadi minus setelah perubahan data
                If row.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri As Decimal
                    JmlDGV = Me.GetJmlDGV_(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPO.Index).Value)
                    JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPO.Index).Value, _
                        .Cells(Me.ColPrimaryCodeNo.Index).Value)
                    JmlObjLama = Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColPrimaryPO.Index).Value)
                    JmlSendiri = .Cells(Me.ColTotalQty.Index).Value
                    'perhitungan menggunakan nilai jumlah dgv yang baru
                    JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama

                    If JmlPOBBaru < 0 Then
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(JmlPOB - (JmlDGV - JmlSendiri) + JmlObjLama, Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value)) + _
                            " " + CStr(Me.ObjProduct.GetUnitCode(.Cells(Me.ColPrimaryCodeNo.Index).Value)) + ".", _
                            MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer

                        Exit Sub
                    End If

                    'isikan ulang data order balance
                    Me.IsikanCBOPO()
                End If

                If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) = 0 Then
                    MsgBox("Qty restored to previous value, because qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

            ''@rosadinnn
            ''@MUaT
            Me.SetQtyDefaultOnDgv(e.RowIndex)

        End Try
    End Sub

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Decimal, _
            ByVal PrimaryPdO As Decimal, ByVal PrimaryProduct As Decimal, ByVal BatchNo As String, _
            ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal jmlDGV As Decimal, Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim QtyDummy As Decimal = 0

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, PrimaryLocation, PrimaryProduct, _
            BatchNo)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV_Default(PrimaryProduct, PrimaryLocation, BatchNo)
            JmlSetelahHapus = JmlGudang - JmlObjLama + jmlDGV

            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make your stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    QtyDummy = jmlDGV - JmlGudang - JmlSetelahHapus
                    QtyDummy = QtyDummy * CDec(row.Cells(Me.ColRatioReference.Index).Value) / CDec(row.Cells(Me.ColRatioDefault.Index).Value)
                    QtyDummy = ObjCD.RoundUp1(QtyDummy, Me.ObjProduct.GetAngkaDecimal(PrimaryProduct, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))

                    MsgBox("Minimum value you must can enter is " + Me.ObjCD.DigitGrouping(QtyDummy, _
                    Me.ObjProduct.GetAngkaDecimal(PrimaryProduct, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)) + _
                    " " + row.Cells(Me.ColUnitsCodeMulti.Index).Value.ToString + ".", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' menggabungkan 2 baris yang kemudian mengembalikan nilai baris hasil penggabungan.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GabungkanBaris(ByVal vIndexAsal As Integer, ByVal vIndexTujuan As Integer) As Integer
        Dim rowAsal, rowTujuan As DataGridViewRow

        Try
            rowAsal = Me.dgvItem.Rows(vIndexAsal)
            rowTujuan = Me.dgvItem.Rows(vIndexTujuan)
            'tambahakan qty pada rowtujuan
            rowTujuan.Cells(Me.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowAsal.Cells(Me.ColTotalQtyMulti.Index).Value) + _
            CDec(rowTujuan.Cells(Me.ColTotalQtyMulti.Index).Value), Me.ObjProduct.GetAngkaDecimal(rowTujuan.Cells(Me.ColPrimaryCodeNo.Index).Value, rowTujuan.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
            Me.SetQtyDefaultOnDgv(rowTujuan.Index)

            'reset nilai adjustment
            rowTujuan.Cells(Me.ColItemAdjustment.Index).Value = 0
            rowTujuan.Cells(Me.ColItemAdjustmentReason.Index).Value = ""
            rowTujuan.Cells(Me.ColItemAdjustmentTimestamp.Index).Value = "-"
            rowTujuan.Cells(Me.ColItemAdjustmentUser.Index).Value = ""

            'hitung ulang amount pada row tujuan
            Me.Hitung_Amount_Multi(rowTujuan)
            'buang baris row asal
            Me.dgvItem.Rows.Remove(rowAsal)
            'set ulang nomor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'kembalikan index rowtujuan
            Return rowTujuan.Index
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' memberikan nilai true jika baris yang diberikan indexnya harus melakukan penggabungan baris.
    ''' 2 baris digabungkan bila field berikut ini sama:
    ''' a. primary code number
    ''' b. unit price
    ''' c. discount
    ''' d. primary location
    ''' e. batch number
    ''' f. primary unit
    ''' g. primary currency
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndexSama(ByVal rowIndexAcuan As Integer) As Integer
        Dim rowDGV, rowAcuan As DataGridViewRow

        Try
            rowAcuan = Me.dgvItem.Rows(rowIndexAcuan)
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> rowIndexAcuan Then
                    With rowAcuan
                        If .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        CDec(.Cells(Me.ColUnitPrice.Index).Value) = CDec(rowDGV.Cells(Me.ColUnitPrice.Index).Value) And _
                        CDec(.Cells(Me.ColDiscount.Index).Value) = CDec(rowDGV.Cells(Me.ColDiscount.Index).Value) And _
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value And _
                        .Cells(Me.ColLotNumber.Index).Value = rowDGV.Cells(Me.ColLotNumber.Index).Value And _
                        .Cells(Me.ColPrimaryUnit.Index).Value = rowDGV.Cells(Me.ColPrimaryUnit.Index).Value And _
                        .Cells(Me.ColPrimaryCurrency.Index).Value = rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value Then
                            Return rowDGV.Index
                        End If
                    End With
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
    '    Try
    '        'yang boleh diedit hanya qty, unit price, remarks item, dan discount
    '        If e.ColumnIndex = Me.ColTotalQty.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or _
    '        e.ColumnIndex = Me.ColRemarksItem.Index Or e.ColumnIndex = Me.ColDiscount.Index Then
    '            Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
    '        Else
    '            e.Cancel = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang boleh diedit hanya qty, unit price, remarks item, dan discount
            If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or _
                    e.ColumnIndex = Me.ColRemarksItem.Index Or e.ColumnIndex = Me.ColDiscount.Index Or _
                    e.ColumnIndex = Me.ColDiscAmount.Index Then
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCurrencyCode1.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
                    Me.TxtUnitPrice.Focus()
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
                Me.SetCalculateBy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.Leave
        Dim i As New FrmCariMasterCurrency

        Try
            If Me.TxtCurrencyCode1.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)

                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        If Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        ElseIf Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        Else
                            If Me.TxtCodeNo.Visible = True Then
                                Me.TxtCodeNo.Focus()
                            Else
                                Me.cboCodeNo.Focus()
                            End If
                        End If
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)

                        If My.Computer.Keyboard.ShiftKeyDown = True Then
                            If Me.TxtLocation.Visible = True Then
                                Me.TxtLocation.Focus()
                            ElseIf Me.txtBatchNo.Visible = True Then
                                Me.txtBatchNo.Focus()
                            Else
                                If Me.TxtCodeNo.Visible = True Then
                                    Me.TxtCodeNo.Focus()
                                Else
                                    Me.cboCodeNo.Focus()
                                End If
                            End If
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCurrency1.Text = "-1"

                If My.Computer.Keyboard.ShiftKeyDown = True Then
                    If Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    ElseIf Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    Else
                        If Me.TxtCodeNo.Visible = True Then
                            Me.TxtCodeNo.Focus()
                        Else
                            Me.cboCodeNo.Focus()
                        End If
                    End If
                Else
                    Me.TxtQtyMulti.Focus()
                End If
            End If
            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try

        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCurrency2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCurrency2.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                Me.TxtUnitPrice.Focus()
                Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.TxtCurrencyCode1.Text <> "" Then
                    i.txtCari.Text = Me.TxtCurrencyCode1.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                    Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                    Me.TxtUnitPrice.Focus()
                Else
                    Me.TxtCurrencyCode1.Clear()
                    Me.TxtCurrencyCode1.Focus()
                End If
            End If

            If Me.TxtRemarkItem.Enabled = True Then
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
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

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemark.Enabled = True Then
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

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemark.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemark.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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

    Private Sub txtBatchNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtBatchNo.ReadOnly = False And Me.txtBatchNo.Enabled = True Then
                    txtBatchNo.BackColor = BackColorAll
                Else
                    txtBatchNo.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
                End If


            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBatchNo.KeyUp
        Try
            If e.KeyData = Keys.Tab And e.Shift = False Then
                If Me.DtpExpired.Visible = True Then
                    Me.DtpExpired.Focus()
                ElseIf Me.TxtLocation.Visible = True Then
                    Me.TxtLocation.Focus()
                Else
                    Me.TxtTotal.Focus()
                End If
            ElseIf e.KeyData = Keys.Tab And e.Shift = True Then
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.Leave
        Try
            Me.txtBatchNo.Text = Me.txtBatchNo.Text.Trim
            If Me.TxtRemark.Enabled = True Then
                If Me.txtBatchNo.Text <> "" And Me.chkExpiryDateControlled.Checked = True Then
                    Me.DtpExpired.Value = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                    If Me.DtpExpired.Value = #1/1/1950# Then
                        Me.DtpExpired.Value = Today
                    End If
                Else
                    Me.DtpExpired.Value = #1/1/1950#
                End If
            End If

            If Me.txtBatchNo.Focused = True Then
                If My.Computer.Keyboard.ShiftKeyDown = True Then
                    Me.TxtCodeNo.Focus()
                Else
                    If Me.DtpExpired.Visible = True Then
                        Me.DtpExpired.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    ElseIf Me.TxtCurrencyCode1.ReadOnly = False Then
                        Me.TxtCurrencyCode1.Focus()
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanCBOPO()
        Try
            Me.cboPurchaseOrder.DataSource = Nothing
            Me.cboCodeNo.DataSource = Nothing

            If Me.chkPONumber.Checked = True Then
                Me.IsiCboPurchaseOrder(Me.txtPrimarySupplier.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboPurchaseOrder(ByVal vPrimarySupplier As Decimal)
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data awal
            Me.cboPurchaseOrder.DataSource = Nothing
            Me.cboPurchaseOrder.Items.Clear()

            Me.TBL_PO.Rows.Clear()
            Me.TBL_PO2.Rows.Clear()
            Me.TBL_PO3.Rows.Clear()
            Me.TBL_PO4.Rows.Clear()

            'isikan dari basis data-persedian terakhir
            Me.ObjPOItem.IsikanTBLPO_Default(vPrimarySupplier, Me.TBL_PO, Me.cboPurchaseOrder)

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiAOP.tbltrans_aop_item_view1Row
            Dim row2 As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row
            Dim ketemu As Boolean

            If Me.baru = False Then
                Me.ObjAOPItem.IsiTBLPO2(Me.TxtPrimaryTransaksiAOP.Text, Me.TBL_PO2)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL_PO2.Rows.Count > 0
                    row1 = Me.TBL_PO2.Rows(0)
                    For Each row2 In Me.TBL_PO.Rows
                        If row1.Primary_Supplier = row2.Primary_Supplier _
                        And row1.Primary_Product = row2.Primary_Product And _
                        row2.Primary_PO = row1.Primary_PO_Purchase Then
                            row2.OrderBalance = row2.OrderBalance + row1.TotalQty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL_PO.Rows.Add(row1.Primary_PO_Purchase, row1.Primary_Supplier, row1.Primary_Product, _
                        row1.Product_Code, row1.No_PO, row1.TotalQty)
                    End If
                    Me.TBL_PO2.Rows.Remove(row1)
                End While
            End If

            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBL_PO()
            While Me.TBL_PO2.Rows.Count > 0
                row1 = Me.TBL_PO2.Rows(0)
                For Each row2 In Me.TBL_PO.Rows
                    If row1.Primary_Supplier = row2.Primary_Supplier _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_PO = row1.Primary_PO_Purchase Then
                        row2.OrderBalance = row2.OrderBalance - row1.TotalQty
                        If row2.OrderBalance = 0 Then
                            Me.TBL_PO.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL_PO.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL_PO2.Rows.Remove(row1)
            End While

            'pastikan hanya 1 so disebutkan hanya 1x
            Dim idx As Integer = 0
            Dim idx2 As Integer = 0
            Dim dummyRow As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row
            While idx < Me.TBL_PO.Rows.Count
                row2 = Me.TBL_PO(idx)
                idx2 = idx + 1
                While idx2 <= Me.TBL_PO.Rows.Count - 1
                    dummyRow = Me.TBL_PO.Rows(idx2)
                    If dummyRow.PO_Number = row2.PO_Number Then
                        Me.TBL_PO.Rows.Remove(dummyRow)
                    Else
                        idx2 += 1
                    End If
                End While

                'pindah ke baris selanjutnya
                idx += 1
            End While

            Me.cboPurchaseOrder.DataSource = Me.TBL_PO.DefaultView
            Me.cboPurchaseOrder.ValueMember = "PO Number"
            Me.cboPurchaseOrder.DisplayMember = "PO Number"
            Me.cboPurchaseOrder.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub IsikanTBL_PO()
        Try
            Dim row As DataGridViewRow
            Dim row1, row2 As DseTransaksiAOP.tbltrans_aop_item_view1Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL_PO2.Rows
                        If row1.Primary_Supplier = Me.txtPrimarySupplier.Text And _
                        row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        row1.Primary_PO_Purchase = row.Cells(Me.ColPrimaryPO.Index).Value Then
                            'kalo ketemu ubah data order balance
                            row1.TotalQty = row1.TotalQty + CDec(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        row2 = Me.TBL_PO2.Newtbltrans_aop_item_view1Row

                        With row2
                            .Primary_Transaksi = Me.TxtPrimaryTransaksiAOP.Text
                            .Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                            .Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                            .TotalQty = row.Cells(Me.ColTotalQty.Index).Value
                            .BatchNumber = row.Cells(Me.ColLotNumber.Index).Value
                            .Primary_PO_Purchase = row.Cells(Me.ColPrimaryPO.Index).Value
                            .No_PO = row.Cells(Me.ColPONO.Index).Value
                            .Primary_Supplier = Me.txtPrimarySupplier.Text
                            .Supplier_Code = Me.txtSupplierCode.Text
                            .At_Amount = row.Cells(Me.ColAtAmount.Index).Value
                        End With

                        Me.TBL_PO2.Rows.Add(row2)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboProduct(ByVal vPrimaryPO As Decimal)
        'Try
        '    Me.TBL_PO3.Rows.Clear()
        '    Me.TBL_PO4.Rows.Clear()
        '    Windows.Forms.Cursor.Current = Cursors.WaitCursor

        '    'isikan dari basis data-persedian terakhir
        '    Me.cboCodeNo.DataSource = Nothing
        '    Me.cboCodeNo.Items.Clear()

        '    Me.ObjPOItem.IsikanTBLProduct(Me.txtPrimaryPO.Text, Me.TBL_PO3, Me.cboCodeNo)

        '    'tambahkan item yang sudah dicatatkan ke dalam database
        '    Dim row1 As DseTransaksiAOP.tbltrans_aop_item_view1Row
        '    Dim row2 As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row
        '    Dim ketemu As Boolean
        '    If Me.baru = False Then
        '        Me.ObjAOPItem.IsiTBLProductByPO(Me.TxtPrimaryTransaksiAOP.Text, _
        '        Me.TBL_PO4, Me.txtPrimaryPO.Text)
        '        'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
        '        While Me.TBL_PO4.Rows.Count > 0
        '            row1 = Me.TBL_PO4.Rows(0)
        '            For Each row2 In Me.TBL_PO3.Rows
        '                If row1.Primary_Supplier = row2.Primary_Supplier _
        '                And row1.Primary_Product = row2.Primary_Product And _
        '                row2.Primary_PO = row1.Primary_PO_Purchase Then
        '                    row2.OrderBalance = row2.OrderBalance + row1.TotalQty
        '                    ketemu = True
        '                    Exit For
        '                End If
        '            Next
        '            If ketemu = False Then
        '                'tambahkan baris baru
        '                Me.TBL_PO3.Rows.Add(row1.Primary_PO_Purchase, row1.Primary_Supplier, row1.Primary_Product, _
        '                row1.Product_Code, row1.No_PO, row1.TotalQty)
        '            End If
        '            Me.TBL_PO4.Rows.Remove(row1)
        '        End While
        '    End If

        '    'kurangkan item yang sudah dimasukkan ke dgv semua
        '    Me.IsiTBLProduct(Me.txtPrimaryPO.Text)
        '    While Me.TBL_PO4.Rows.Count > 0
        '        row1 = Me.TBL_PO4.Rows(0)
        '        For Each row2 In Me.TBL_PO3.Rows
        '            If row1.Primary_Supplier = row2.Primary_Supplier _
        '            And row1.Primary_Product = row2.Primary_Product And _
        '            row2.Primary_PO = row1.Primary_PO_Purchase Then
        '                row2.OrderBalance = row2.OrderBalance - row1.TotalQty
        '                If row2.OrderBalance = 0 Then
        '                    Me.TBL_PO3.Rows.Remove(row2)
        '                End If
        '                ketemu = True
        '                Exit For
        '            End If
        '        Next
        '        If ketemu = False And TBL_PO3.Rows.Count > 0 Then
        '            'tidak mungkin tidak ketemu
        '            MsgBox("Logic error. Please contact your vendor if you see this message")
        '            Exit Sub
        '        End If
        '        Me.TBL_PO4.Rows.Remove(row1)
        '    End While

        '    Me.TBL_PO3.DefaultView.Sort = "[Product Code] ASC"
        '    Me.cboCodeNo.DataSource = Me.TBL_PO3.DefaultView
        '    Me.cboCodeNo.DisplayMember = "Product Code"
        '    Me.cboCodeNo.Refresh()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub IsiTBLProduct(ByVal vPrimaryPO As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiAOP.tbltrans_aop_item_view1Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And row.Cells(Me.ColReferedToPO.Index).Value = 1 Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL_PO4.Rows
                        If row1.Primary_Supplier = Me.txtPrimarySupplier.Text And _
                        row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        row1.Primary_PO_Purchase = row.Cells(Me.ColPrimaryPO.Index).Value Then
                            'kalo ketemu ubah data order balance
                            row1.TotalQty = row1.TotalQty + CDec(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        row1 = Me.TBL_PO4.Newtbltrans_aop_item_view1Row

                        With row1
                            .Primary_Transaksi = Me.TxtPrimaryTransaksiAOP.Text
                            .Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                            .Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                            .TotalQty = row.Cells(Me.ColTotalQty.Index).Value
                            .BatchNumber = row.Cells(Me.ColLotNumber.Index).Value
                            .Primary_PO_Purchase = row.Cells(Me.ColPrimaryPO.Index).Value
                            .No_PO = row.Cells(Me.ColPONO.Index).Value
                            .Primary_Supplier = Me.txtPrimarySupplier.Text
                            .Supplier_Code = Me.txtSupplierCode.Text
                            .At_Amount = row.Cells(Me.ColAtAmount.Index).Value
                        End With
                        'Me.TBL_PO4.Rows.Add(Me.TxtPrimaryTransaksiAOP.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        'row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColTotalQty.Index).Value, row.Cells(Me.ColLotNumber.Index).Value, _
                        'row.Cells(Me.ColPrimaryPO.Index).Value, row.Cells(Me.ColPONO.Index).Value, Me.txtPrimarySupplier.Text, _
                        'Me.txtSupplierCode.Text)

                        Me.TBL_PO4.Rows.Add(row1)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetOrderBalance()
        Try
            'berdasarkan po number
            If IsNumeric(Me.TxtPrimaryCodeNo.Text) = True And IsNumeric(Me.txtPrimaryPO.Text) = True Then
                Me.TxtOrderBalance.Text = Me.ObjCD.DigitGrouping(Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) _
                + Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) - _
                Me.GetJmlDGV_(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlDGV_(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim retVal As Decimal
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And rowDGV.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO Then
                    retVal += CDec(rowDGV.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub chkPONumber_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPONumber.CheckedChanged
        Try
            If Me.chkPONumber.Checked = True Then
                Me.IsikanCBOPO()
                Me.cboCodeNo.TabStop = True

                Me.TxtCodeNo.Visible = False
                Me.TxtCodeNo.TabStop = False
                Me.BtnFindCodeNo.Visible = False
                Me.BtnFindCodeNo.TabStop = False

                Me.cboPurchaseOrder.TabStop = True
                Me.cboPurchaseOrder.Enabled = True

                Me.cboCodeNo.Visible = True
                Me.cboCodeNo.TabStop = True
                Me.TxtOrderBalance.Visible = True
                Me.LblOrderBalance.Visible = True
                Me.SplitContainer6.Visible = True
            Else
                Me.cboPurchaseOrder.TabStop = False
                Me.cboPurchaseOrder.Enabled = False
                Me.cboPurchaseOrder.DataSource = Nothing
                Me.cboPurchaseOrder.Items.Clear()

                Me.KosongkanItem()
                Me.txtPrimaryPO.Text = "-1"
                Me.cboCodeNo.TabStop = False

                Me.TxtCodeNo.Visible = True
                Me.TxtCodeNo.TabStop = True
                Me.BtnFindCodeNo.Visible = True
                Me.cboCodeNo.DataSource = Nothing
                Me.cboCodeNo.Items.Clear()

                Me.cboCodeNo.Visible = False
                Me.cboCodeNo.TabStop = False
                Me.TxtOrderBalance.Visible = False
                Me.LblOrderBalance.Visible = False
                Me.SplitContainer6.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
            Dim PrimaryLocation, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
            JmlDGV, JmlSetelahHapus As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = e.Row

            With row
                PrimaryLocation = .Cells(Me.ColPrimaryLocation.Index).Value
                PrimaryProduct = .Cells(Me.ColPrimaryCodeNo.Index).Value
                BatchNo = .Cells(Me.ColLotNumber.Index).Value
                If IsDate(.Cells(Me.ColExpiry.Index).Value) = True Then
                    ExpiryDate = CDate(.Cells(Me.ColExpiry.Index).Value)
                Else
                    ExpiryDate = #1/1/1950#
                End If
                SecondaryQty = -1
                JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                'karena akan dihapus, perhitungan menggunakan nilai total jumlah dgv - qty dgv yang akan dihapuskan

                JmlDGV = Me.GetJmlDGV_Default(PrimaryProduct, PrimaryLocation, BatchNo) - CDec(.Cells(Me.ColTotalQty.Index).Value)

                If Me.baru = True Then
                    JmlObjLama = 0
                Else
                    JmlObjLama = Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, PrimaryLocation, PrimaryProduct, _
                    BatchNo)
                End If

                JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
                If JmlSetelahHapus < 0 Then
                    MsgBox("You cannot delete this data, because this will created minus quantity .", _
                    MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNumber As String) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim retVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                rowDGV.Cells(Me.ColLotNumber.Index).Value = vBatchNumber Then
                    retVal += rowDGV.Cells(Me.ColTotalQty.Index).Value
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNumber As String) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim retVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                rowDGV.Cells(Me.ColLotNumber.Index).Value = vBatchNumber Then
                    retVal += rowDGV.Cells(Me.ColTotalQtyMulti.Index).Value
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cboCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cboCodeNo.KeyDown

    End Sub

    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Try
            If Me.cboCodeNo.SelectedIndex <> -1 Then
                Dim RowView As Data.DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
                Dim row2 As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row = RowView.Row
                Dim PrimaryCodeNoBaru As Double = row2.Primary_Product

                If Me.TxtPrimaryCodeNo.Text <> CStr(PrimaryCodeNoBaru) Then
                    Me.TxtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtPrimaryUnitsDefault.Text = ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)

                    'isikan data currency item, unit price, disc, dan nett price
                    Dim mPrimaryCurrency, mUnitPrice, mDisc, mNettPrice, mDecimal As Decimal
                    mPrimaryCurrency = Me.ObjPO.GetPrimaryCurrency(Me.txtPrimaryPO.Text)
                    Me.ObjPOItem2.GetHargaDetail(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, mUnitPrice, mDisc, mNettPrice)

                    Me.TxtPrimaryCurrency1.Text = mPrimaryCurrency
                    Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                    mDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)
                    Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(mUnitPrice, mDecimal)
                    If mDisc <> 0 Then
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(mDisc, 2)
                    Else
                        Me.TxtDiscount.Text = "0"
                    End If
                    Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(mNettPrice, mDecimal)
                End If
            Else
                'reset tampilan
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.TxtPrimaryUnitsDefault.Text = "-1"
                Me.TxtQtyMulti.Text = "0"
                Me.TxtOrderBalance.Text = "0"
            End If

            'format tampilan
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.LblUnitsMulti.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                Me.TxtQtyMulti.Text = "0"
                Me.SetOrderBalance()
            Else
                Me.LblUnitsMulti.Text = "Units"
            End If
            Me.SetGboItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPurchaseOrder_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPurchaseOrder.SelectedValueChanged
        Try
            If Me.cboPurchaseOrder.SelectedIndex <> -1 And Me.cboPurchaseOrder.Text <> "" Then
                Dim RowView As Data.DataRowView = Me.cboPurchaseOrder.Items(Me.cboPurchaseOrder.SelectedIndex)
                Dim row2 As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row = RowView.Row
                Dim PrimaryPOBaru As Decimal = row2.Primary_PO
                If Me.txtPrimaryPO.Text <> CStr(PrimaryPOBaru) Then
                    Me.txtPrimaryPO.Text = PrimaryPOBaru
                    Me.IsiCboProduct(Me.txtPrimaryPO.Text)
                End If

                Me.cboCodeNo.TabStop = True

                Me.TxtCodeNo.Visible = False
                Me.TxtCodeNo.TabStop = False
                Me.BtnFindCodeNo.Visible = False
                Me.BtnFindCodeNo.TabStop = False
            Else
                Me.txtPrimaryPO.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim obj As New ClsTransaksiAOP
        Dim ij As New FrmCetakTransaksiAOP
        Dim ObjSettin As New Core_Module.ClsTransaksiSetting

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If Me.TxtPrimaryTransaksiAOP.Text = "" Or Me.TxtPrimaryTransaksiAOP.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryTransaksiAOP.Text <> "" And Me.TxtPrimaryTransaksiAOP.Text <> "-1" Then
                If ObjSettin.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    Dim objCustom As New CustomSMESS.ClsTransaksiAOP
                    If objCustom.CetakCustomAOP(ij.crv1, Me.TxtPrimaryTransaksiAOP.Text, Me.txtPrimaryCurrency.Text) = True Then
                        ij.ShowDialog()
                    End If
                    objCustom = Nothing
                    Exit Sub
                Else
                    obj.CetakAop(ij.crv1, Me.TxtPrimaryTransaksiAOP.Text)
                    ij.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSettin = Nothing
            obj = Nothing
            ij.Dispose()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
        End Try
    End Sub

    Private Sub SplitContainer3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer3.GotFocus
        Try
            If My.Computer.Keyboard.ShiftKeyDown = True Then
                Me.SelectNextControl(Me.SplitContainer3, False, False, True, True)
            Else
                Me.SelectNextControl(Me.SplitContainer3, False, True, True, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SplitContainer5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer5.GotFocus
        Try
            If My.Computer.Keyboard.ShiftKeyDown = True Then
                Me.SelectNextControl(Me.SplitContainer5.Panel1, False, False, True, True)
            Else
                Me.SelectNextControl(Me.SplitContainer5.Panel1, False, True, True, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCodeNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedIndexChanged

    End Sub

    Private Sub cboCodeNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboCodeNo.KeyPress
        MsgBox("")
    End Sub

    Private Sub TxtCodeNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCodeNo.TextChanged

    End Sub

    Private Sub cboCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.Leave
        Try
            If My.Computer.Keyboard.ShiftKeyDown = False Then
                If Me.txtBatchNo.Visible = True Then
                    Me.txtBatchNo.Focus()
                ElseIf Me.TxtLocation.Visible = True Then
                    Me.TxtLocation.Focus()
                Else
                    If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
            Else
                Me.cboPurchaseOrder.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtLocation.ReadOnly = False And Me.TxtLocation.Enabled = True Then
                TxtLocation.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
            End If
        End If
    End Sub

    Private Sub txtBatchNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtBatchNo.ReadOnly = False And Me.txtBatchNo.Enabled = True Then
                txtBatchNo.BackColor = Me.TxtPrimaryTransaksiAOP.BackColor
            End If
        End If
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

                ''@rosadinnn
                ''@MUaT
                Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)

                Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtPrimaryTransaksiAOP.BackColor, False)
            End If
            If Me.TabControl1.SelectedIndex = 1 Then
                Me.chkPONumber.Checked = False
                Me.KosongkanItem()
                Me.SetLocation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoiceNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoiceNumber.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtTaxInvoiceNumber.ReadOnly = True Then
                    Me.TxtTaxInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtTaxInvoiceNumber.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoiceNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoiceNumber.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtTaxInvoiceNumber.ReadOnly = True Then
                    Me.TxtTaxInvoiceNumber.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtTaxInvoiceNumber.BackColor = Me.TxtPrimaryCodeNo.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtNPWPNo.ReadOnly = True Then
                    Me.TxtNPWPNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtNPWPNo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNPWPNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtNPWPNo.ReadOnly = True Then
                    Me.TxtNPWPNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtNPWPNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
                End If
            End If
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

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.ObjAOP.CariDetail(vPrimary, _
               Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut) _
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

    Private Sub FrmTransaksiAOP_Quick2_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim ij As New Threading.Thread(AddressOf Me.UpdateAtAmount1)

        Try
            If Me.UpdateAtAmount = True Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.WindowState = FormWindowState.Normal
                Me.Left = Me.Width * -1
                Me.Top = Me.Height * -1

                'ij.Start()
                Me.UpdateAtAmount1()
            End If
            Me.SetLocation()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateAtAmount1()
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aopDataTable
        Dim row As DseTransaksiAOP.tbltrans_aopRow
        Dim FrmPO As New FrmTransaksiAOP_Quick_DetailNumber
        Dim ObjAOPItem As New ClsTransaksiAOPItem

        Try
            'siapkan form
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ObjAOP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP
            Me.ColATCost.Visible = True
            Me.ColShowDetail.Visible = True
            'Me.SplitContainer1.Panel2Collapsed = True
            'Me.SplitContainer3.TabStop = False
            'Me.SplitContainer3.Parent = Me.SplitContainer5.Panel1

            'hitung ulang at amount
            DTA.FillBy_QuickOnly(TBL)
            Me.UpdatePrg.Value = 0
            Me.UpdatePrg.Maximum = TBL.Rows.Count

            For Each row In TBL.Rows
                Me.UpdatePrg.Value += 1

                'cari detail
                If Me.ObjAOP.CariDetail(row.Primary, _
                             Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut) _
                             = True Then

                    'salin ke field
                    Me.SalinKeField()

                    'hitung at cost
                    Me.Hitung_AtCost_Multi()

                    FrmPO.PrimaryAOPQuickAcuan = Me.TxtPrimaryTransaksiAOP.Text
                    FrmPO.BackColorAll = Me.BackColorAll
                    If Me.baru = False Then
                        FrmPO.NoPO_Lama = Me.ObjPO.GetNoTransaksiPO_Quick(Me.TxtPrimaryTransaksiAOP.Text)
                    Else
                        FrmPO.NoPO_Lama = ""
                    End If

                    'simpan perubahan
                    Me.SalinKeObjek()
                    ObjAOPItem.Hapus(Me.ObjAOP.Primary)
                    ObjAOPItem.Simpan_Multi(Me.dgvItem, Me.dgvExchageRates, Me.DGVdp, Me.ObjAOP.Primary, Me.DGVitemPR, Me.dgvCost, Me, Me.TblCustomDummy)
                End If
            Next

            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub setLocation_Product()

        Try
            If Me.CboUnitMulti.Visible = True Then
                Me.TxtQtyMulti.Width = Me.TxtDiscount.Width - Me.CboUnitMulti.Width - 5
                Me.CboUnitMulti.Left = Me.TxtDiscount.Right - Me.CboUnitMulti.Width
                visibleCBO = True
            Else
                Me.TxtQtyMulti.Width = Me.TxtDiscount.Width - Me.LblUnitsMulti.Width - 5
                Me.LblUnitsMulti.Left = Me.TxtDiscount.Right - Me.LblUnitsMulti.Width
                visibleCBO = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SetLocation()

        Try
            'Me.gboDetail.Width = Me.Width - 10
            'Me.gboDetail.Left = Me.Left + 5
            ''Me.gboDetail.Top = Me.Top
            ''Me.gboDetail.Height = Me.Bottom - 50
            ''Me.gboDetail.Bottom = Me.Bottom - 50

            Me.SplitContainer5.Width = Me.dgvItem.Width
            Me.SplitContainer5.Left = Me.dgvItem.Left
            'Me.SplitContainer5.Panel1.Width = Me.SplitContainer5.Width / 2
            'Me.SplitContainer5.Panel2.Width = Me.SplitContainer5.Width / 2

            '''@rosadinnn
            Me.SplitContainer6.Width = Me.cboPurchaseOrder.Width

            '''@rosadinnn
            Me.cboCodeNo.Width = Me.SplitContainer6.Panel1.Width - 5 - Me.cboCodeNo.Left
            Me.TxtOrderBalance.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.TxtOrderBalance.Left

            '''@rosadinnn
            '''@MUaT
            ''Me.txtDesc.Width = Me.cboPurchaseOrder.Width

            Me.SplitContainer3.TabStop = False
            Me.SplitContainer5.TabStop = False
            'Me.SplitContainer3.Width = Me.SplitContainer5.Panel1.Width
            Me.btnFindLocation.Left = Me.BtnFindCodeNo.Left
            Me.TxtLocation.Width = Me.TxtCodeNo.Width
            Me.txtBatchNo.Width = Me.SplitContainer3.Panel1.Width - 5 - Me.txtBatchNo.Left
            Me.DtpExpired.Width = Me.SplitContainer3.Panel2.Width - 5 - Me.DtpExpired.Left
            Me.btnFindCurrency2.Left = Me.SplitContainer5.Panel2.Width - Me.btnFindCurrency2.Width - 5
            Me.TxtCurrencyCode1.Width = Me.txtAmount.Width - 5 - Me.btnFindCurrency2.Width


            Me.BtnFindCodeNo.Left = Me.SplitContainer5.Panel1.Width - 5 - Me.BtnFindCodeNo.Width
            Me.TxtCodeNo.Width = Me.BtnFindCodeNo.Left - 5 - Me.TxtCodeNo.Left
            Me.txtDesc.Width = Me.SplitContainer5.Panel1.Width - 5 - Me.txtDesc.Left
            Me.btnFindLocation.Left = Me.BtnFindCodeNo.Left
            Me.btnFindCurrency2.Left = Me.BtnFindCodeNo.Left
            Me.btnFindCurrency2.Width = Me.BtnFindCodeNo.Width
            Me.TxtCurrencyCode1.Width = Me.TxtCodeNo.Width
            'Me.CboUnitMulti.Left = Me.btnInput.Left + Me.btnInput.Width - Me.CboUnitMulti.Width

            Me.dtpDate.Left = Me.SplitContainer7.Panel1.Width - Me.dtpDate.Width
            Me.Label3.Left = Me.dtpDate.Left - Me.Label3.Width - 5
            Me.TxtAOP.Width = Me.dtpDate.Left - Me.TxtAOP.Left - 5
            Me.TxtSupplierName.Width = Me.dtpDate.Left + Me.dtpDate.Width - Me.TxtSupplierName.Left
            Me.llbAddress.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.llbAddress.Width
            Me.cboSupplierAddress.Width = Me.llbAddress.Left - 5 - Me.cboSupplierAddress.Left
            Me.txtSupplierInvoiceNo.Width = Me.Label14.Left - 5 - Me.txtSupplierInvoiceNo.Left
            Me.TxtNPWPNo.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.TxtNPWPNo.Width
            Me.Label13.Left = Me.TxtNPWPNo.Left - 5 - Me.Label13.Width
            Me.TxtTaxInvoiceNumber.Width = Me.Label13.Left - 5 - Me.TxtTaxInvoiceNumber.Left
            Me.cboCalculateBy.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.cboCalculateBy.Width
            Me.Label20.Left = Me.cboCalculateBy.Left - 5 - Me.Label20.Width
            Me.txtCurrencyCode.Anchor = Me.TxtTaxInvoiceNumber.Anchor
            Me.btnFindCurrency.Anchor = Me.cboCalculateBy.Anchor
            Me.btnFindCurrency.Left = Me.TxtTaxInvoiceNumber.Left + Me.TxtTaxInvoiceNumber.Width - Me.btnFindCurrency.Width
            Me.txtCurrencyCode.Width = Me.btnFindCurrency.Left - 5 - Me.txtCurrencyCode.Left
            Me.cboAccount.Left = Me.txtCurrencyCode.Left
            'Me.cboPrimaryAccount.Visible = False
            Me.cboAccount.Width = Me.dtpDate.Left + Me.dtpDate.Width - Me.cboAccount.Left

            Me.txtSubTotalBeforeDP.Width = Me.SplitContainer7.Panel2.Width - Me.txtSubTotalBeforeDP.Left - 5
            Me.txtSubTotalAfterDP.Width = Me.txtSubTotalBeforeDP.Width
            Me.TxtDutyStamp.Width = Me.txtSubTotalBeforeDP.Left + Me.txtSubTotalBeforeDP.Width - Me.TxtDutyStamp.Left
            Me.txtTaxAmount.Width = Me.TxtDutyStamp.Width - Me.cboPPNDipungut.Width - 5
            Me.txtTaxAmount.Left = Me.TxtDutyStamp.Right - Me.txtTaxAmount.Width
            Me.txtTaxExchangeRate.Width = Me.TxtDutyStamp.Width - Me.txtTax.Width - Me.Label17.Width - Me.Label25.Width - 20
            Me.txtTaxExchangeRate.Left = Me.txtSubTotalBeforeDP.Left + Me.txtSubTotalBeforeDP.Width - Me.txtTaxExchangeRate.Width
            Me.TxtTotal.Width = Me.TxtDutyStamp.Width
            Me.Label7.Left = Me.txtSubTotalBeforeDP.Left + Me.txtSubTotalBeforeDP.Width - Me.Label7.Width
            Me.TxtPaymentsTerms.Width = Me.Label7.Left - 5 - Me.TxtPaymentsTerms.Left
            Me.TxtRemark.Width = Me.TxtTotal.Width
            Me.btnFindLocation.Left = Me.BtnFindCodeNo.Left
            Me.btnFindLocation.Width = Me.BtnFindCodeNo.Width
            Me.TxtLocation.Width = Me.TxtCodeNo.Width



            If Me.visibleCBO = False Then
                Me.CboUnitMulti.SendToBack()
                Me.CboUnitMulti.Visible = False
            End If

            Me.ColDiscAmount.Visible = True
            Me.ColDiscAmount.DisplayIndex = Me.ColDiscount.DisplayIndex + 1
            Me.lblQty.Left = Me.SplitContainer5.Panel1.Left + 15
            Me.lblDiscPersen.Left = Me.lblQty.Left
            Me.lblNettPrice.Left = Me.lblQty.Left
            Me.lblremarks.Left = Me.lblQty.Left

            Me.TxtDiscount.Width = (Me.SplitContainer5.Panel1.Width / 2) - Me.lblDiscPersen.Width - Me.lblDiscPersen.Left - 5
            Me.TxtUnitPrice.Width = Me.TxtDiscount.Width
            Me.txtNettPrice.Width = Me.TxtDiscount.Width
            Me.txtAmount.Width = Me.TxtDiscount.Width
            Me.TxtDiscAmount.Width = Me.TxtDiscount.Width

            Me.TxtQtyMulti.Left = Me.lblDiscPersen.Right + 5


            If Me.visibleCBO = True Then
                Me.TxtQtyMulti.Width = Me.TxtDiscount.Width - Me.CboUnitMulti.Width - 5
                Me.CboUnitMulti.Left = Me.TxtDiscount.Right - Me.CboUnitMulti.Width
            Else
                Me.TxtQtyMulti.Width = Me.TxtDiscount.Width - Me.LblUnitsMulti.Width - 5
                Me.LblUnitsMulti.Left = Me.TxtDiscount.Right - Me.LblUnitsMulti.Width
            End If

            Me.TxtDiscount.Left = Me.TxtQtyMulti.Left
            Me.txtNettPrice.Left = Me.TxtQtyMulti.Left
            Me.TxtRemarkItem.Left = Me.TxtQtyMulti.Left
            'Me.CboUnitMulti.Left = Me.TxtDiscount.Right - Me.CboUnitMulti.Width - Me.LblUnitsMulti.Width

            'Me.LblUnitsMulti.SendToBack() 


            Me.LblUnitPrice.Left = Me.TxtDiscount.Right + 10
            Me.lblDiscAmount.Left = Me.LblUnitPrice.Left
            Me.lblAmount.Left = Me.LblUnitPrice.Left

            Me.cboPriceByMultiUnit.Left = Me.LblUnitPrice.Right + 5
            Me.TxtUnitPrice.Left = Me.cboPriceByMultiUnit.Right + 5
            Me.TxtUnitPrice.Width = Me.SplitContainer5.Panel2.Width - 5 - Me.TxtUnitPrice.Left
            Me.txtAmount.Left = Me.TxtUnitPrice.Left
            Me.txtAmount.Width = Me.TxtUnitPrice.Width
            Me.TxtDiscAmount.Left = Me.TxtUnitPrice.Left
            Me.TxtDiscAmount.Width = Me.TxtUnitPrice.Width

            Me.btnInput.Left = Me.TxtUnitPrice.Right - Me.btnInput.Width
            Me.TxtRemarkItem.Width = Me.btnInput.Left - Me.TxtRemarkItem.Left - 5
            'Me.btnInput.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.btnInput.Width

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpExpired_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpExpired.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.DtpExpired.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        Me.txtBatchNo.Focus()
                    Else
                        If Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        ElseIf Me.TxtCurrencyCode1.ReadOnly = False Then
                            Me.TxtCurrencyCode1.Focus()
                        Else
                            Me.TxtUnitPrice.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub PrepareMultiUnit(ByVal vPrimaryProduct As Double)
        Dim dtaMultiUnit As New DseTransaksiPOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Try
            Me.CboUnitMulti.DataSource = Nothing
            If Me.ObjPO.IsMultiUnitProduct(vPrimaryProduct) = False Then
                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtPrimaryUnitsMulti.Text = ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
                Me.TxtPrimaryUnitsDefault.Text = Me.TxtPrimaryUnitsMulti.Text
                Me.LblUnitsMulti.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsMulti.Text)
                Me.LblUnitsDefault.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Visible = True
                Me.CboUnitMulti.Visible = False
                Me.CboUnitMulti.TabStop = False
                Me.LblUnitPrice.Text = "Unit Price/" & LblUnitsMulti.Text
                Me.SetQtyDefault()

                If Me.TxtPrimaryUnitsDefault.Text = "" Or Me.TxtPrimaryUnitsDefault.Text = "0" Then
                    Me.TxtPrimaryUnitsDefault.Text = "-1"
                End If
                If Me.TxtPrimaryUnitsMulti.Text = "" Or Me.TxtPrimaryUnitsMulti.Text = "0" Then
                    Me.TxtPrimaryUnitsMulti.Text = "-1"
                End If
                If Me.LblUnitsMulti.Text = "" Then
                    Me.LblUnitsMulti.Text = "Units"
                End If
                If Me.LblUnitsDefault.Text = "" Then
                    Me.LblUnitsDefault.Text = "Units"
                End If
            Else
                Me.LblUnitsMulti.Visible = False
                Me.CboUnitMulti.Visible = True
                Me.CboUnitMulti.TabStop = True
                dtaMultiUnit.Connection.ConnectionString = Core_Module.GetConstring
                dtaMultiUnit.FillBy_PrimaryProduct(Me.tblMultiUnit, Me.TxtPrimaryCodeNo.Text)
                Me.ObjPO.IsiCboUnit(vPrimaryProduct, Me.CboUnitMulti, Me.tblCboUnit, Me.tblMultiUnit, Me.TxtPrimaryUnitsMulti)
                Me.CboUnitMulti.SelectedIndex = -1
                Me.CboUnitMulti.SelectedIndex = 0
            End If

            setLocation_Product()
            Me.SetCboPriceByMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCboPriceByMultiUnit()
        Dim UnitCode1 As String = ""
        Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal

        Try
            Me.cboPriceByMultiUnit.Items.Clear()

            If Me.CboUnitMulti.Items.Count = 0 Then
                'tanpa multi unit
                Me.LblUnitPrice.Text = "Unit Price / " + Me.LblUnitsDefault.Text
                Me.cboPriceByMultiUnit.Visible = False
                Me.cboPriceByMultiUnit.TabStop = False
                Me.cboPriceByMultiUnit.Items.Add(Me.LblUnitPrice.Text)
                Me.cboPriceByMultiUnit.SelectedIndex = 0
                Me.txtPriceByMultiUnit.Text = "1"

                Me.ObjPO.GetHistoricalPrice(Me.TxtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.txtPrimaryCurrency.Text, UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc, Me.txtPriceByMultiUnit.Text)

                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.DefaultPrecision)
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, Me.DefaultPrecision)
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(DiscAmount, Me.DefaultPrecision)
                Me.TxtDiscFromPerct.Text = DiscFromPerc
            Else
                If Me.CboUnitMulti.Items.Count = 1 Then
                    Me.LblUnitPrice.Text = "Unit Price / " + Me.LblUnitsDefault.Text
                    Me.cboPriceByMultiUnit.Visible = False
                    Me.cboPriceByMultiUnit.TabStop = False
                    Me.cboPriceByMultiUnit.Items.Add(Me.LblUnitPrice.Text)
                    Me.cboPriceByMultiUnit.SelectedIndex = 0
                    Me.txtPriceByMultiUnit.Text = "1"
                Else
                    Me.LblUnitPrice.Text = "Unit Price"
                    UnitCode1 = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)

                    Me.cboPriceByMultiUnit.Visible = True
                    Me.cboPriceByMultiUnit.TabStop = True
                    Me.cboPriceByMultiUnit.Items.Add(UnitCode1)
                    If Me.CboUnitMulti.Text <> UnitCode1 Then
                        Me.cboPriceByMultiUnit.Items.Add(Me.CboUnitMulti.Text)
                        Me.cboPriceByMultiUnit.SelectedIndex = Me.ObjProduct.GetPriceByMultiUnit(Me.TxtPrimaryCodeNo.Text)
                    Else
                        Me.cboPriceByMultiUnit.SelectedIndex = 0
                    End If
                    'Me.txtPriceByMultiUnit.Text = "1"
                    If Me.CboUnitMulti.Text <> UnitCode1 Then
                        Me.txtPriceByMultiUnit.Text = "1"
                    Else
                        Me.txtPriceByMultiUnit.Text = "0"
                    End If
                    Me.cboPriceByMultiUnit.Left = Me.LblUnitPrice.Left + Me.LblUnitPrice.Width
                End If
            End If

            Me.Hitung_Amount2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboUnitMulti_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnitMulti.Leave
        Try
            If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                ''@rosadinnn
                ''MUaT
                Dim DefPrec_ As Integer = 0
                If Me.ObjPO.IsMultiUnitProduct(Me.TxtPrimaryCodeNo.Text) = False Then
                    DefPrec_ = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                Else
                    DefPrec_ = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
                End If

                Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, DefPrec_)
                Me.SetQtyDefault()

                If Me.CboUnitMulti.Text <> Me.cboPriceByMultiUnit.Text Then
                    Me.txtPriceByMultiUnit.Text = 0
                Else
                    Me.txtPriceByMultiUnit.Text = 1
                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboUnitMulti_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboUnitMulti.SelectedIndexChanged
        'Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Dim row As DseTransaksiPO.CboUnitRow
        Try
            If Me.CboUnitMulti.SelectedIndex = -1 Then
                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtPrimaryUnitsMulti.Text = ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
                Me.TxtPrimaryUnitsDefault.Text = Me.TxtPrimaryUnitsMulti.Text
                Me.LblUnitsMulti.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsMulti.Text)
                Me.LblUnitsDefault.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Visible = True
                Me.CboUnitMulti.Visible = False
                Me.CboUnitMulti.TabStop = False
                Me.LblUnitPrice.Text = "Unit Price/" & LblUnitsMulti.Text

                If Me.TxtPrimaryUnitsDefault.Text = "" Or Me.TxtPrimaryUnitsDefault.Text = "0" Then
                    Me.TxtPrimaryUnitsDefault.Text = "-1"
                End If
                If Me.TxtPrimaryUnitsMulti.Text = "" Or Me.TxtPrimaryUnitsMulti.Text = "0" Then
                    Me.TxtPrimaryUnitsMulti.Text = "-1"
                End If
                If Me.LblUnitsMulti.Text = "" Then
                    Me.LblUnitsMulti.Text = "Units"
                End If
                If Me.LblUnitsDefault.Text = "" Then
                    Me.LblUnitsDefault.Text = "Units"
                End If
            Else
                row = Me.tblCboUnit.DefaultView(Me.CboUnitMulti.SelectedIndex).Row
                Me.TxtRatioReference.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, row.Primary), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, row.Primary))
                'Me.TxtRatioDefault.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, row.Primary), ObjProduct.GetAngkaDecimal(row.Primary))
                Me.TxtRatioDefault.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, row.Primary), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, row.Primary))
                Me.TxtPrimaryUnitsMulti.Text = row.Primary
                Me.TxtQtyMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                Me.TxtPrimaryUnitsDefault.Text = ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
                Me.LblUnitsMulti.Text = Me.CboUnitMulti.Text
                Me.LblUnitsDefault.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Visible = False
                Me.CboUnitMulti.Visible = True
                Me.CboUnitMulti.TabStop = True
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnitMulti.Text
            End If
            Me.SetQtyDefault()
            Me.SetCboPriceByMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIdx).Value) * _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value)), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetRatioAndQtyDefault(ByVal vDgvRow As DataGridViewRow, ByVal vSetQtyDefault As Boolean)
        Try
            With vDgvRow
                If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnit.Index).Value Then
                    .Cells(Me.ColRatioDesc.Index).Value = .Cells(Me.ColRatioReference.Index).Value & " " & .Cells(Me.ColUnitsCodeMulti.Index).Value & " = " & _
                    .Cells(Me.ColRatioDefault.Index).Value & " " & .Cells(Me.ColUnit.Index).Value
                Else
                    .Cells(Me.ColRatioDesc.Index).Value = "-"
                End If

                If vSetQtyDefault = True Then
                    Me.SetQtyDefaultOnDgv(vDgvRow.Index)
                End If
            End With
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


    Private Sub llbAdjustDPP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAdjustDPP.LinkClicked
        Dim ij As New FrmTransaksiAOP_AdjustDPP2
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
            ij.dgvItem.Columns(ij.ColAtAmount.Index).Visible = False
            ij.dgvItem.Columns(ij.ColShowDetail.Index).Visible = False
            ij.dgvItem.Columns(ij.ColItemAdjustment.Index).Visible = True
            ij.dgvItem.Columns(ij.ColItemAdjustmentReason.Index).Visible = True
            ij.dgvItem.Columns(ij.ColItemAdjustmentUser.Index).Visible = True
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

    Private Sub FrmTransaksiAOP_Quick2_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        Try
            Me.SetLocation()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtDiscAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Enter
        Try
            Me.buffer2 = Me.TxtDiscAmount.Text
            Me.TxtDiscFromPerct.Text = "0"
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscAmount, sender, e, Me.TxtUnitPrice.Text)
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

    Private Sub TxtDiscAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.GotFocus
        If Me.TxtDiscAmount.ReadOnly = False And Me.TxtDiscAmount.Enabled = True Then
            TxtDiscAmount.BackColor = BackColorAll
        Else
            TxtDiscAmount.BackColor = Me.txtPrimaryAccount.BackColor
        End If
    End Sub

    Private Sub TxtDiscAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.LostFocus
        If Me.TxtDiscAmount.ReadOnly = False And Me.TxtDiscAmount.Enabled = True Then
            TxtDiscAmount.BackColor = Me.txtPrimaryAccount.BackColor
        End If
    End Sub

    Private Sub TxtDiscAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscAmount.Text, Me.DefaultPrecision)
                If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtDiscAmount.Text = Me.buffer2
                    Exit Sub
                End If
                If CDec(Me.TxtDiscAmount.Text) <> CDec(Me.buffer2) Then
                    'If IsNumeric(Me.TxtUnitPrice.Text) AndAlso CDec(Me.TxtUnitPrice.Text) <> 0 Then
                    '    Me.TxtDiscount.Text = CDec(CDec(Me.TxtDiscAmount.Text) / CDec(Me.TxtUnitPrice.Text)) * 100
                    'Else
                    '    Me.TxtDiscount.Text = 0
                    'End If
                    Me.TxtDiscFromPerct.Text = "0"
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            'gabungkan baris
            Me.GabungkanBaris(Me.IndexAsal, Me.IndexTujuan)
            'hitung ulang total
            Me.Hitung_Total_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_AOPQuick_Create(ByVal vPrimaryTransaksi As Double, ByVal vNoAOP_Obs As String)

        Try
            Me.baru = False
            Me.ObjAOP.ResetObject(Me.dgvItem, Me.DGVdp, Me.dgvExchageRates, Me.DGVitemPR, Me.dgvCost)
            Me.ObjAOP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP

            Me.ObjAOP.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut)
            Me.ObjAOP.NoTransaksi = vNoAOP_Obs
            Me.UserCode = Me.ObjAOP.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_AOPQuick_Save(ByVal vNoPO As String, ByVal vNoSIP As String) As Double

        Try
            Me.baru = True
            Me.ObjAOP.SimpanBaru1(Me.dgvItem, Me.DGVdp, Me.DGVitemPR, Me.dgvExchageRates, Me.UserCode, Me.dgvCost, Me, Me.TblCustomDummy, Me.TblItemCustomInOut, False, vNoPO, vNoSIP, "", False, True)
            Return Me.ObjAOP.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    '
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(LinkLabel1, New Point(0, 15))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub



    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim ObjCustIdentifier As New CustomSMESS.ClsIdentifier
        Try
            If ObjCustIdentifier.GetIdentifier = "smess_admiral2" Or ObjCustIdentifier.GetIdentifier = "smess_linda" Then
                Me.LoadCustomForAdmiral(sender, e)
            Else
                Me.LoadCustomForBinaya(sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadCustomForBinaya(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Frm As New CustomSMESS.FrmTransaksiAOP_Quick2_AddItembyCustom
        Dim r2 As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            Frm.PrimaryTransaksiAOP = Me.TxtPrimaryTransaksiAOP.Text
            Frm.PrimarySupplier = Me.txtPrimarySupplier.Text
            Frm.visibleCBO = Me.visibleCBO
            Frm.DefaultPrecision = Me.DefaultPrecision
            Frm.ExchangeRate = Me.txtTaxExchangeRate.Text
            Frm.baru = Me.baru
            Frm.SupplierCode = Me.txtSupplierCode.Text
            Frm.TanggalTransaksi = Me.dtpDate.Value.Date
            Frm.PrimaryCurrency = Me.txtPrimaryCurrency.Text
            Frm.cboCalculateBy.SelectedIndex = Me.cboCalculateBy.SelectedIndex
            Frm.dgvCost.Rows.Clear()
            For Each r As DataGridViewRow In Me.dgvCost.Rows
                r2 = Frm.dgvCost.Rows(Frm.dgvCost.Rows.Add)
                r2.Cells(0).Value = r.Cells(0).Value
                r2.Cells(1).Value = r.Cells(1).Value
                r2.Cells(2).Value = r.Cells(2).Value
                r2.Cells(3).Value = r.Cells(3).Value
                r2.Cells(4).Value = r.Cells(4).Value
                r2.Cells(5).Value = r.Cells(5).Value
                r2.Cells(6).Value = r.Cells(6).Value
                r2.Cells(7).Value = r.Cells(7).Value
                r2.Cells(8).Value = r.Cells(8).Value
                r2.Cells(9).Value = r.Cells(9).Value
                r2.Cells(10).Value = r.Cells(10).Value
            Next

            Frm.dgvExchageRates.Rows.Clear()
            For Each r As DataGridViewRow In Me.dgvExchageRates.Rows
                r2 = Frm.dgvExchageRates.Rows(Frm.dgvExchageRates.Rows.Add)
                r2.Cells(0).Value = r.Cells(0).Value
                r2.Cells(1).Value = r.Cells(1).Value
                r2.Cells(2).Value = r.Cells(2).Value
                r2.Cells(3).Value = r.Cells(3).Value
                r2.Cells(4).Value = r.Cells(4).Value
                r2.Cells(5).Value = r.Cells(5).Value
                r2.Cells(6).Value = r.Cells(6).Value
            Next
            If btnOK.Enabled = False Then
                Frm.AktifForm = False
            End If
            Frm.TblDummy = Me.TblCustomDummy
            If Frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblCustomDummy = Frm.TblDummy
                Me.IsikanDgvFromTblDummyBinaya(sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Public Sub LoadCustomForAdmiral(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '###TR
        Dim i As New CustomSMESS.FrmTransaksiAOP_Quick2_AddByCustomProduct
        Dim primaryCurr As Double = -1
        Dim primaryCurrTax As Double = -1
        Dim primaryTransaksiAOP As Double = -1
        Dim PrimarySupplier As Double = -1
        Dim primaryPO As Double = -1

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtPrimaryCurrency.Text <> "" Then
                primaryCurr = Me.txtPrimaryCurrency.Text
            End If
            If Me.txtPrimaryCurrTax.Text <> "" Then
                primaryCurrTax = Me.txtPrimaryCurrTax.Text
            End If
            If Me.TxtPrimaryTransaksiAOP.Text <> "" Then
                primaryTransaksiAOP = Me.TxtPrimaryTransaksiAOP.Text
            End If
            If Me.txtPrimarySupplier.Text <> "" Then
                PrimarySupplier = Me.txtPrimarySupplier.Text
            End If
            If Me.txtPrimaryPO.Text <> "" Then
                primaryPO = Me.txtPrimaryPO.Text
            End If
            i.NoAOP = Me.TxtAOP.Text.ToString
            i.CurrencyCode = Me.txtCurrencyCode.Text
            i.PrimaryCurrency = primaryCurr
            i.PrimaryCurrTax = primaryCurrTax
            i.PrimaryTransaksiAOP = primaryTransaksiAOP
            i.PrimarySupplier = PrimarySupplier
            i.PrimaryPO = primaryPO
            i.UserCode = Me.UserCode
            i.SupplierCode = Me.txtSupplierCode.Text
            i.baru = Me.baru
            i.TanggalTransaksi = Me.dtpDate.Value
            i.ExchangeRate = Me.txtTaxExchangeRate.Text
            i.TblItemCustom = Me.TblCustomDummy.Copy
            i.CalculateByIndex = Me.cboCalculateBy.SelectedIndex

            If Me.TxtRemark.Enabled = False Then
                i.AktifForm = False
            End If

            If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblCustomDummy.Clear()
                Me.TblItemCustomInOut.Clear()
                Me.TblCustomDummy = i.TblItemCustom.Copy
                Me.TblItemCustomInOut = i.TblItemCustomInOut.Copy
                Me.IsikanDgvFromTblDummyAdmiral(sender, e)
                Me.Hitung_AtCost_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsikanDgvFromTblDummyBinaya(ByVal sender As System.Object, ByVal e As System.EventArgs) As Boolean
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1
        Dim rowDummy As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Try
            Me.dgvItem.Rows.Clear()
            For Each rowDummy In Me.TblCustomDummy.Rows
                indexSama = Me.GetIndex2(rowDummy)
                If indexSama = -1 Then
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product)
                    AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)

                    With row
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColCodeNo.Index).Value = rowDummy.Product_Code
                        .Cells(Me.ColDescription.Index).Value = rowDummy.Description
                        .Cells(Me.ColSuppliersCodeNumber.Index).Value = ""
                        .Cells(Me.ColSuppliersProductDesc.Index).Value = ""
                        .Cells(Me.ColPONO.Index).Value = rowDummy.Primary_PO
                        .Cells(Me.ColSuppDONo.Index).Value = rowDummy.Supplier_Do_No
                        .Cells(Me.ColGRNNO.Index).Value = rowDummy.GRN_No
                        .Cells(Me.ColStockOutNo.Index).Value = ""
                        .Cells(Me.ColLotNumber.Index).Value = rowDummy.BatchNumber
                        If Me.DtpExpired.Visible = False Then
                            .Cells(Me.ColExpiry.Index).Value = #1/1/1950#
                        Else
                            .Cells(Me.ColExpiry.Index).Value = rowDummy.Expiry_Date
                        End If
                        .Cells(Me.ColLocation.Index).Value = rowDummy.Storage_Area_Code
                        .Cells(Me.ColQtyPerMaster.Index).Value = rowDummy.Qty_Per_Master
                        .Cells(Me.ColQtyMaster.Index).Value = rowDummy.Qty_Master
                        .Cells(Me.ColMasterUnits.Index).Value = rowDummy.Master_Unit
                        .Cells(Me.ColQty.Index).Value = rowDummy.Qty
                        .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Disc_Amount), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscFromPerct.Index).Value = rowDummy.Disc_From_Perct

                        .Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Total_Qty, AngkaDecimal)
                        .Cells(Me.ColUnit.Index).Value = rowDummy.Units

                        .Cells(Me.ColSecQty.Index).Value = rowDummy.Secondary_Qty
                        .Cells(Me.ColSecUnit.Index).Value = rowDummy.Secondary_Unit
                        .Cells(Me.ColCurrency.Index).Value = rowDummy.Currency_Code
                        .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Unit_Price), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscount.Index).Value = CDec(rowDummy.Discount)
                        .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Nett_Price), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Amount), AngkaDecimalCurrency)
                        .Cells(Me.ColRemarksItem.Index).Value = rowDummy.Remarks
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowDummy.Primary_Location
                        .Cells(Me.ColPrimaryGRN.Index).Value = rowDummy.Primary_GRN
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDummy.Primary_Product
                        .Cells(Me.ColPrimaryPO.Index).Value = rowDummy.Primary_PO
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = rowDummy.Primary_Master_Unit
                        .Cells(Me.ColPrimaryUnit.Index).Value = rowDummy.Primary_Units
                        .Cells(Me.ColPrimarySecUnit.Index).Value = rowDummy.Primary_Secondary_Units
                        .Cells(Me.ColPrimaryCurrency.Index).Value = rowDummy.Primary_Currency
                        .Cells(Me.ColShowDetail.Index).Value = "Show"
                        .Cells(Me.ColReferedToPO.Index).Value = rowDummy.Refered_To_PO
                        '<i> diblock karna if nya berlaku saat input di custom 
                        'If rowDummy.Primary_PO <> -1 Then
                        '    .Cells(Me.ColReferedToPO.Index).Value = 1
                        'Else
                        '    .Cells(Me.ColReferedToPO.Index).Value = 0
                        'End If
                        .Cells(Me.ColItemAdjustment.Index).Value = 0
                        .Cells(Me.ColItemAdjustmentReason.Index).Value = ""
                        .Cells(Me.ColItemAdjustmentUser.Index).Value = ""
                        .Cells(Me.ColItemAdjustmentTimestamp.Index).Value = "-"

                        .Cells(Me.ColTotalQtyMulti.Index).Value = rowDummy.Total_Qty_Multi
                        .Cells(Me.ColUnitsCodeMulti.Index).Value = rowDummy.Unit_Code_Multi
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = rowDummy.Primary_Units_Multi
                        .Cells(Me.ColRatioReference.Index).Value = rowDummy.Ratio_Reference
                        .Cells(Me.ColRatioDefault.Index).Value = rowDummy.Ratio_Default
                        If rowDummy.Primary_Units <> rowDummy.Primary_Units_Multi Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Ratio_Reference, _
                                    ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Units_Multi)) _
                                    & " " & rowDummy.Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(rowDummy.Ratio_Default, _
                                    Me.ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product)) & " " & rowDummy.Units
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If

                    End With
                Else
                    row = Me.dgvItem.Rows(indexSama)
                    'tambahkan qty, 

                    row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(rowDummy.Total_Qty_Multi), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                    Me.SetQtyDefaultOnDgv(indexSama)

                    'hitung ulang amount
                    row = Me.dgvItem.Rows(indexSama)
                    Me.Hitung_Amount_Multi(row)
                End If

                PrimaryCurrenct = Me.TxtPrimaryCurrency1.Text

                'Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
                'Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)


                'Me.cboCalculateBy.SelectedIndex = -1
                'Me.cboCalculateBy.SelectedIndex = 0

                ''fokuskan ke baris terakhir
                'While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                '        Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                '    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                'End While

                Me.Hitung_Total_Multi()

                If Me.chkPONumber.Checked = True Then
                    'perhitungan menggunakan nilai jumlah dgv yang baru
                    Dim JmlObjLama, JmlPOBBaru, JmlDGV, JmlPOB As Decimal
                    JmlDGV = Me.GetJmlDGV_(rowDummy.Primary_Product, rowDummy.Primary_PO)
                    JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(rowDummy.Primary_PO, rowDummy.Primary_Product)
                    JmlObjLama = Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, rowDummy.Primary_Product, rowDummy.Primary_PO)
                    JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV

                    If JmlPOBBaru <= 0 Then
                        'remove row, pada code no ybs
                        Dim rowView As Data.DataRowView = Me.cboCodeNo.SelectedItem
                        Dim row1 As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row = rowView.Row
                        Me.TBL_PO3.Rows.Remove(row1)
                        Me.cboCodeNo_SelectedValueChanged(sender, e)
                        Me.cboCodeNo.Focus()
                        If Me.cboCodeNo.Items.Count = 0 Then
                            Me.IsikanCBOPO()
                            Me.cboPurchaseOrder.Focus()
                        End If
                    Else
                        If Me.cboCodeNo.SelectedIndex <> Me.cboCodeNo.Items.Count - 1 Then
                            Me.cboCodeNo.SelectedIndex += 1
                            Me.cboCodeNo.Focus()
                        Else
                            'focus pada cbo code no
                            Me.cboCodeNo.Focus()
                            'hitung kembali order balance dan sisa free qty
                            Me.SetOrderBalance()
                        End If
                    End If
                End If
            Next
            Me.FormatDGV2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex2(ByVal rowDummy As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row) As Integer
        Dim BatchNo As String = ""
        Dim PrimaryLocation As Double = -1
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                BatchNo = row.Cells(Me.ColLotNumber.Index).Value

                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = rowDummy.Primary_Product And BatchNo = rowDummy.BatchNumber And _
                CDec(rowDummy.Unit_Price) = CDec(row.Cells(Me.ColUnitPrice.Index).Value) And _
                CDec(rowDummy.Discount) = CDec(row.Cells(Me.ColDiscount.Index).Value) And _
                CDec(rowDummy.Primary_Location) = row.Cells(Me.ColPrimaryLocation.Index).Value Then
                    If Me.chkPONumber.Checked = False Then
                        If row.Cells(Me.ColReferedToPO.Index).Value = 0 Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColReferedToPO.Index).Value = 1 And row.Cells(Me.ColPrimaryPO.Index).Value = rowDummy.Primary_PO Then
                            Return row.Index
                        End If
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetOrderBalance2()
        Try
            'berdasarkan po number
            If IsNumeric(Me.TxtPrimaryCodeNo.Text) = True And IsNumeric(Me.txtPrimaryPO.Text) = True Then
                Me.TxtOrderBalance.Text = Me.ObjCD.DigitGrouping(Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) _
                + Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) - _
                Me.GetJmlDGV_(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        ''###TR
        'Dim i As New CustomSMESS.FrmTransaksiAOP_Quick2_AddByCustomProduct
        'Dim primaryCurr As Double = -1
        'Dim primaryCurrTax As Double = -1
        'Dim primaryTransaksiAOP As Double = -1
        'Dim PrimarySupplier As Double = -1
        'Dim primaryPO As Double = -1

        'Try
        '    i.BackColorAll = Me.BackColorAll
        '    If Me.txtPrimaryCurrency.Text <> "" Then
        '        primaryCurr = Me.txtPrimaryCurrency.Text
        '    End If
        '    If Me.txtPrimaryCurrTax.Text <> "" Then
        '        primaryCurrTax = Me.txtPrimaryCurrTax.Text
        '    End If
        '    If Me.TxtPrimaryTransaksiAOP.Text <> "" Then
        '        primaryTransaksiAOP = Me.TxtPrimaryTransaksiAOP.Text
        '    End If
        '    If Me.txtPrimarySupplier.Text <> "" Then
        '        PrimarySupplier = Me.txtPrimarySupplier.Text
        '    End If
        '    If Me.txtPrimaryPO.Text <> "" Then
        '        primaryPO = Me.txtPrimaryPO.Text
        '    End If
        '    i.NoAOP = Me.TxtAOP.Text.ToString
        '    i.CurrencyCode = Me.txtCurrencyCode.Text
        '    i.PrimaryCurrency = primaryCurr
        '    i.PrimaryCurrTax = primaryCurrTax
        '    i.PrimaryTransaksiAOP = primaryTransaksiAOP
        '    i.PrimarySupplier = PrimarySupplier
        '    i.PrimaryPO = primaryPO
        '    i.UserCode = Me.UserCode
        '    i.SupplierCode = Me.txtSupplierCode.Text
        '    i.baru = Me.baru
        '    i.TanggalTransaksi = Me.dtpDate.Value
        '    i.TaxExchageRate = Me.txtTaxExchangeRate.Text
        '    i.TblItemCustom = Me.TblCustomDummy.Copy
        '    i.CalculateByIndex = Me.cboCalculateBy.SelectedIndex



        '    If Me.TxtRemark.Enabled = False Then
        '        i.CanEditForm = False
        '    End If

        '    If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
        '        Me.TblCustomDummy.Clear()
        '        Me.TblItemCustomInOut.Clear()
        '        Me.TblCustomDummy = i.TblItemCustom.Copy
        '        Me.TblItemCustomInOut = i.TblItemCustomInOut.Copy
        '        Me.IsikanDgvFromTblDummy(sender, e)
        '        Me.Hitung_AtCost_Multi()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub cboPriceByMultiUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPriceByMultiUnit.SelectedIndexChanged
        Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Dim row As DseTransaksiPO.CboUnitRow

        Try
            If Me.cboPriceByMultiUnit.SelectedIndex <> -1 Then
                If Me.TxtRemarkItem.Enabled = True Then
                    If Me.cboPriceByMultiUnit.Text <> Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) Then
                        Me.txtPriceByMultiUnit.Text = "1"
                    Else
                        Me.txtPriceByMultiUnit.Text = "0"
                    End If

                    If Me.GetHistoricalPrice = True Then
                        Me.ObjPO.GetHistoricalPrice(Me.objUnit.GetPrimary(Me.cboPriceByMultiUnit.Text), Me.TxtPrimaryCodeNo.Text, _
                        Me.txtPrimaryCurrency.Text, UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc, Me.txtPriceByMultiUnit.Text)

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.DefaultPrecision)
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                        Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, Me.DefaultPrecision)
                        Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(DiscAmount, Me.DefaultPrecision)
                        Me.TxtDiscFromPerct.Text = DiscFromPerc
                    End If

                    If IsNumeric(Me.TxtUnitPrice.Text) = True Then
                        If CDec(Me.TxtUnitPrice.Text) > 0 Then
                            Me.Hitung_Amount2()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Sub Atur1()
        Try
            'If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOP_Quick2_AddByCustomProduct") = False Then
            '    Me.GroupBox2.Enabled = True
            '    Me.ToolStripMenuItem1.Visible = False
            '    Me.dgvItem.ReadOnly = False
            '    Me.dgvItem.AllowUserToDeleteRows = True
            'Else
            '    Me.dgvItem.ReadOnly = True
            '    Me.GroupBox2.Enabled = False
            '    Me.ToolStripMenuItem1.Visible = True
            '    Me.dgvItem.AllowUserToDeleteRows = False
            '    Me.LinkLabel1.Enabled = True
            'End If
            Me.ToolStripMenuItem1.Visible = False
            Me.AddItemByCustom.Visible = False
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiAOP_Quick2_AddByCustomProduct") = False Then
                Me.GroupBox2.Enabled = True
                Me.ToolStripMenuItem1.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
                Me.AddItemByCustom.Visible = False
            Else
                Me.dgvItem.ReadOnly = True
                Me.GroupBox2.Enabled = False
                'Me.AddItemByCustom.Visible = True
                Me.ToolStripMenuItem1.Visible = True
                Me.dgvItem.AllowUserToDeleteRows = False
                Me.LinkLabel1.Enabled = True
                Me.AddItemByCustom.Visible = False
            End If

            If Me.baru = True Then
                Me.chkTax.Checked = False
                Me.chkTax.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '##TR
    Public TblCustomDummy As New CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable
    Public TblItemCustomInOut As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

    Private Function IsikanDgvFromTblDummyAdmiral(ByVal sender As System.Object, ByVal e As System.EventArgs) As Boolean
        Dim rowItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1
        Dim PrimaryLocationDummy As Double = -1
        Dim obj As New Core_Module.clsMasterStorageArea
        Dim primaryroll As Double = 22

        Try
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'If Me.ApakahDataItemBenar = True Then
            If Me.txtPrimaryLocation.Text.Trim <> -1 Then
                PrimaryLocationDummy = Me.txtPrimaryLocation.Text
            End If
            For Each rowItemCustom In TblCustomDummy.Rows
                With rowItemCustom
                    indexSama = Me.GetIndex2(rowItemCustom)
                    If indexSama = -1 Then
                        'tambahkan baris baru
                        Me.dgvItem.Rows.Add()
                        row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                        AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(rowItemCustom.Primary_Product)
                        AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(rowItemCustom.Primary_Currency)

                        With row
                            .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                            .Cells(Me.ColCodeNo.Index).Value = rowItemCustom.Product_Code
                            .Cells(Me.ColDescription.Index).Value = rowItemCustom.Description
                            .Cells(Me.ColSuppliersCodeNumber.Index).Value = ""
                            .Cells(Me.ColSuppliersProductDesc.Index).Value = ""
                            .Cells(Me.ColPONO.Index).Value = rowItemCustom.PO_No
                            .Cells(Me.ColSuppDONo.Index).Value = ""
                            .Cells(Me.ColGRNNO.Index).Value = ""
                            .Cells(Me.ColStockOutNo.Index).Value = ""
                            .Cells(Me.ColLotNumber.Index).Value = rowItemCustom.BatchNumber
                            If rowItemCustom.Expiry_Date = CDate("#1/1/1950#") Then
                                .Cells(Me.ColExpiry.Index).Value = #1/1/1950#
                            Else
                                .Cells(Me.ColExpiry.Index).Value = rowItemCustom.Expiry_Date
                            End If
                            .Cells(Me.ColLocation.Index).Value = rowItemCustom.Storage_Area_Code
                            .Cells(Me.ColQtyPerMaster.Index).Value = 0
                            .Cells(Me.ColQtyMaster.Index).Value = 0
                            .Cells(Me.ColMasterUnits.Index).Value = ""
                            .Cells(Me.ColQty.Index).Value = 0
                            .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Disc_Amount, AngkaDecimalCurrency)
                            .Cells(Me.ColDiscFromPerct.Index).Value = rowItemCustom.Disc_From_Perct

                            ''@rosadinnn
                            ''@MUaT
                            .Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Total_Qty, rowItemCustom.Decimal_Product)
                            .Cells(Me.ColUnit.Index).Value = rowItemCustom.Unit_Code

                            .Cells(Me.ColSecQty.Index).Value = -1
                            .Cells(Me.ColSecUnit.Index).Value = ""
                            .Cells(Me.ColCurrency.Index).Value = rowItemCustom.Currency_Code
                            .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Unit_Price, AngkaDecimalCurrency)
                            .Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Discount, AngkaDecimalCurrency)
                            .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Nett_Price, AngkaDecimalCurrency)
                            .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Amount, AngkaDecimalCurrency)
                            .Cells(Me.ColATCost.Index).Value = "0"
                            .Cells(Me.ColAtAmount.Index).Value = "0"
                            .Cells(Me.ColRemarksItem.Index).Value = rowItemCustom.Remarks
                            .Cells(Me.ColPrimaryLocation.Index).Value = rowItemCustom.Primary_Location
                            .Cells(Me.ColPrimaryGRN.Index).Value = rowItemCustom.Primary_GRN
                            .Cells(Me.ColPrimaryCodeNo.Index).Value = rowItemCustom.Primary_Product
                            .Cells(Me.ColPrimaryPO.Index).Value = rowItemCustom.Primary_PO
                            .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                            .Cells(Me.ColPrimaryUnit.Index).Value = rowItemCustom.Primary_Units
                            .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                            .Cells(Me.ColPrimaryCurrency.Index).Value = rowItemCustom.Primary_Currency
                            .Cells(Me.ColShowDetail.Index).Value = "Show"
                            If rowItemCustom.Primary_PO <> "-1" Then
                                .Cells(Me.ColReferedToPO.Index).Value = 1
                            Else
                                .Cells(Me.ColReferedToPO.Index).Value = 0
                            End If
                            .Cells(Me.ColItemAdjustment.Index).Value = 0
                            .Cells(Me.ColItemAdjustmentReason.Index).Value = ""
                            .Cells(Me.ColItemAdjustmentUser.Index).Value = ""
                            .Cells(Me.ColItemAdjustmentTimestamp.Index).Value = "-"

                            .Cells(Me.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Total_Qty_Multi, rowItemCustom.Decimal_Multi)
                            .Cells(Me.ColUnitsCodeMulti.Index).Value = rowItemCustom.Unit_Code_Multi
                            .Cells(Me.ColPrimaryUnitsMulti.Index).Value = rowItemCustom.Primary_Units_Multi
                            .Cells(Me.ColRatioReference.Index).Value = rowItemCustom.Ratio_Reference
                            .Cells(Me.ColRatioDefault.Index).Value = rowItemCustom.Ratio_Default
                            If rowItemCustom.Primary_Units_Multi <> rowItemCustom.Primary_Units Then
                                If rowItemCustom.Is_Kain = False Then
                                    .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(rowItemCustom.Ratio_Reference, _
                                       ObjProduct.GetAngkaDecimal(rowItemCustom.Primary_Product, rowItemCustom.Primary_Units_Multi)) _
                                       & " " & rowItemCustom.Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(rowItemCustom.Ratio_Default, _
                                       Me.ObjProduct.GetAngkaDecimal(rowItemCustom.Primary_Product)) & " " & rowItemCustom.Unit_Code
                                Else
                                    .Cells(Me.ColRatioDesc.Index).Value = .Cells(Me.ColTotalQtyMulti.Index).Value & " " & rowItemCustom.Unit_Code_Multi & " = " & .Cells(Me.ColTotalQty.Index).Value & " " & rowItemCustom.Unit_Code
                                End If
                            Else
                                .Cells(Me.ColRatioDesc.Index).Value = "-"
                            End If

                            .Cells(Me.ColPriceByMultiUnit.Index).Value = rowItemCustom.Price_by_Multi_Unit
                            If rowItemCustom.Price_by_Multi_Unit = "1" Then
                                .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rowItemCustom.Unit_Code_Multi
                            Else
                                .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rowItemCustom.Unit_Code
                            End If
                            If rowItemCustom.Is_Kain = True Then
                                .Cells(Me.ColRatioReference.Index).Value = rowItemCustom.Total_Qty_Multi
                                .Cells(Me.ColRatioDefault.Index).Value = rowItemCustom.Total_Qty
                            End If
                            .Cells(Me.ColReferedToPO.Index).Value = 0
                        End With
                        PrimaryCurrenct = rowItemCustom.Primary_Currency
                    Else

                        row = Me.dgvItem.Rows(indexSama)
                        'tambahkan qty, 

                        ''@rosadinnn
                        ''@MUaT
                        'row.Cells(Me.ColTotalQty.Index).Value = CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(Me.TxtQtyMulti.Text)
                        If rowItemCustom.Primary_Units_Multi <> row.Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                            MsgBox("There is already  " & row.Cells(Me.ColTotalQtyMulti.Index).Value & " " & row.Cells(Me.ColUnit.Index).Value & " of " & rowItemCustom.Product_Code & " exist." & vbNewLine & "You cannot input this item with different units.", MsgBoxStyle.Information, "Information")
                            Exit Function
                        End If
                        row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + rowItemCustom.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))

                        If rowItemCustom.Is_Kain = True Then
                            row.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQty.Index).Value) + rowItemCustom.Total_Qty, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                            row.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQtyMulti.Index).Value, Me.ObjProduct.GetAngkaDecimal(rowItemCustom.Primary_Product, primaryroll)) + _
                            " " + row.Cells(Me.ColUnitsCodeMulti.Index).Value + " = " + Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value, Me.ObjProduct.GetAngkaDecimal(rowItemCustom.Primary_Product)) + " " + row.Cells(Me.ColUnit.Index).Value
                            'row.Cells(Me.ColRatioReference.Index).Value = rowItemCustom.Total_Qty_Multi
                            'row.Cells(Me.ColRatioDefault.Index).Value = rowItemCustom.Total_Qty
                            row.Cells(Me.ColRatioReference.Index).Value = row.Cells(Me.ColTotalQtyMulti.Index).Value
                            row.Cells(Me.ColRatioDefault.Index).Value = row.Cells(Me.ColTotalQty.Index).Value
                        Else
                            Me.SetQtyDefaultOnDgv(indexSama)
                        End If

                        'row.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(Me.TxtQtyDefault.Text), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))



                        'hitung ulang amount
                        row = Me.dgvItem.Rows(indexSama)
                        Me.Hitung_Amount_Multi(row)

                        If String.IsNullOrEmpty(row.Cells(Me.ColRemarksItem.Index).Value) = False Then
                            row.Cells(Me.ColRemarksItem.Index).Value = row.Cells(Me.ColRemarksItem.Index).Value & ", " & .Remarks
                        Else
                            row.Cells(Me.ColRemarksItem.Index).Value = .Remarks
                        End If
                    End If
                End With
            Next



            Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
            Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
            Me.FormatDGV2()

            Me.cboCalculateBy.SelectedIndex = -1
            Me.cboCalculateBy.SelectedIndex = 0

            'fokuskan ke baris terakhir
            If Me.dgvItem.Rows.Count <> 0 Then
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

            End If

            Me.Hitung_Total_Multi()

            If Me.chkPONumber.Checked = True Then
                'perhitungan menggunakan nilai jumlah dgv yang baru
                Dim JmlObjLama, JmlPOBBaru, JmlDGV, JmlPOB As Decimal
                JmlDGV = Me.GetJmlDGV_(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
                JmlObjLama = Me.ObjAOPItem.GetJml(Me.TxtPrimaryTransaksiAOP.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV

                If JmlPOBBaru <= 0 Then
                    'remove row, pada code no ybs
                    Dim rowView As Data.DataRowView = Me.cboCodeNo.SelectedItem
                    Dim row1 As Inventory_Module.DseTransaksiPO.orderbalance_report_view0Row = rowView.Row
                    Me.TBL_PO3.Rows.Remove(row1)
                    Me.cboCodeNo_SelectedValueChanged(sender, e)
                    Me.cboCodeNo.Focus()
                    If Me.cboCodeNo.Items.Count = 0 Then
                        Me.IsikanCBOPO()
                        Me.cboPurchaseOrder.Focus()
                    End If
                Else
                    If Me.cboCodeNo.SelectedIndex <> Me.cboCodeNo.Items.Count - 1 Then
                        Me.cboCodeNo.SelectedIndex += 1
                        Me.cboCodeNo.Focus()
                    Else
                        'focus pada cbo code no
                        Me.cboCodeNo.Focus()
                        'hitung kembali order balance dan sisa free qty
                        Me.SetOrderBalance()
                    End If
                End If

                Me.TxtQtyMulti.Text = "0"
                Me.txtAmount.Text = "0"
                Me.TxtRemarkItem.Text = ""
            Else
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If

            If PrimaryLocationDummy <> -1 Then
                Me.txtPrimaryLocation.Text = PrimaryLocationDummy
                Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Function
    Public Sub UpdateATCOSTTblItemCustom()
        Dim rowItemCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row
        Dim ED As String = ""
        Try
            For i As Integer = 0 To Me.dgvItem.Rows.Count - 1
                If Me.dgvItem.Item(Me.ColExpiry.Index, i).Value = "" Or Me.dgvItem.Item(Me.ColExpiry.Index, i).Value = "-" Then
                    ED = "#1/1/1950#"
                Else
                    ED = Me.dgvItem.Item(Me.ColExpiry.Index, i).Value
                End If
                Me.TblCustomDummy.DefaultView.RowFilter = Nothing
                Me.TblCustomDummy.DefaultView.RowFilter = "[Primary Product] = " & Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, i).Value & " AND [BatchNumber] = '" & _
                Me.dgvItem.Item(Me.ColLotNumber.Index, i).Value & "' AND [Expiry Date] = #" & CDate(ED) & "#"
                For Each dv As Data.DataRowView In Me.TblCustomDummy.DefaultView
                    rowItemCustom = dv.Row
                    rowItemCustom.At_Amount = CDbl(Me.dgvItem.Item(Me.ColAtAmount.Index, i).Value)
                    rowItemCustom.At_Price = CDbl(Me.dgvItem.Item(Me.ColATCost.Index, i).Value)
                    rowItemCustom.User_Code = Me.UserCode
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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