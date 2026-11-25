Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiSalesInvoiceDP
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0

    Dim baru As Boolean
    Dim objInvoiceDP As New ClsTransaksiSalesInvoice
    Dim ObjBBDP As New Core_Module.ClsTransaksiBalanceDP
    Dim ObjCD As New ClsCekData
    Dim ObjSalesPerson As New Core_Module.ClsMasterSalesPerson
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCustomer_sProductCode As New clsTransaksiSettingcustomerProductCodeNo
    Dim objQuotationOutItem As New ClsTransaksiQuotationItem
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCustomer As New clsMasterCustomer
    Dim objInvoiceDPItem As New ClsTransaksiSalesInvoiceItem
    Dim ObjCL As New ClsTransaksiPurchaseClosing
    Dim ObjNumbering As New ClsMasterTransactionNumbering
    Dim ObjTaxInvoiceSetting As New ClsTransaksiSalesInvoice_SettingTaxInvoice
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Dim DefaultPrecision As Integer = 0
    Dim DefaultPrecision_DefCur As Integer = 0
    Dim buffer As String = ""

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Dim DseAccount As New DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccount As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccount As DseMultiAccount.tblisi_cbo_ar_akunRow

    Dim DseCurrTax As New DseMultiAccountTableAdapters.tblCurrencyTaxTableAdapter
    Dim tblCurrTax As New DseMultiAccount.tblCurrencyTaxDataTable
    Dim rowCurrTax As DseMultiAccount.tblCurrencyTaxRow

    Dim DsAccountDP As New DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccountDP As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccountDP As DseMultiAccount.tblisi_cbo_ar_akunRow

    Dim ObjAkun As New Core_Module.ClsMasterAkun1

    '2025 Core Tax Tahap 2.1
    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
    '2025 Core Tax Tahap 2.2
    Dim SPPBDate As Date = #1/1/1950#
    Dim AddInfo As Double = -1
    Dim FacilityStamp As Double = -1

    Private Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Receivable
        Dim objMulti As New ClsMultiAccount

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

    Private Sub IsiCBOCurrencyTax()
        '- isicbo customer
        Dim objMulti As New ClsMultiAccount
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

    Private Sub IsiCBOAkunDP()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Liabilities
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Other_Payable
        Dim tblAccountDP1 As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
        Dim rowAccountDP1 As DseMultiAccount.tblisi_cbo_ar_akunRow
        Dim objMulti As New ClsMultiAccount

        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboDPAccount.DataSource = Nothing
            Me.cboDPAccount.Items.Clear()
            Me.cboPrimaryDPAccount.DataSource = Nothing
            Me.cboPrimaryDPAccount.Items.Clear()
            Me.tblAccountDP.Rows.Clear()

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccountDP, 1)

            'tmbahkan akun income
            tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Income
            tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Income
            objMulti.IsiTBLAPARAcount(tp1, tp2, tblAccountDP1, 1, False)
            For Each rowAccountDP1 In tblAccountDP1.Rows
                rowAccountDP = Me.tblAccountDP.Newtblisi_cbo_ar_akunRow
                For hj As Integer = 0 To Me.tblAccountDP.Columns.Count - 1
                    rowAccountDP.Item(hj) = rowAccountDP1.Item(hj)
                Next
                Me.tblAccountDP.Rows.Add(rowAccountDP)
            Next

            Me.cboDPAccount.DataSource = Me.tblAccountDP.DefaultView
            Me.cboDPAccount.DisplayMember = "Account Code"
            Me.cboPrimaryDPAccount.DataSource = Me.tblAccountDP.DefaultView
            Me.cboPrimaryDPAccount.DisplayMember = "Primary"
            Me.txtPrimaryDPAccount.Text = Me.cboPrimaryDPAccount.Text
            'Me.txtPrimaryDPAccount.Visible = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged
        Try
            If Me.TxtRemarks.Enabled = True AndAlso _
            IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso _
            (Me.txtPrimaryCurrency.Text) > 0 Then
                Me.IsiCBOCurrencyTax()
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOAkunDP()
                'ini ditambahkan untuk mencegah akun dp yang dipilih ke item pertama
                'tidak masalah ditaro sini karena, apapun primary currencynya, akun dpnya tetap. kecuali one day berubah
                Me.txtPrimaryDPAccount.Text = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)
                SetTaxRate(True)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub cboCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            'Me.txtPrimaryAccount.Visible = True
            If Me.TxtRemarks.Enabled = True AndAlso IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.txtDutyStamp.Text) Then
                    Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                Me.Hitung_Total()
                SetTaxRate(True)
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
                Me.Hitung_Total()
                SetTaxRate(True)
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
            Me.lblCurrTax.Text = Me.cboCurrTax.Text

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

    Private Sub cboDPAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDPAccount.SelectedIndexChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryDPAccount.Text) AndAlso Me.cboPrimaryDPAccount.Text > -1 Then
                Me.txtPrimarydpaccount.Text = Me.cboPrimaryDPAccount.Text
            End If
            'MsgBox("")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryDPAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryDPAccount.SelectedIndexChanged
        Try
            ' Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryDPAccount.Text) AndAlso Me.cboPrimaryDPAccount.Text > -1 Then
                Me.txtPrimarydpAccount.Text = Me.cboPrimaryDPAccount.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryDPAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryDPAccount.TextChanged
        Dim PrimaryDPAkun As Double = -1

        Try
            If IsNumeric(Me.txtPrimaryDPAccount.Text) Then
                PrimaryDPAkun = CDbl(Me.txtPrimaryDPAccount.Text)
            End If

            If Me.tblAccountDP.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblAccountDP.Rows.Count - 1
                    Me.rowAccountDP = Me.tblAccountDP.Rows(IntA)
                    With Me.rowAccountDP
                        If .Primary = PrimaryDPAkun Then
                            Me.cboPrimaryDPAccount.SelectedIndex = IntA
                            Me.cboDPAccount.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim UseTaxSetting, TaxPercentSetting As String
        Dim PrimaryAccDP As String = ""

        Try
            PrimaryAccDP = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)
            If PrimaryAccDP = "" Or PrimaryAccDP = "-1" Then
                MsgBox("Please specify sales down payment account first.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            'Me.dtpDate.Focus()
            Me.objInvoiceDP.ResetObject()
            Me.txtInvoiceNo.Text = Me.objInvoiceDP.GetNoBaru1(Me.dtpDate.Value)
            '2025 Core Tax Tahap 2.2
            ' Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
            Me.objInvoiceDP.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP
            Me.CboFakturPajak.SelectedIndex = 1

            '2025 Core Tax Tahap 2.2
            Me.cboTaxtype.SelectedIndex = Me.objInvoiceDP.getLastTaxType()
            Me.txtNoFakturPajak.Text = ""

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
            Else
                If UseTaxSetting = "0" Then
                    Me.chkTax.Checked = False
                Else
                    Me.chkTax.Checked = True
                    TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(TaxPercentSetting) = True Then
                        Me.txtTax.Text = TaxPercentSetting
                    End If
                End If
            End If
            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            Me.SetDefaultPrecision()
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiCBOCurrencyTax()
            Me.IsiCBOAkunDP()
            Me.txtPrimaryDPAccount.Text = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)

            Me.TxtNamaSigned.Text = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.SignedByName)
            Me.TxtJabatanSigned.Text = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.SignedByJabatan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)

            Me.txtPrimaryAccount.Text = "-1"
            Me.txtPrimaryDPAccount.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.TxtPrimaryInvoice.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryCustomer.Text = "-1"
            Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
            Me.TxtPaymentTerm.Text = "0"
            Me.txtTaxExchangeRate.Text = "1.00"
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur)
            Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur) 'new tax amount1
            Me.TxtTotal.Text = "0"
            Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur)
            Me.txtSubTotal.Text = "0"
            Me.txtExchangeRateToUSD.Text = "1"
            Me.txtExhangeRateToIDR.Text = "1"

            Me.Txt_E_Faktur_Jalan.Text = "-"
            Me.Txt_E_Faktur_Blok.Text = "-"
            Me.Txt_E_Faktur_No.Text = "-"
            Me.Txt_E_Faktur_RT.Text = "0"
            Me.Txt_E_Faktur_RW.Text = "0"
            Me.Txt_E_Faktur_Kecamatan.Text = "-"
            Me.Txt_E_Faktur_Kelurahan.Text = "-"
            Me.Txt_E_Faktur_Propinsi.Text = "-"
            Me.Txt_E_Faktur_Kabupaten.Text = "-"
            Me.Txt_E_Faktur_KodePos.Text = "00000"
            Me.Txt_E_Faktur_NoTelepon.Text = "-"
            Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.chkTax.Checked = True
            Me.cboCustomerAddress.DataSource = Nothing
            Me.cboCustomerAddress.Text = ""
            Me.cboCustomerAddress.Items.Clear()

            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryDPAccount.DataSource = Nothing
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.cboDPAccount.DataSource = Nothing
            Me.cboCurrTax.DataSource = Nothing

            Me.cboDPAccount.Text = ""
            Me.cboAccount.Text = ""
            Me.cboCurrTax.Text = ""

            Me.cboPrimaryAccount.Items.Clear()
            Me.cboPrimaryDPAccount.Items.Clear()
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.cboAccount.Items.Clear()
            Me.cboDPAccount.Items.Clear()
            Me.cboCurrTax.Items.Clear()

            '2025 Core Tax Tahap 2.2 
            Me.AddInfo = -1
            Me.FacilityStamp = -1
            Me.SPPBDate = #1/1/1950#
            'Me.TxtSPPBNo.Text = ""
            Me.txtDPPNilaiLain.Text = "0"
            Me.txtRevisedInvoice.Text = "00" '//#2732

            Me.Txt_E_Faktur_Jalan.Text = ""
            Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryInvoice.BackColor, False)

            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TLayoutPanel2.Controls, True, True, True, True, True, True, True, True, True, True) '//#2732
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.llbAddress.Enabled = True
            Me.llbEditNoFakturPajak.Enabled = True
            Me.llbEditNoInvoice.Enabled = True
            Me.llbUpDateNPWP.Enabled = True

            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                Me.txtInvoiceNo.ReadOnly = True
                Me.txtInvoiceNo.BackColor = Me.TxtReadOnly.BackColor
                Me.txtInvoiceNo.TabStop = False
                Me.dtpDate.Focus()
                Me.llbEditNoInvoice.Visible = True
            Else
                Me.txtInvoiceNo.ReadOnly = False
                Me.txtInvoiceNo.BackColor = Me.TxtPrimaryInvoice.BackColor
                Me.txtInvoiceNo.TabStop = True
                Me.txtInvoiceNo.Focus()
                Me.llbEditNoInvoice.Visible = False
                Me.txtInvoiceNo.Width = Me.dtpDate.Width
            End If

            'set default precision for ide
            Me.DefaultPrecision_DefCur = Me.ObjCurrency.GetDecimal_DefaultCurrency

            'set llb faktur pajak
            If Me.chkTax.Checked = True Then
                Me.llbEditNoFakturPajak.Enabled = True
            Else
                Me.llbEditNoFakturPajak.Enabled = False
            End If

            Me.cboAccount.Enabled = True
            Me.cboDPAccount.Enabled = True
            Me.cboCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
                Me.IsiCBOAkunDP()
                Me.txtPrimaryDPAccount.Text = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)
            End If

            txtNoFakturPajakSederhana.ReadOnly = True
            txtNoFakturPajakSederhana.BackColor = Me.TxtReadOnly.BackColor


            '2025 Core Tax Tahap 2.2
            Me.llbEditNoFakturPajak.Visible = False
            Me.txtNoSeriFakturPajak.Visible = False
            Me.txtNoFakturPajak.Visible = True
            Me.LlbViewDetailTax.Enabled = False
            If Me.baru Then
                Me.chkRevisedInvoice.Enabled = False
            Else
                Me.chkRevisedInvoice.Enabled = True
            End If
            'Me.lblSPPBNo.Visible = False
            'Me.TxtSPPBNo.Visible = False
            Me.txtDPPNilaiLain.ReadOnly = True

            Me.SetTaxRate(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TLayoutPanel2.Controls, True, False, True, False, True, False, True, False, True, False) '//#2732
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit

            Me.llbAddress.Enabled = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.BackColor, False)

            Me.llbEditNoFakturPajak.Enabled = False
            Me.llbEditNoInvoice.Enabled = False
            Me.llbUpDateNPWP.Enabled = False
            Me.cboAccount.Enabled = False
            Me.cboDPAccount.Enabled = False
            Me.cboCurrTax.Enabled = False

            '2025 Core Tax Tahap 2.2
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
            Me.LlbViewDetailTax.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date
        Dim DefTaxRate As Decimal = 0
        Dim SO_Amount As Decimal = 0
        Dim Tgl_SO As Date
        Dim Dummy As String = ""

        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("invoice number")
            KoleksiTextBox.Add(Me.txtInvoiceNo)
            KoleksiNamaData.Add("customer code")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            KoleksiNamaData.Add("currency")
            KoleksiTextBox.Add(Me.txtCurrencyCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            'cek bilangan positif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            KoleksiNamaData.Add("down payment")
            KoleksiTextBox.Add(Me.txtSubTotal)
            If ModCekData.CekBilanganPositif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber
            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.objInvoiceDP.TimeStamp <> Me.objInvoiceDP.GetTimestamp(Me.objInvoiceDP.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objInvoiceDP.CariDetail(Me.objInvoiceDP.Primary)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If


            ''cek berkaitan dengan tanggal closing
            'TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            'If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
            '    MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.objInvoiceDP.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objInvoiceDP.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    'Me.txtInvoiceNo.Text = Me.objInvoiceDP.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objInvoiceDP.TanggalTransaksi
                        Me.txtInvoiceNo.Text = Me.objInvoiceDP.NoTransaksi
                        Return False
                    End If
                End If
            End If

            'data salah jika, tidak ada kurs pajak untuk tgl ybs
            If Me.txtPrimaryCurrency.Text <> "1" Then
                If Me.dtpDate.Value.Date >= Me.ObjTaxInvoiceSetting.BatasStartDate Then
                    If Me.ObjTaxRate.ApakahSudahAda(Me.dtpDate.Value.Date) = False Then
                        MsgBox("Please specify tax invoice rate data current transaction date.", MsgBoxStyle.Information, "Information")
                        Me.ShowTaxSetting()
                        Return False
                    End If
                End If

                'jika kursnya tidak sama, maka berikan peringatan
                DefTaxRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.txtPrimaryCurrency.Text)
                If CDec(Me.txtTaxExchangeRate.Text) <> DefTaxRate Then
                    If MsgBox("The tax exchange rate differs from default tax exchange rate for " + Me.txtCurrencyCode.Text + " on " + _
                    Format(Me.dtpDate.Value, "dd MMM yyyy") + " which is " + Me.ObjCD.DigitGrouping(DefTaxRate, Me.ObjTaxRate.DefaultPrec) + " continue save?", MsgBoxStyle.YesNo, _
                    "Confirm") <> MsgBoxResult.Yes Then
                        Me.txtTaxExchangeRate.Focus()
                        Return False
                    End If
                End If
            End If

            'jika tidak ada no faktur pajak standar
            If Me.CboFakturPajak.SelectedIndex = 1 And Me.chkTax.Checked = True Then
                If Me.txtNoFakturPajak.Text = "" Then
                    '2025 Core Tax Tahap 2.2 
                    'Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    'If Me.txtNoFakturPajak.Text = "" Then
                    '    MsgBox("There are no new tax invoice number available. Please specify  new tax invoice number in your global setting.", _
                    '    MsgBoxStyle.Information, "Information")
                    '    Return False
                    'End If
                    'MsgBox("Please enter tax invoice number", MsgBoxStyle.Information, "Information")
                    'Return False
                End If
            End If


            If Me.chkTax.Checked = True Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Me.txtTax.Focus()
                    Return False
                End If
                '2025 Core Tax Tahap 2.1, di comment update
                'If CDec(Me.txtTax.Text) <> 10 Then
                '    If MsgBox("Tax rate <> 10%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                '        Me.txtTax.Focus()
                '        Return False
                '    End If
                'End If
            End If

            'berkaitan dengan tax exchange rate
            If Me.txtCurrencyCode.Text = "IDR" Then
                If CDbl(Me.txtTaxExchangeRate.Text) <> 1 Then
                    MsgBox("For invoice with IDR currency, the tax exchange rate value must be 1.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", 2)
                    Return False
                End If
            Else
                'jika bukan IDR dan tax ratenya bernilai 1 maka berikan peringatan
                If CDbl(Me.txtTaxExchangeRate.Text) = 1 Then
                    If MsgBox("This invoice use non-IDR currency and the tax exchange rate value = 1. Usually this is wrong data input, continue?", _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If

            'nilai dp maksimum adalah sama dengan nilai po - nilai dp lain
            SO_Amount = Me.ObjSO.GetSubTotal(Me.txtPrimarySO.Text)
            SO_Amount -= Me.objInvoiceDP.GetDPAmount_ExcludePrimaryInv(Me.txtPrimarySO.Text, Me.TxtPrimaryInvoice.Text)
            SO_Amount -= Me.ObjBBDP.GetJml(ClsTransaksiBalanceDP.EnumTipeDP.DP_Invoice, Me.txtPrimarySO.Text)
            If SO_Amount < CDec(Me.txtSubTotal.Text) Then
                MsgBox("Maximum down payment for " + Me.txtSONo.Text + " is " + Me.txtCurrencyCode.Text + " " + _
                Me.ObjCD.DigitGrouping(SO_Amount, ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)), MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek tanggalnya, minimum sama dengan tanggal PO
            Tgl_SO = Me.ObjSO.GetTanggal(Me.txtPrimarySO.Text)
            If Me.dtpDate.Value < Tgl_SO Then
                MsgBox("Minimum down payment date is: " + Format(Tgl_SO, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")

                Return False
            End If

            '2025 Core Tax Tahap 2.2 
            'kalo pake pajak, maka no seri faktur pajak harus diisikan
            'Me.txtNoSeriFakturPajak.Text = Me.txtNoSeriFakturPajak.Text.Trim
            'If Me.chkTax.Checked = True And Me.txtNoSeriFakturPajak.Text = "" Then
            '    MsgBox("Please specify tax invoice serial number.", MsgBoxStyle.Information, "Information")
            '    Me.txtNoSeriFakturPajak.Focus()
            '    Return False
            'End If



            If Me.chkTax.Checked = True Then
                'kalo ada VAT, maka akun VAT out harus ada
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify VAT out account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If CDbl(Me.txtDutyStamp.Text) > 0 Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify dutystamp expense account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '2025 Core Tax Tahap 2.2 
            'berkaitan dengan data e-faktur yang tidak lengkap
            'If Me.chkTax.Checked = True Then
            '    Me.Txt_E_Faktur_Blok.Text = Me.Txt_E_Faktur_Blok.Text.Trim
            '    Me.Txt_E_Faktur_Jalan.Text = Me.Txt_E_Faktur_Jalan.Text.Trim
            '    Me.Txt_E_Faktur_Kelurahan.Text = Me.Txt_E_Faktur_Kelurahan.Text.Trim
            '    Me.Txt_E_Faktur_KodePos.Text = Me.Txt_E_Faktur_KodePos.Text.Trim
            '    Me.Txt_E_Faktur_No.Text = Me.Txt_E_Faktur_No.Text.Trim
            '    Me.Txt_E_Faktur_NoTelepon.Text = Me.Txt_E_Faktur_NoTelepon.Text.Trim
            '    Me.Txt_E_Faktur_Propinsi.Text = Me.Txt_E_Faktur_Propinsi.Text.Trim
            '    Me.Txt_E_Faktur_RT.Text = Me.Txt_E_Faktur_RT.Text.Trim
            '    Me.Txt_E_Faktur_RW.Text = Me.Txt_E_Faktur_RW.Text.Trim
            '    Me.Txt_E_Faktur_Kabupaten.Text = Me.Txt_E_Faktur_Kabupaten.Text.Trim
            '    Me.Txt_E_Faktur_Kecamatan.Text = Me.Txt_E_Faktur_Kecamatan.Text.Trim

            '    If Me.Txt_E_Faktur_Blok.Text = "" Or _
            '    Me.Txt_E_Faktur_Jalan.Text = "" Or _
            '    Me.Txt_E_Faktur_Kelurahan.Text = "" Or _
            '    Me.Txt_E_Faktur_KodePos.Text = "" Or _
            '    Me.Txt_E_Faktur_No.Text = "" Or _
            '    Me.Txt_E_Faktur_NoTelepon.Text = "" Or _
            '    Me.Txt_E_Faktur_Propinsi.Text = "" Or _
            '    Me.Txt_E_Faktur_RT.Text = "" Or _
            '    Me.Txt_E_Faktur_RW.Text = "" Or _
            '    Me.Txt_E_Faktur_Kecamatan.Text = "" Or _
            '    Me.Txt_E_Faktur_Kabupaten.Text = "" Then
            '        MsgBox("Incomplete e-Faktur data, please fix.", MsgBoxStyle.Information, "Information")
            '        Me.ShowDetailFakturPajak()
            '        If Me.Txt_E_Faktur_Blok.Text = "" Or _
            '           Me.Txt_E_Faktur_Jalan.Text = "" Or _
            '           Me.Txt_E_Faktur_Kelurahan.Text = "" Or _
            '           Me.Txt_E_Faktur_KodePos.Text = "" Or _
            '           Me.Txt_E_Faktur_No.Text = "" Or _
            '           Me.Txt_E_Faktur_NoTelepon.Text = "" Or _
            '           Me.Txt_E_Faktur_Propinsi.Text = "" Or _
            '           Me.Txt_E_Faktur_RT.Text = "" Or _
            '           Me.Txt_E_Faktur_RW.Text = "" Or _
            '           Me.Txt_E_Faktur_Kecamatan.Text = "" Or _
            '           Me.Txt_E_Faktur_Kabupaten.Text = "" Then
            '            Return False
            '        End If
            '    End If
            'End If

            'no seri faktur pajak harus: 010, 020, 030, 040, 060, 070, 080, 090
            'If Me.chkTax.Checked = True Then
            '    If Me.txtNoSeriFakturPajak.Text <> "010" And _
            '                Me.txtNoSeriFakturPajak.Text <> "020" And _
            '                Me.txtNoSeriFakturPajak.Text <> "030" And _
            '                Me.txtNoSeriFakturPajak.Text <> "040" And _
            '                Me.txtNoSeriFakturPajak.Text <> "060" And _
            '                Me.txtNoSeriFakturPajak.Text <> "070" And _
            '                Me.txtNoSeriFakturPajak.Text <> "080" And _
            '                Me.txtNoSeriFakturPajak.Text <> "090" And _
            '                 Me.txtNoSeriFakturPajak.Text <> "011" And _
            '                Me.txtNoSeriFakturPajak.Text <> "021" And _
            '                Me.txtNoSeriFakturPajak.Text <> "031" And _
            '                Me.txtNoSeriFakturPajak.Text <> "041" And _
            '                Me.txtNoSeriFakturPajak.Text <> "061" And _
            '                Me.txtNoSeriFakturPajak.Text <> "071" And _
            '                Me.txtNoSeriFakturPajak.Text <> "081" And _
            '                Me.txtNoSeriFakturPajak.Text <> "091" Then
            '        txtNoSeriFakturPajak.Text = txtNoSeriFakturPajak.Text.Trim
            '        MsgBox("The serial number of tax invoice must be chosen from one of the following alternatives: 010, 011, 020, 021, 030, 031, 040, 041, 060, 061, 070, 071, 080, 081, 091 or 090.", _
            '        MsgBoxStyle.Information, "Information")
            '        Me.txtNoSeriFakturPajak.Focus()
            '        Return False
            '    End If

            '    If Me.txtNoSeriFakturPajak.Text = "070" Or _
            '                Me.txtNoSeriFakturPajak.Text = "080" Or _
            '                Me.txtNoSeriFakturPajak.Text = "071" Or _
            '                Me.txtNoSeriFakturPajak.Text = "081" Then
            '        Dim ijk As New FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur
            '        ijk.TxtTaxInvoiceSeries.Text = Me.txtNoSeriFakturPajak.Text
            '        ijk.vIndexAcuan = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
            '        ijk.ShowDialog()
            '        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = ijk.cboKeteranganTambahan.SelectedIndex
            '    End If
            'End If


            'Account receivable dan currency tax tidak boleh berbeda dengan data Invoice
            If IsNumeric(Me.txtPrimarySO.Text) AndAlso _
               IsNumeric(Me.txtPrimaryAccount.Text) AndAlso _
               IsNumeric(Me.txtPrimaryCurrTax.Text) Then
                If Me.objInvoiceDP.GetJmlAOPSama(Me.txtPrimarySO.Text, _
                                               Me.txtPrimaryAccount.Text, _
                                               Me.txtPrimaryCurrTax.Text, ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice) <= 0 Then
                    Dim PrimaryInvoiceDPFirst As Double = Me.objInvoiceDP.GetPrimaryAOPwithPOSama(Me.txtPrimarySO.Text, ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice)
                    If PrimaryInvoiceDPFirst > -1 Then
                        Dim InvoiceNO As String = Me.objInvoiceDP.GetNoInvoice(PrimaryInvoiceDPFirst)
                        Dim PrimaryAccountB4 As Double = Me.objInvoiceDP.GetPrimaryAccountAR(PrimaryInvoiceDPFirst)
                        Dim PrimaryCurrTaxB4 As Double = Me.objInvoiceDP.GetPrimaryCurrTax(PrimaryInvoiceDPFirst)
                        Dim PrimaryAccountDPB4 As Double = Me.objInvoiceDP.GetPrimaryAccountDP(PrimaryInvoiceDPFirst)
                        If PrimaryAccountDPB4 <> Me.txtPrimaryDPAccount.Text OrElse _
                           PrimaryCurrTaxB4 <> Me.txtPrimaryCurrTax.Text Then
                            MsgBox("There's some of invoice is using the similiar SO Number, with different account payable or different currency tax. " & _
                                   "please use the same account payable and currency tax for this transaction. Invoice No. " & _
                                   InvoiceNO & ".", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If


                If Me.objInvoiceDP.GetJmlAOPwithPOSama(Me.txtPrimarySO.Text, ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP) > 0 Then
                    Dim PrimaryInvoiceDPFirst As Double = Me.objInvoiceDP.GetPrimaryAOPwithPOSama(Me.txtPrimarySO.Text, ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP)
                    Dim InvoiceNO As String = Me.objInvoiceDP.GetNoInvoice(PrimaryInvoiceDPFirst)
                    Dim PrimaryAccountB4 As Double = Me.objInvoiceDP.GetPrimaryAccountAR(PrimaryInvoiceDPFirst)
                    Dim PrimaryCurrTaxB4 As Double = Me.objInvoiceDP.GetPrimaryCurrTax(PrimaryInvoiceDPFirst)
                    Dim PrimaryAccountDPB4 As Double = Me.objInvoiceDP.GetPrimaryAccountdp(PrimaryInvoiceDPFirst)

                    If PrimaryAccountDPB4 <> Me.txtPrimaryDPAccount.Text OrElse _
                       PrimaryAccountB4 <> Me.txtPrimaryAccount.Text OrElse _
                       PrimaryCurrTaxB4 <> Me.txtPrimaryCurrTax.Text Then
                        MsgBox("There's some of invoice down payment using the similiar SO Number, " & _
                        "with diferent account DP or account payable or diferent currency tax, " & _
                        "please use the same account DP and account payable and currency tax for this transaction. Invoice No. " & _
                        InvoiceNO & ".", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            If CDec(Me.txtDutyStamp.Text) > 0 Then
                If Me.txtPrimaryCurrency.Text <> "1" Then
                    MsgBox("Duty stamp is for IDR invoices only.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '2. SO yang berasal dari quick invoice tidak boleh diedit.
            If Me.txtPrimarySO.Text = "" Then
                Me.txtPrimarySO.Text = -1
            End If
            Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryQuickInvoice(Me.txtPrimarySO.Text).HasValue = True Then
                Dim PrimaryQuickInvoice As Double = DTA.ScalarQuery_GetPrimaryQuickInvoice(Me.txtPrimarySO.Text).Value
                If PrimaryQuickInvoice <> -1 Then
                    Dim ObjInv As New ClsTransaksiSalesInvoice
                    MsgBox("This SO is refered to quick invoice: " + _
                    ObjInv.GetNoInvoice(PrimaryQuickInvoice) + ". You cannot use this data.", MsgBoxStyle.Information, "Information")
                    ObjInv = Nothing
                    Return False
                End If
            End If

            'Comment: Perubahan NPWP 16 Digit
            If chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                '2025 Core Tax Tahap 2.1
                If ObjSetNPWP.SetNPWP_New(Me.dtpDate.Value.Date, Me.txtPrimaryCustomer.Text, Me.txtNoNPWP.Text, _
                ClsTrans_SettingNPWPNo.EnumFromMaster.Customer, Me.txtNoNIKPass.Text, frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionSales, _
                Me.txtPrimaryAccount.Text, Me.UserCode) = False Then
                    Return False
                End If
            End If

            '2025 Core Tax Tahap 2.2, show info pajak, ketika pilihan 08, nanti diganti menjadi yang bawah, ketika full
            If Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax07 Or Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax08 Then
                If ShowInfoPajak() = False Then
                    Return False
                End If
            End If

            '2025 Coretax Tahap 2.2, cek apakah boleh 01 atau tidak
            If chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                If Me.dtpDate.Value <= ObjSetNPWP.TglAkhirTax04 And Me.cboTaxtype.SelectedIndex = 0 Then
                    MsgBox("The tax type 01 can't be used until " & ObjSetNPWP.TglAkhirTax04.AddDays(1) & ".", MsgBoxStyle.Information, _
                            "Information")
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

    Private Sub SalinKeField()
        '2025 Core Tax Tahap 2.1
        Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
        Try
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.objInvoiceDP.PrimaryCurrency)
            Me.DefaultPrecision_DefCur = Me.ObjCurrency.GetDecimal_DefaultCurrency

            With Me.objInvoiceDP
                Me.txtNoFakturPajak.Text = .NoFakturPajak
                Me.txtNoNPWP.Text = .NoNPWP
                '2025 Core Tax Tahap 2.1
                If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                    'Me.llbUpDateNPWP.Visible = False
                    Me.cboNIKPASS.Visible = True
                    Me.cboNIKPASS.TabStop = True
                    Me.txtNoNIKPass.Visible = True
                    Me.txtNoNPWP.Visible = False
                    '2025 Core Tax Tahap 2.1
                    Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                    Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(.PrimaryCustomer)
                    If Me.baru = True Then
                        Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(.PrimaryCustomer)
                    Else
                        Me.txtNoNIKPass.Text = Me.objInvoiceDP.GetNoNikPassbyInvoice(.Primary)
                    End If
                Else
                    'Me.llbUpDateNPWP.Visible = True
                    Me.cboNIKPASS.Visible = False
                    Me.cboNIKPASS.TabStop = False
                    Me.txtNoNIKPass.Visible = False
                    Me.txtNoNPWP.Visible = True
                End If
                Me.txtNoSeriFakturPajak.Text = .NoSeriFakturPajak

                Me.CboFakturPajak.SelectedIndex = .JenisFakturPajak
                Me.TxtPrimaryInvoice.Text = .Primary
                Me.txtInvoiceNo.Text = .NoTransaksi
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.IsiCBOArAkunByTipePerkiraan1()
                If .PrimaryAccount > -1 Then
                    Me.txtPrimaryAccount.Text = .PrimaryAccount
                    Me.cboPrimaryAccount.Text = .PrimaryAccount
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                End If

                Me.IsiCBOAkunDP()
                If .PrimaryAccount > -1 Then
                    Me.txtPrimaryDPAccount.Text = .PrimaryAccountDP
                    Me.cboPrimaryDPAccount.Text = .PrimaryAccountDP
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryDPAccount.DataSource = Nothing
                    Me.txtPrimaryDPAccount.Text = -1
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


                Me.SetDefaultPrecision()
                Me.txtPrimarySO.Text = .PrimarySO4DP
                Me.txtSONo.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                Me.SetSalesPerson()
                Me.txtPrimaryCustomer.Text = .PrimaryCustomer
                Me.txtCustomerAddIdentifier.Text = .CustIdentifierAdd
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotalB4DP, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax)))
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax))) 'new tax amount1
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxExchangeRate, Me.DefaultPrecision))
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.DutyStamp, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax)))
                Me.txtExhangeRateToIDR.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExchangRateToIDR, Me.DefaultPrecision))
                Me.txtExchangeRateToUSD.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExhangeRateToUSD, Me.DefaultPrecision))
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(.TaxPercentage, 2)
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))
                Me.chkTax.Checked = .UseTax
                Me.TxtRemarks.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtNoFakturPajakSederhana.Text = .NoFakturPajakSederhana
                Me.TxtPaymentTerm.Text = .PaymentTerm
                Me.TxtNamaSigned.Text = .NamaSigned
                Me.TxtJabatanSigned.Text = .JabatanSigned
                If .VATDipungut = True Then
                    Me.cboPPNDipungut.SelectedIndex = 0
                Else
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If

                Me.Txt_E_Faktur_Blok.Text = .EFak_Blok
                Me.Txt_E_Faktur_Jalan.Text = .EFak_NamaJalan
                Me.Txt_E_Faktur_Kelurahan.Text = .EFak_Kelurahan
                Me.Txt_E_Faktur_KodePos.Text = .EFak_KodePos
                Me.Txt_E_Faktur_No.Text = .EFak_No
                Me.Txt_E_Faktur_NoTelepon.Text = .EFak_Telepon
                Me.Txt_E_Faktur_Propinsi.Text = .EFak_Propinsi
                Me.Txt_E_Faktur_RT.Text = .EFak_RT
                Me.Txt_E_Faktur_RW.Text = .EFak_RW
                Me.Txt_E_Faktur_Kabupaten.Text = .EFak_Kabupaten
                Me.Txt_E_Faktur_Kecamatan.Text = .EFak_Kecamatan
                Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = .EFak_KodeJenisTransaksiPajak

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.LblPrintStatus.Text = "Print Status: " + Me.objInvoiceDP.GetPrintStatus(Me.TxtPrimaryInvoice.Text)
                If Me.LblPrintStatus.Text = "Print Status: " Then
                    Me.LblPrintStatus.Text = "Print Status: -"
                End If

                '2025 Core Tax Tahap 2.2
                Me.SPPBDate = .SPPBDate
                'Me.chkRevisedInvoice.Checked = .TaxRevised
                '//#2732
                If .TaxRevised = True Then
                    Me.txtRevisedInvoice.Text = "01"
                Else
                    Me.txtRevisedInvoice.Text = "00"
                End If

                Me.AddInfo = .PrimaryAddInfo
                Me.FacilityStamp = .PrimaryFacilityStamp
                Me.cboTaxtype.SelectedIndex = .TaxType
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(.DPPNilaiLain, Me.DefaultPrecision)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objInvoiceDP
                'invoice DP: DP=0 dan Calculates by invoice
                .PrimarySO4DP = Me.txtPrimarySO.Text
                .CalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice
                .DownPayment = 0
                .DutyStamp = Me.txtDutyStamp.Text
                .ExchangRateToIDR = Me.txtExhangeRateToIDR.Text
                .ExhangeRateToUSD = Me.txtExchangeRateToUSD.Text
                .NoFakturPajak = Me.txtNoFakturPajak.Text
                .NoNPWP = Me.txtNoNPWP.Text
                .NoSeriFakturPajak = Me.txtNoSeriFakturPajak.Text
                'invoice DP tidak ada revisi
                .Revisi = 0
                .SubTotalB4DP = Me.txtSubTotal.Text
                .TaxExchangeRate = Me.txtTaxExchangeRate.Text
                .CustIdentifierAdd = Me.txtCustomerAddIdentifier.Text
                .NoTransaksi = Me.txtInvoiceNo.Text
                .PaymentTerm = ""
                .PrimaryCurrency = Me.txtPrimaryCurrency.Text
                .PrimaryCustomer = Me.txtPrimaryCustomer.Text
                .SubTotal = Me.txtSubTotal.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .TaxAmount1 = Me.txtTaxAmount1.Text 'new tax amount1
                .TaxPercentage = Me.txtTax.Text
                .Total = Me.TxtTotal.Text
                .UseTax = Me.chkTax.Checked
                .Remarks = Me.TxtRemarks.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .JenisFakturPajak = Me.CboFakturPajak.SelectedIndex
                .NoFakturPajakSederhana = Me.txtNoFakturPajakSederhana.Text
                .PaymentTerm = Me.TxtPaymentTerm.Text
                .NamaSigned = Me.TxtNamaSigned.Text
                .JabatanSigned = Me.TxtJabatanSigned.Text
                .UrutanFakturPajak = 0
                If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    .VATDipungut = True
                Else
                    .VATDipungut = False
                End If

                .EFak_Blok = Me.Txt_E_Faktur_Blok.Text
                .EFak_NamaJalan = Me.Txt_E_Faktur_Jalan.Text
                .EFak_Kelurahan = Me.Txt_E_Faktur_Kelurahan.Text
                .EFak_KodePos = Me.Txt_E_Faktur_KodePos.Text
                .EFak_No = Me.Txt_E_Faktur_No.Text
                .EFak_Telepon = Me.Txt_E_Faktur_NoTelepon.Text
                .EFak_Propinsi = Me.Txt_E_Faktur_Propinsi.Text
                .EFak_RT = Me.Txt_E_Faktur_RT.Text
                .EFak_RW = Me.Txt_E_Faktur_RW.Text
                .EFak_Kabupaten = Me.Txt_E_Faktur_Kabupaten.Text
                .EFak_Kecamatan = Me.Txt_E_Faktur_Kecamatan.Text
                .EFak_KodeJenisTransaksiPajak = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text

                'jika datanya baru, maka update data alamat pajak 1 dan alamat pajak 2
                If Me.baru = True Then
                    Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, .AlamatPajak1, .AlamatPajak2)
                Else
                    'jika ada perubahan data customer dari invoice lama maka update data alamat customer
                    If Me.objInvoiceDP.Get_Primary_Cust(Me.TxtPrimaryInvoice.Text) <> Me.txtPrimaryCustomer.Text Then
                        Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, .AlamatPajak1, .AlamatPajak2)
                    End If
                End If
                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
                .PrimaryAccountDP = Me.txtPrimaryDPAccount.Text

                '2025 Core Tax Tahap 2.1
                .OptionNIKPASS = Me.cboNIKPASS.SelectedIndex
                .NoNikPass = Me.txtNoNIKPass.Text

                '2025 Core Tax Tahap 2.2
                .SPPBDate = Me.SPPBDate
                '.TaxRevised = Me.chkRevisedInvoice.Checked
                '//#2732
                If Me.txtRevisedInvoice.Text = "01" Then
                    .TaxRevised = True
                Else
                    .TaxRevised = False
                End If
                If Me.chkTax.Checked = False Then
                    .NoFakturPajak = ""
                End If

                .PrimaryAddInfo = Me.AddInfo
                .PrimaryFacilityStamp = Me.FacilityStamp
                .TaxType = Me.cboTaxtype.SelectedIndex
                .NoSeriFakturPajak = Me.cboTaxtype.SelectedItem.ToString
                .DPPNilaiLain = Me.txtDPPNilaiLain.Text
            End With
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
                    Me.SetDefaultPrecision()
                    Me.txtSubTotal.Focus()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision()
                        Me.txtSubTotal.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtDutyStamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                ' Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, Me.ObjCurrency.GetDecimal_DefaultCurrency)
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================
    '====================================================================================================================================================================

    Private Sub dtpDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.GotFocus
        Me.dtpDate.BackColor = BackColorAll
    End Sub

    Private Sub txtPONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSONo.GotFocus
        Me.txtSONo.BackColor = BackColorAll
    End Sub

    Private Sub txtPONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSONo.LostFocus
        Me.txtSONo.BackColor = Me.TxtPrimaryInvoice.BackColor
    End Sub

    Private Sub txtSubTotal_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubTotal.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtSubTotal, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubTotal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubTotal.GotFocus
        Me.txtSubTotal.BackColor = BackColorAll
    End Sub

    Private Sub txtSubTotal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSubTotal.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSubTotal, sender, e)
    End Sub

    Private Sub txtSubTotal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSubTotal.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSubTotal, sender, e)
    End Sub

    Private Sub txtSubTotal_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubTotal.LostFocus
        Try
            Me.txtSubTotal.BackColor = Me.TxtPrimaryInvoice.BackColor
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.GotFocus
        Me.txtDutyStamp.BackColor = BackColorAll
    End Sub

    Private Sub txtDutyStamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.LostFocus
        If Me.txtDutyStamp.Enabled = True And Me.txtDutyStamp.ReadOnly = False Then
            Me.txtDutyStamp.BackColor = Me.TxtPrimaryInvoice.BackColor
            Me.Hitung_Total()
        End If
    End Sub

    Private Sub FrmTransaksiAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.baru = True
            Me.CboFakturPajak.SelectedIndex = 0
            Me.cboPPNDipungut.SelectedIndex = 0
            Me.NonAktifkanDetail()
            'If Me.UserCode <> "USER" Then
            '    Me.LblPrintStatus.Visible = False
            'End If
            '2025 Core Tax Tahap 2.2
            Me.llbEditNoFakturPajak.Visible = False
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
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
        txtExchangeRateToUSD.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub

    Private Sub txtExhangeRateToIDR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.GotFocus
        Me.txtExhangeRateToIDR.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtExhangeRateToIDR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.LostFocus
        txtExhangeRateToIDR.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        TxtRemarks.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        TxtRemarks.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                Dim ij As New FrmCetakInvoice
                If Me.objInvoiceDP.CetakInvoiceDP(ij.crv1, Me.TxtPrimaryInvoice.Text, _
                Me.dtpDate.Value, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text), Me.UserCode) Then
                    ij.ShowDialog()
                    ij.Dispose()
                    If MsgBox("Is invoice printed correctly?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        'update printed invoice
                        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
                        DTA.Connection.ConnectionString = Core_Module.GetConstring
                        DTA.UpdateQuery_SetPrintStatus(Me.UserCode, Now, Me.TxtPrimaryInvoice.Text)
                        DTA.Dispose()

                        Me.LblPrintStatus.Text = "Print Status: " + Me.objInvoiceDP.GetPrintStatus(Me.TxtPrimaryInvoice.Text)
                        If Me.LblPrintStatus.Text = "Print Status: " Then
                            Me.LblPrintStatus.Text = "Print Status: -"
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowTaxSetting()
        Dim ObjHakAkses As New Core_Module.ClsUser

        Try
            If ObjHakAkses.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoiceSettingPajak") = True Then
                Dim ij As Sales_Module.FrmTransaksiSalesInvoiceSettingPajak
                Dim j As Windows.Forms.Form

                For Each j In Me.MdiChildren
                    ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoiceSettingPajak)
                    If IsNothing(ij) = False Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                Next

                Dim i As New Sales_Module.FrmTransaksiSalesInvoiceSettingPajak
                i.UserCode = Me.UserCode
                ' Set the caption to be unique.
                i.BackColorAll = Me.BackColorAll
                i.MdiParent = Me.ParentForm
                i.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjHakAkses = Nothing
        End Try
    End Sub



    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Down Payment Invoice"
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.objInvoiceDP.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP
            Control.CheckForIllegalCrossThreadCalls = False
            Me.ShowIcon = False
            '2025 Core Tax Tahap 2.2
            Me.llbEditNoFakturPajak.Visible = False
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objInvoiceDP.SimpanBaru(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                            Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                            Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                            Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objInvoiceDP.SimpanPerubahan(Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                            Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                            Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                            Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
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
            If Me.TxtRemarks.Text = "" And CDbl(Me.TxtTotal.Text) = 0 And Me.txtSONo.Text = "" _
            And Me.txtCurrencyCode.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.objInvoiceDP.CariDetail(Me.TxtPrimaryInvoice.Text)
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
        Dim hj As New FrmCariTransaksiSalesInvoice
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP

            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.objInvoiceDP.CariDetail(hj.PrimaryTransaksi) = True Then
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
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.objInvoiceDP.ApakahBolehEdit(Me.objInvoiceDP.Primary, Me.UserCode, True) = True Then
                    If Me.objInvoiceDP.ApakahQuickInvoice(Me.objInvoiceDP.Primary) = True Then
                        MsgBox("This invoice is not sales invoice. Use quick invoice form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.baru = False
                    Me.objInvoiceDP.CariDetail(Me.objInvoiceDP.Primary)
                    Me.AktifkanDetail()
                    Me.SalinKeField()
                    Me.dtpDate.Focus()
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
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.objInvoiceDP.ApakahBolehHapus(Me.objInvoiceDP.Primary, True, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    If Me.objInvoiceDP.ApakahQuickInvoice(Me.objInvoiceDP.Primary) = True Then
                        MsgBox("This invoice is not sales invoice. Use quick invoice form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objInvoiceDP.Hapus(Me.TxtPrimaryInvoice.Text, Me.UserCode, True, False) = True Then
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
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objInvoiceDP.SimpanBaru(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                                Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objInvoiceDP.SimpanPerubahan(Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                                Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
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

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            Me.TxtRemarks.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            Me.TxtRemarks.BackColor = Me.TxtPrimaryInvoice.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtInvoiceNo.ReadOnly = False And Me.txtInvoiceNo.Enabled = True Then
                    txtInvoiceNo.BackColor = BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtInvoiceNo.ReadOnly = False And Me.txtInvoiceNo.Enabled = True Then
                    txtInvoiceNo.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = True Then
                    Me.txtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtCustomerCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = True Then
                    Me.txtCustomerCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtCustomerCode.BackColor = Me.txtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        txtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        txtCurrencyCode.BackColor = Me.TxtPrimaryInvoice.BackColor
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub Hitung_Total()
        Dim exchangRate As Double = 0
        Dim DefPrec As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Total_MiddlePoint()
                Exit Sub
            End If
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            'hitung pajak
            Dim DefPrecTax As Decimal = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefaultPrecision
            End If


            If Me.chkTax.Checked = True Then
                'If Me.cboPPNDipungut.SelectedIndex = 0 Then
                'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                If Me.txtPrimaryCurrency.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    CDec(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                Else
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                              CDec(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    Else

                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                               CDec(Me.txtSubTotal.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    End If
                End If

                Me.txtTaxAmount1.Text = Me.txtTaxAmount.Text 'new tax amount1
                If Me.cboPPNDipungut.SelectedIndex = 1 Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                End If
                'Else
                '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                'End If
            Else
                Me.txtTax.Text = "0.00"
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax) 'new tax amount1
            End If
            Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DefPrecTax)


            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.txtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotal.Text) Then
                    decTotal += CDec(Me.txtSubTotal.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.txtDutyStamp.Text) Then
                    decTotal += CDec(Me.txtDutyStamp.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
            Else
                If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotal.Text) Then
                        decTotal += CDec(Me.txtSubTotal.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.txtDutyStamp.Text) Then
                        decTotal += CDec(Me.txtDutyStamp.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotal.Text) Then
                        decTotal += CDec(Me.txtSubTotal.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                End If
            End If

            ''hitung pajak
            'If Me.cboPPNDipungut.SelectedIndex = 0 Then
            '    If Me.txtPrimaryCurrency.Text = 1 Then
            '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * _
            '        CDbl(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrec)
            '    Else
            '        'perhitungan pajak dengan exchange rate
            '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Math.Round(CDec(Me.txtSubTotal.Text) * _
            '        CDec(Me.txtTaxExchangeRate.Text), DefPrec, MidpointRounding.AwayFromZero) * _
            '        CDec(Me.txtTax.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrec)
            '    End If
            'Else
            '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrec)
            'End If

            ''hitung total = sub total + tax
            'If Me.txtPrimaryCurrency.Text = "1" Then
            '    'kalo IDR, duty stamp diperhitungkan
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) + _
            '    CDbl(Me.txtTaxAmount.Text) + CDbl(Me.txtDutyStamp.Text), DefaultPrecision))
            'Else
            '    'kalo bukan IDR duty stamp dan tax tidak diperhitungkan
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) , DefaultPrecision))
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Total_MiddlePoint()
        Dim exchangRate As Double = 0
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency

            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            'hitung pajak
            Dim DefPrecTax As Decimal = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefaultPrecision
            End If


            If Me.chkTax.Checked = True Then
                'If Me.cboPPNDipungut.SelectedIndex = 0 Then
                'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                If Me.txtPrimaryCurrency.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    CDec(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                Else
                    If Me.txtPrimaryCurrTax.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                        CDec(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                    Else

                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                        CDec(Me.txtSubTotal.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                    End If
                End If

                Me.txtTaxAmount1.Text = Me.txtTaxAmount.Text 'new tax amount1
                If Me.cboPPNDipungut.SelectedIndex = 1 Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                End If
                'Else
                '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                'End If
            Else
                Me.txtTax.Text = "0.00"
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax) 'new tax amount1
            End If
            Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DefPrecTax)


            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.txtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotal.Text) Then
                    decTotal += CDec(Me.txtSubTotal.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.txtDutyStamp.Text) Then
                    decTotal += CDec(Me.txtDutyStamp.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
            Else
                If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotal.Text) Then
                        decTotal += CDec(Me.txtSubTotal.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.txtDutyStamp.Text) Then
                        decTotal += CDec(Me.txtDutyStamp.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotal.Text) Then
                        decTotal += CDec(Me.txtSubTotal.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                End If
            End If

            '2025 Core Tax Tahap 2.2
            'DPP Nilai Lain untuk 04, sementara 11/12, tapi bisa diedit
            Me.UbahDPPNilaiLain()

            ''hitung pajak
            'If Me.cboPPNDipungut.SelectedIndex = 0 Then
            '    If Me.txtPrimaryCurrency.Text = 1 Then
            '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * _
            '        CDbl(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrec)
            '    Else
            '        'perhitungan pajak dengan exchange rate
            '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Math.Round(CDec(Me.txtSubTotal.Text) * _
            '        CDec(Me.txtTaxExchangeRate.Text), DefPrec, MidpointRounding.AwayFromZero) * _
            '        CDec(Me.txtTax.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrec)
            '    End If
            'Else
            '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrec)
            'End If

            ''hitung total = sub total + tax
            'If Me.txtPrimaryCurrency.Text = "1" Then
            '    'kalo IDR, duty stamp diperhitungkan
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) + _
            '    CDbl(Me.txtTaxAmount.Text) + CDbl(Me.txtDutyStamp.Text), DefaultPrecision))
            'Else
            '    'kalo bukan IDR duty stamp dan tax tidak diperhitungkan
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) , DefaultPrecision))
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.2
    Private Sub UbahDPPNilaiLain()
        Try
            'Update SMESS, 2026-10-12
            Me.txtDPPNilaiLain.ReadOnly = True
            Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotal.Text) * 11 / 12, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))

            ''DPP Nilai Lain untuk 04, sementara 11/12, tapi bisa diedit
            'If Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax04 Then
            '    Me.txtDPPNilaiLain.ReadOnly = True
            '    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotal.Text) * 11 / 12, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'ElseIf Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax05 Then
            '    'DPP Nilai Lain untuk 05, sementara 10% dari total DPP, bisa diedit
            '    Me.txtDPPNilaiLain.ReadOnly = True
            '    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotal.Text) * 100 / 100, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'Else
            '    Me.txtDPPNilaiLain.ReadOnly = True
            '    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotal.Text), Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
            End If
            Me.Hitung_Total()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' melakukan perubahan currency setelah so dituliskan
    ''' </summary>
    ''' <param name="vPrimarySO"></param>
    ''' <remarks></remarks>
    Private Sub SetCurrency(ByVal vPrimarySO As Double)
        Dim PrimaryCur As Double = -1
        Try
            PrimaryCur = Me.ObjSO.GetPrimaryCurrency(vPrimarySO)


            If PrimaryCur <> -1 Then
                Me.txtPrimaryCurrency.Text = PrimaryCur
                Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision()
            Else
                Me.txtPrimaryCurrency.Text = -1
                Me.txtCurrencyCode.Text = ""
                Me.SetDefaultPrecision()
            End If
            If Me.ObjSO.GetUseTax(vPrimarySO) = True Then
                Me.chkTax.Checked = True
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.ObjSO.GetTaxPersentage(vPrimarySO), 2)
                Me.Hitung_Total()
            Else
                Me.chkTax.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                Me.txtSubTotal.Focus()
                Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                Me.SetDefaultPrecision()
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
                    Me.txtSubTotal.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetDefaultPrecision()
                End If
            End If

            Me.Hitung_Total()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetDefaultPrecision()
        Try
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
            Me.Hitung_Total()
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim UseTaxSetting, TaxPercentSetting As String

        Try
            If Me.chkTax.Checked = True Then
                Me.txtTax.ReadOnly = False
                Me.txtTax.TabStop = True
                Me.txtTax.BackColor = Me.TxtPrimaryInvoice.BackColor

                UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.UseTax)
                If UseTaxSetting = "" Then
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
                Else
                    If UseTaxSetting = "0" Then
                        Me.chkTax.Checked = False
                    Else
                        Me.chkTax.Checked = True
                        TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                        If IsNumeric(TaxPercentSetting) = True Then
                            Me.txtTax.Text = TaxPercentSetting
                        End If
                    End If
                End If

                '2025 Core Tax Tahap 2.2
                'txtNoFakturPajakSederhana.Visible = False
                'txtNoSeriFakturPajak.Visible = True
                'Label7.Visible = True
                'txtNoFakturPajak.Visible = True
                'llbEditNoFakturPajak.Visible = True

            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.TabStop = False
                Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                '2025 Core Tax Tahap 2.2
                'txtNoFakturPajakSederhana.Visible = True
                'txtNoSeriFakturPajak.Visible = False
                'Label7.Visible = False
                'txtNoFakturPajak.Visible = False
                'llbEditNoFakturPajak.Visible = False
            End If

            If Me.TxtRemarks.Enabled = True Then
                Me.SetTaxRate(True)
                Me.Hitung_Total()

                If Me.chkTax.Checked = True Then
                    '2025 Core Tax Tahap 2.2
                    'Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    'Me.llbEditNoFakturPajak.Enabled = True
                    Me.cboTaxtype.SelectedIndex = 0
                    Me.cboTaxtype.Enabled = True
                    Me.cboTaxtype.TabStop = True
                    Me.cboPPNDipungut.Enabled = True
                    Me.cboPPNDipungut.TabStop = True
                    '//#2732
                    Me.txtRevisedInvoice.Enabled = True
                    Me.txtNoFakturPajak.Enabled = True
                Else
                    Me.llbEditNoFakturPajak.Enabled = False
                    '2025 Core Tax Tahap 2.2
                    'Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    'If Me.txtNoFakturPajak.Text.Length >= 8 Then
                    '    Me.txtNoFakturPajak.Text = Me.txtNoFakturPajak.Text.Substring(0, 8) + "00000000"
                    'End If
                    Me.cboTaxtype.SelectedIndex = 0
                    Me.cboTaxtype.Enabled = False
                    Me.cboTaxtype.TabStop = False
                    Me.cboPPNDipungut.Enabled = False
                    Me.cboPPNDipungut.TabStop = False
                    '//#2732
                    Me.txtRevisedInvoice.Enabled = False
                    Me.txtNoFakturPajak.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCustomerAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustomerAddress.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.cboCustomerAddress.SelectedIndex <> -1 Then
                    Me.txtCustomerAddIdentifier.Text = Me.ObjCustomer.GetIdentifierAdd(Me.txtPrimaryCustomer.Text, _
                    Me.cboCustomerAddress.Text)
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

    Private Sub TxtCustomerName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerName.TextChanged
        Try
            If Me.txtPrimaryCustomer.Text <> "-1" And Me.txtPrimaryCustomer.Text <> "" Then
                If Me.TxtCustomerName.Text <> "" Then
                    Me.Text = "Down Payment Invoice: " + Me.txtCustomerCode.Text + " - " + Me.TxtCustomerName.Text
                Else
                    Me.Text = "Down Payment Invoice"
                End If
            Else
                Me.Text = "Down Payment Invoice"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
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
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDutyStamp.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtDutyStamp, sender, e)
    End Sub

    Private Sub txtDutyStamp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDutyStamp.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtDutyStamp, sender, e)
    End Sub

    Private Sub txtTaxExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtTaxExchangeRate, sender, e)
    End Sub

    Private Sub txtTaxExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtTaxExchangeRate, sender, e)
    End Sub

    Private Sub txtSalesOrderNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSONo.KeyDown
        Dim i As New FrmCariTransaksiSO
        Dim Jml As Long = 0
        '2025 Core Tax Tahap 2.1
        Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSO.ApakahSudahAda(Me.txtSONo.Text) = True Then
                    Me.txtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.txtSONo.Text)
                    Me.txtPrimaryCustomer.Text = Me.ObjSO.GetPrimaryCustomer(Me.txtPrimarySO.Text)
                    Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                    Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    '2025 Core Tax Tahap 2.1
                    '//TR : Tambahan baru Nik/Pas
                    '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                    End If
                    '2025 Core Tax Tahap 2.1
                    If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                    Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                    Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                        'Me.llbUpDateNPWP.Visible = False
                        Me.cboNIKPASS.Visible = True
                        Me.cboNIKPASS.TabStop = True
                        Me.txtNoNIKPass.Visible = True
                        Me.txtNoNPWP.Visible = False
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                        Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                        If Me.baru = True Then
                            Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNIKPass.Text = Me.objInvoiceDP.GetNoNikPassbyInvoice(Me.TxtPrimaryInvoice.Text)
                        End If
                    Else
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If

                    Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                    Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                    Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                    If Me.cboCustomerAddress.Items.Count > 0 Then
                        Me.cboCustomerAddress.SelectedIndex = 0
                    End If
                    Me.cboCustomerAddress.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtSONo.Text <> "" Then
                        i.txtCari.Text = Me.txtSONo.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryTransaksi > 0 Then
                        Me.txtPrimarySO.Text = i.PrimaryTransaksi
                        Me.txtSONo.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                        Me.txtPrimaryCustomer.Text = Me.ObjSO.GetPrimaryCustomer(Me.txtPrimarySO.Text)
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                        Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        '//TR : Tambahan baru Nik/Pas
                        '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                        If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                            NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                        End If
                        '2025 Core Tax Tahap 2.1
                        If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                        Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                        Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                            'Me.llbUpDateNPWP.Visible = False
                            Me.cboNIKPASS.Visible = True
                            Me.cboNIKPASS.TabStop = True
                            Me.txtNoNIKPass.Visible = True
                            Me.txtNoNPWP.Visible = False
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                            Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                            If Me.baru = True Then
                                Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                            Else
                                Me.txtNoNIKPass.Text = Me.objInvoiceDP.GetNoNikPassbyInvoice(Me.TxtPrimaryInvoice.Text)
                            End If
                        Else
                            'Me.llbUpDateNPWP.Visible = True
                            Me.cboNIKPASS.Visible = False
                            Me.cboNIKPASS.TabStop = False
                            Me.txtNoNIKPass.Visible = False
                            Me.txtNoNPWP.Visible = True
                        End If
                        Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                        Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                        Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                        If Me.cboCustomerAddress.Items.Count > 0 Then
                            Me.cboCustomerAddress.SelectedIndex = 0
                        End If
                        Me.cboCustomerAddress.Focus()
                    Else
                        Me.txtSONo.Clear()
                        Me.txtPrimarySO.Text = "-1"
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtNoSeriFakturPajak.Clear()
                        Me.txtSONo.Focus()
                        Me.Txt_E_Faktur_Jalan.Text = ""
                        Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                    End If
                End If
                Me.SetCurrency(Me.txtPrimarySO.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSalesOrderNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSONo.Leave
        Dim i As New FrmCariTransaksiSO
        '2025 Core Tax Tahap 2.1
        Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
        Try
            If Me.txtSONo.Text <> "" Then
                If Me.ObjSO.ApakahSudahAda(Me.txtSONo.Text) = True Then
                    Me.txtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.txtSONo.Text)
                    Me.txtPrimaryCustomer.Text = Me.ObjSO.GetPrimaryCustomer(Me.txtPrimarySO.Text)
                    Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                    Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    '2025 Core Tax Tahap 2.1
                    '//TR : Tambahan baru Nik/Pas
                    '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                    End If
                    '2025 Core Tax Tahap 2.1
                    If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                    Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                    Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                        'Me.llbUpDateNPWP.Visible = False
                        Me.cboNIKPASS.Visible = True
                        Me.cboNIKPASS.TabStop = True
                        Me.txtNoNIKPass.Visible = True
                        Me.txtNoNPWP.Visible = False
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                        Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                        If Me.baru = True Then
                            Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNIKPass.Text = Me.objInvoiceDP.GetNoNikPassbyInvoice(Me.TxtPrimaryInvoice.Text)
                        End If
                    Else
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If
                    Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                    Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                    Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                    If Me.cboCustomerAddress.Items.Count > 0 Then
                        Me.cboCustomerAddress.SelectedIndex = 0
                    End If
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.txtSONo.Text <> "" Then
                        i.txtCari.Text = Me.txtSONo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        Me.txtPrimarySO.Text = i.PrimaryTransaksi
                        Me.txtSONo.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                        Me.txtPrimaryCustomer.Text = Me.ObjSO.GetPrimaryCustomer(Me.txtPrimarySO.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                        Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        '//TR : Tambahan baru Nik/Pas
                        '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                        If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                            NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                        End If
                        '2025 Core Tax Tahap 2.1
                        If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                        Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                        Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                            'Me.llbUpDateNPWP.Visible = False
                            Me.cboNIKPASS.Visible = True
                            Me.cboNIKPASS.TabStop = True
                            Me.txtNoNIKPass.Visible = True
                            Me.txtNoNPWP.Visible = False
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                            Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                            If Me.baru = True Then
                                Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                            Else
                                Me.txtNoNIKPass.Text = Me.objInvoiceDP.GetNoNikPassbyInvoice(Me.TxtPrimaryInvoice.Text)
                            End If
                        Else
                            'Me.llbUpDateNPWP.Visible = True
                            Me.cboNIKPASS.Visible = False
                            Me.cboNIKPASS.TabStop = False
                            Me.txtNoNIKPass.Visible = False
                            Me.txtNoNPWP.Visible = True
                        End If
                        Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                        Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                        Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                        If Me.cboCustomerAddress.Items.Count > 0 Then
                            Me.cboCustomerAddress.SelectedIndex = 0
                        End If
                    Else
                        Me.txtSONo.Clear()
                        Me.txtPrimarySO.Text = "-1"
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtNoSeriFakturPajak.Clear()
                        Me.txtSONo.Focus()
                        Me.Txt_E_Faktur_Jalan.Text = ""
                        Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                    End If
                End If
            Else
                Me.txtSONo.Clear()
                Me.txtPrimarySO.Text = "-1"
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtNoSeriFakturPajak.Clear()
                Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, True, True)
                Me.Txt_E_Faktur_Jalan.Text = ""
            End If
            Me.SetCurrency(Me.txtPrimarySO.Text)
            Me.SetSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesOrder.Click
        Dim i As New FrmCariTransaksiSO
        '2025 Core Tax Tahap 2.1
        Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
        Try
            i.BackColorAll = Me.BackColorAll

            If Me.txtSONo.Text <> "" Then
                i.txtCari.Text = Me.txtSONo.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.txtPrimarySO.Text = i.PrimaryTransaksi
                Me.txtSONo.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                Me.txtPrimaryCustomer.Text = Me.ObjSO.GetPrimaryCustomer(Me.txtPrimarySO.Text)
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                '2025 Core Tax Tahap 2.1
                '//TR : Tambahan baru Nik/Pas
                '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    '2025 Core Tax Tahap 2.1
                    Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                    NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                Else
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                End If
                '2025 Core Tax Tahap 2.1
                If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                    'Me.llbUpDateNPWP.Visible = False
                    Me.cboNIKPASS.Visible = True
                    Me.cboNIKPASS.TabStop = True
                    Me.txtNoNIKPass.Visible = True
                    Me.txtNoNPWP.Visible = False
                    '2025 Core Tax Tahap 2.1
                    Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                    Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                    If Me.baru = True Then
                        Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNIKPass.Text = Me.objInvoiceDP.GetNoNikPassbyInvoice(Me.TxtPrimaryInvoice.Text)
                    End If
                Else
                    'Me.llbUpDateNPWP.Visible = True
                    Me.cboNIKPASS.Visible = False
                    Me.cboNIKPASS.TabStop = False
                    Me.txtNoNIKPass.Visible = False
                    Me.txtNoNPWP.Visible = True
                End If
                Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                If Me.cboCustomerAddress.Items.Count > 0 Then
                    Me.cboCustomerAddress.SelectedIndex = 0
                End If
                Me.cboCustomerAddress.Focus()
            Else
                Me.txtSONo.Clear()
                Me.txtPrimarySO.Text = "-1"
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtNoSeriFakturPajak.Clear()
                Me.txtSONo.Focus()
                Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                Me.Txt_E_Faktur_Jalan.Text = ""
            End If

            Me.SetCurrency(Me.txtPrimarySO.Text)
            Me.SetSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetSalesPerson()
        Try
            Me.TxtPrimarySalesPerson.Text = Me.ObjSO.GetPrimarySalesPerson(Me.txtPrimarySO.Text)
            Me.TxtSalesPersonCode.Text = Me.ObjSalesPerson.GetCode(Me.TxtPrimarySalesPerson.Text)
            Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoSeriFakturPajak_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoSeriFakturPajak.ReadOnly = True Then
                    Me.txtNoSeriFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtNoSeriFakturPajak.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoSeriFakturPajak_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoSeriFakturPajak.ReadOnly = True Then
                    Me.txtNoSeriFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtNoSeriFakturPajak.BackColor = Me.TxtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboFakturPajak_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboFakturPajak.SelectedIndexChanged
        Try
            If Me.CboFakturPajak.SelectedIndex = 0 Then
                'faktur pajak sederhana
                Me.txtNoFakturPajakSederhana.Visible = True
                Me.txtNoFakturPajakSederhana.TabStop = True
                Me.txtNoFakturPajakSederhana.BringToFront()
                Me.txtNoFakturPajakSederhana.BackColor = Me.TxtPrimaryInvoice.BackColor
                Me.txtNoSeriFakturPajak.TabStop = False
                'Me.txtNoFakturPajak.TabStop = False '//#2732
            Else
                'faktur pajak standar
                Me.txtNoFakturPajakSederhana.Visible = False
                Me.txtNoFakturPajakSederhana.TabStop = False
                Me.txtNoFakturPajakSederhana.SendToBack()
                Me.txtNoSeriFakturPajak.TabStop = True
                'Me.txtNoFakturPajak.TabStop = False '//#2732

                '2025 Core Tax Tahap 2.2
                'If Me.baru = False And Me.txtNoFakturPajak.Text = "" Then
                '    'If Me.txtPrimaryCustomer.Text <> "-1" And Me.txtPrimaryCustomer.Text <> "" Then
                '    '    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                '    'End If
                '    'Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '    If Me.chkTax.Checked = True Then
                '        Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '        Me.llbEditNoFakturPajak.Enabled = True
                '    Else
                '        Me.llbEditNoFakturPajak.Enabled = False
                '        Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '        If Me.txtNoFakturPajak.Text <> "" Then
                '            Me.txtNoFakturPajak.Text = Me.txtNoFakturPajak.Text.Substring(0, 8) + "00000000"
                '        End If
                '    End If
                'End If
            End If

            
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajakSederhana_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajakSederhana.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoFakturPajakSederhana.ReadOnly = False Then
                    Me.txtNoFakturPajakSederhana.BackColor = Me.BackColorAll
                Else
                    Me.txtNoFakturPajakSederhana.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajakSederhana_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajakSederhana.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoFakturPajakSederhana.ReadOnly = False Then
                    Me.txtNoFakturPajakSederhana.BackColor = Me.txtPrimaryCurrency.BackColor
                Else
                    Me.txtNoFakturPajakSederhana.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
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

    Private Sub txtTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtTax.ReadOnly = False Then
                    Me.txtTax.BackColor = Me.txtPrimaryCurrency.BackColor
                Else
                    Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = False Then
                    Me.txtTaxExchangeRate.BackColor = Me.BackColorAll
                Else
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = False Then
                    Me.txtTaxExchangeRate.BackColor = Me.txtPrimaryCurrency.BackColor
                Else
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbEditNoInvoice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditNoInvoice.LinkClicked
        Dim i As New FrmTransaksiSalesInvoice_EditNo
        Dim Dummy As String = ""
        Dim PrimaryInvoice As Double = -1
        Dim JmlRujukan As Integer = 0
        Dim JmlRujukan1 As Integer = 0
        Dim JmlRujukan2 As Integer = 0
        Dim Dummy1 As String = ""

        Try
            If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryInvoice.Text) AndAlso CDec(Me.TxtPrimaryInvoice.Text)) Then
                i.BackColorAll = Me.BackColorAll
                JmlRujukan = Me.ObjNumbering.GetJmlRunningNumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice)
                While Dummy.Length < JmlRujukan
                    Dummy = Dummy + "0"
                End While
                i.mtb_Number.Mask = Dummy
                i.MinLength = JmlRujukan
                JmlRujukan1 = Me.ObjNumbering.GetUrutan(Me.txtInvoiceNo.Text, ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, Me.dtpDate.Value)

                i.mtb_Number.Text = JmlRujukan1
                While i.mtb_Number.Text.Length < JmlRujukan
                    i.mtb_Number.Text = "0" + i.mtb_Number.Text
                End While
                Dummy1 = i.mtb_Number.Text

                i.Label1.Text = "Invoice Number"
                i.Text = "Edit Invoice Number"

                'cek revisi
                JmlRujukan2 = Me.objInvoiceDP.GetNoBaru1(Me.dtpDate.Value).Length
                If JmlRujukan2 <= Me.txtInvoiceNo.Text.Length Then
                    i.TxtRevision.Text = Me.txtInvoiceNo.Text.Substring(JmlRujukan2)
                Else
                    i.TxtRevision.Text = ""
                End If

                If Me.TxtRemarks.Enabled = False Then
                    i.TxtRevision.Enabled = False
                    i.mtb_Number.Enabled = False
                End If

                If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dummy = Me.txtInvoiceNo.Text.Replace(Dummy1, i.mtb_Number.Text)
                    Dummy = Dummy.Substring(0, JmlRujukan2) + i.TxtRevision.Text

                    PrimaryInvoice = Me.objInvoiceDP.GetPrimaryInvoice(Dummy)
                    If PrimaryInvoice = -1 Then
                        Me.txtInvoiceNo.Text = Dummy
                    Else
                        If CDec(Me.TxtPrimaryInvoice.Text) = PrimaryInvoice Then
                            Me.txtInvoiceNo.Text = Dummy
                        Else
                            MsgBox("The invoice number " + Dummy + " is used. Please enter another number.", _
                            MsgBoxStyle.Information, "Information")
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

    Private Sub llbEditNoFakturPajak_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditNoFakturPajak.LinkClicked
        Try
            Me.ShowDetailFakturPajak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowDetailFakturPajak()
        Dim i As New FrmTransaksiSalesInvoice_EditNo
        Dim j As New FrmTransaksiSalesInvoice_EditNo2
        Dim Dummy As String = ""
        Dim PrimaryInvoice As Double = -1
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim StartingNumber As Double = 0
        Dim EndingNumber As Double = 0

        Try
            If Me.dtpDate.Value < Obj.BatasStartDate Then
                i.mtb_Number.Mask = "00000000"
                i.mtb_Number.Text = Me.txtNoFakturPajak.Text.Substring(3)
                i.MinLength = 8
                i.Text = "Edit Tax Invoice Number"
                If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dummy = Me.txtNoFakturPajak.Text.Substring(0, 3) + i.mtb_Number.Text
                    PrimaryInvoice = Me.objInvoiceDP.GetPrimaryInvoiceByNoFakturPajak(Dummy)
                    If PrimaryInvoice = -1 Then
                        Me.txtNoFakturPajak.Text = Dummy
                    Else
                        If CDbl(Me.TxtPrimaryInvoice.Text) = PrimaryInvoice Then
                            Me.txtNoFakturPajak.Text = Dummy
                        Else
                            MsgBox("The tax invoice number " + Dummy + " is used by " + _
                            Me.objInvoiceDP.GetNoInvoice(PrimaryInvoice) + ".", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            Else
                j.mtb_Number.Mask = "00000000"
                If Me.txtNoFakturPajak.Text.Length > 7 Then
                    j.mtb_Number.Text = Me.txtNoFakturPajak.Text.Substring(8)
                Else
                    j.mtb_Number.Text = ""
                End If
                j.PrimaryInvoice = Me.TxtPrimaryInvoice.Text
                j.InvDate = Me.dtpDate.Value
                j.primaryTax = Me.ObjTaxInvoiceSetting.GetPrimary(Me.dtpDate.Value, Me.txtNoFakturPajak.Text)
                j.MinLength = 8
                j.Text = "Edit Tax Invoice Number"

                j.Txt_E_Faktur_Blok.Text = Me.Txt_E_Faktur_Blok.Text
                j.Txt_E_Faktur_Jalan.Text = Me.Txt_E_Faktur_Jalan.Text
                j.Txt_E_Faktur_Kabupaten.Text = Me.Txt_E_Faktur_Kabupaten.Text
                j.Txt_E_Faktur_Kecamatan.Text = Me.Txt_E_Faktur_Kecamatan.Text
                j.Txt_E_Faktur_Kelurahan.Text = Me.Txt_E_Faktur_Kelurahan.Text
                j.Txt_E_Faktur_KodePos.Text = Me.Txt_E_Faktur_KodePos.Text
                j.Txt_E_Faktur_No.Text = Me.Txt_E_Faktur_No.Text
                j.Txt_E_Faktur_NoTelepon.Text = Me.Txt_E_Faktur_NoTelepon.Text
                j.Txt_E_Faktur_Propinsi.Text = Me.Txt_E_Faktur_Propinsi.Text
                j.Txt_E_Faktur_RT.Text = Me.Txt_E_Faktur_RT.Text
                j.Txt_E_Faktur_RW.Text = Me.Txt_E_Faktur_RW.Text

                If j.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dummy = j.Result
                    Me.txtNoFakturPajak.Text = Dummy

                    Me.Txt_E_Faktur_Blok.Text = j.Txt_E_Faktur_Blok.Text
                    Me.Txt_E_Faktur_Jalan.Text = j.Txt_E_Faktur_Jalan.Text
                    Me.Txt_E_Faktur_Kabupaten.Text = j.Txt_E_Faktur_Kabupaten.Text
                    Me.Txt_E_Faktur_Kecamatan.Text = j.Txt_E_Faktur_Kecamatan.Text
                    Me.Txt_E_Faktur_Kelurahan.Text = j.Txt_E_Faktur_Kelurahan.Text
                    Me.Txt_E_Faktur_KodePos.Text = j.Txt_E_Faktur_KodePos.Text
                    Me.Txt_E_Faktur_No.Text = j.Txt_E_Faktur_No.Text
                    Me.Txt_E_Faktur_NoTelepon.Text = j.Txt_E_Faktur_NoTelepon.Text
                    Me.Txt_E_Faktur_Propinsi.Text = j.Txt_E_Faktur_Propinsi.Text
                    Me.Txt_E_Faktur_RT.Text = j.Txt_E_Faktur_RT.Text
                    Me.Txt_E_Faktur_RW.Text = j.Txt_E_Faktur_RW.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub btnPrintTaxInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintTaxInvoice.Click
        Dim ij As New FrmCetakFakturDP

        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.CboFakturPajak.SelectedIndex = 0 Or Me.chkTax.Checked = False Then
                    MsgBox("This feature is only for standard tax invoice.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                ij.primary = Me.TxtPrimaryInvoice.Text
                ij.mParentForm = Me
                ij.NoFakturPajak = Me.txtNoSeriFakturPajak.Text + "-" + Me.txtNoFakturPajak.Text
                ij.CurrencyCode = Me.txtCurrencyCode.Text
                ij.TotalDitagih = Me.TxtTotal.Text
                ij.TotalPPN = Me.txtTaxAmount.Text
                ij.TotalHargaJual = Me.txtSubTotal.Text
                ij.RatePajak = Me.txtTaxExchangeRate.Text
                If Me.txtPrimaryCurrency.Text <> 1 Then
                    ij.FakturPajakStandarValas = True
                Else
                    ij.FakturPajakStandarValas = False
                End If
                ij.DecimalValas = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
                ij.CurrencyCodeValas = Me.txtCurrencyCode.Text

                If Me.CboFakturPajak.SelectedIndex = 0 Then
                    ij.JenisFaktur = FrmCetakFaktur.EnumJenisFakturPajak.Sederhana
                    'ij.Text = "Faktur Pajak Sederhana"
                Else
                    ij.JenisFaktur = FrmCetakFaktur.EnumJenisFakturPajak.Standar
                    'ij.Text = "Faktur Pajak Standar"
                End If
                ij.FormatDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            If Me.TxtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2)
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Leave
        Try
            If CDec(Me.txtTaxExchangeRate.Text) = 0 Then
                Me.txtTaxExchangeRate.Text = 1
            End If
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.txtTaxExchangeRate.Text, 2)

            If Me.TxtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNamaSigned_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNamaSigned.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtNamaSigned.ReadOnly = False Then
                    Me.TxtNamaSigned.BackColor = Me.BackColorAll
                Else
                    Me.TxtNamaSigned.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNamaSigned_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNamaSigned.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtNamaSigned.ReadOnly = False Then
                    Me.TxtNamaSigned.BackColor = Me.txtPrimaryCurrency.BackColor
                Else
                    Me.TxtNamaSigned.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJabatanSigned_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJabatanSigned.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJabatanSigned.ReadOnly = False Then
                    Me.TxtJabatanSigned.BackColor = Me.BackColorAll
                Else
                    Me.TxtJabatanSigned.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJabatanSigned_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJabatanSigned.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJabatanSigned.ReadOnly = False Then
                    Me.TxtJabatanSigned.BackColor = Me.txtPrimaryCurrency.BackColor
                Else
                    Me.TxtJabatanSigned.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUpDateNPWP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUpDateNPWP.LinkClicked
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If Me.txtPrimaryCustomer.Text <> "" And Me.txtPrimaryCustomer.Text <> "-1" Then
                If MsgBox("Update NPWP number from customer data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    '2024-08-26 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                    Else
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                    End If
                    Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, Me.objInvoiceDP.AlamatPajak1, Me.objInvoiceDP.AlamatPajak2)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtPaymentTerm, sender, e)
    End Sub

    Private Sub TxtPaymentTerm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtPaymentTerm, sender, e, False)
    End Sub

    Private Sub TxtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.GotFocus
        Me.TxtPaymentTerm.BackColor = BackColorAll
    End Sub

    Private Sub TxtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.LostFocus
        If Me.TxtPaymentTerm.Enabled = True And Me.TxtPaymentTerm.ReadOnly = False Then
            Me.TxtPaymentTerm.BackColor = Me.TxtPrimaryInvoice.BackColor
            Me.Hitung_Total()
        End If
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
            If Me.TxtRemarks.Enabled = True Then
                If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If Me.baru = False Then
                        If Me.objInvoiceDP.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objInvoiceDP.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                            'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                            Me.txtInvoiceNo.Text = Me.objInvoiceDP.NoTransaksi
                        Else
                            'jika pindah bulan maka akan ganti no
                            Me.txtInvoiceNo.Text = Me.objInvoiceDP.GetNoBaru1(Me.dtpDate.Value)
                        End If
                    Else
                        Me.txtInvoiceNo.Text = Me.objInvoiceDP.GetNoBaru1(Me.dtpDate.Value)
                    End If
                End If

                '2025 Core Tax Tahap 2.2, di comment
                'If Me.dtpDate.Value.Date < CDate(#4/1/2013#) Then
                '    'update nomor tahun pada faktur pajak standar
                '    If Me.CboFakturPajak.SelectedIndex = 1 Then
                '        Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '    End If
                'Else
                '    If Me.chkTax.Checked = True Then
                '        If Me.CboFakturPajak.SelectedIndex = 1 Then
                '            If Me.baru = True Then
                '                Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value.Date)
                '            Else
                '                If Me.objInvoiceDP.TanggalTransaksi >= #4/1/2013# Then
                '                    Me.txtNoFakturPajak.Text = Me.objInvoiceDP.NoFakturPajak
                '                Else
                '                    Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value.Date)
                '                End If
                '            End If
                '        End If
                '        Me.llbEditNoFakturPajak.Enabled = True
                '    Else
                '        Me.llbEditNoFakturPajak.Enabled = False
                '        Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '        If Me.txtNoFakturPajak.Text <> "" Then
                '            Me.txtNoFakturPajak.Text = Me.txtNoFakturPajak.Text.Substring(0, 8) + "00000000"
                '        End If
                '    End If
                'End If

                Me.SetTaxRate(True)
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Enter
        Try
            If Me.txtTaxExchangeRate.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTaxExchangeRate, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTaxRate(ByVal vSetTaxRate As Boolean)
        Dim Rate As Decimal = 0

        Try
            If Me.txtPrimaryCurrency.Text = 1 Then
                Me.txtTaxExchangeRate.Enabled = False
                Me.txtTaxExchangeRate.ReadOnly = True
                Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                Me.txtTaxExchangeRate.TabStop = False
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Me.ObjCurrency.GetDefaultPrecission(1))
            Else
                Me.txtTaxExchangeRate.Enabled = True
                Me.txtTaxExchangeRate.ReadOnly = False
                Me.txtTaxExchangeRate.BackColor = Me.txtPrimaryCurrency.BackColor
                Me.txtTaxExchangeRate.TabStop = True
                If vSetTaxRate = True Then
                    Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                    If Rate = 0 Then
                        Rate = 1
                    End If
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Me.ObjTaxRate.DefaultPrec)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.Enter
        Try
            If Me.txtDutyStamp.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtDutyStamp, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub txtSubTotal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSubTotal.TextChanged

    End Sub

    Private Sub cboPPNDipungut_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPPNDipungut.SelectedValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemarks.Enabled = False Then
                If Me.objInvoiceDP.CariDetail(vPrimary) = True Then
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

    Public Sub FromCutOff_SalesInvoiceDP_Create(ByVal vPrimaryTransaksi As Double, ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core, _
    ByVal vTblPrimary As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable)
        Try
            Me.baru = False
            Me.objInvoiceDP.ResetObject()
            Me.objInvoiceDP.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP

            Me.objInvoiceDP.CariDetail(vPrimaryTransaksi)
            Me.objInvoiceDP.PrimarySO4DP = ObjCutOff_Core.GetPrimaryUpdate(Me.objInvoiceDP.PrimarySO4DP, vTblPrimary)
            Me.UserCode = Me.objInvoiceDP.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_SalesInvoiceDP_Save(Optional ByVal vApakahSimpanJurnal As Boolean = True) As Double
        Try
            Me.baru = True
            Me.objInvoiceDP.SimpanBaru(Me.UserCode, False, vApakahSimpanJurnal, True)
            Return Me.objInvoiceDP.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    'aul
    Private Sub txtNoSeriFakturPajak_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.objInvoiceDP.SetDontCollect(Me.txtNoSeriFakturPajak.Text, Me.chkTax.Checked, Me.cboPPNDipungut.SelectedIndex) = True Then
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.2
    Private Sub LlbViewDetailTax_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewDetailTax.LinkClicked
        Try
            Me.ShowInfoPajak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ShowInfoPajak() As Boolean
        Dim ij As New FrmTransaksiSalesInvoice_InfoFakturPajak
        Try
            ij.InvoiceDate = Me.dtpDate.Value
            ij.PrimarySalesInvoice = Me.txtPrimaryInvoice.Text
            ij.UserCode = Me.UserCode
            ij.BackColorAll = Me.BackColorAll
            ij.InvoiceNo = Me.txtInvoiceNo.Text
            '2025 Core Tax Tahap 2.2, update ketika semua type bisa dipilih, tanpa percabangan
            'If Me.cboTaxType.SelectedIndex = 0 Then
            '    ij.VATType = ClsTransaksiSalesInvoice.EnumTaxType.Tax04
            'ElseIf Me.cboTaxType.SelectedIndex = 1 Then
            '    ij.VATType = ClsTransaksiSalesInvoice.EnumTaxType.Tax08
            'End If
            ij.VATType = Me.cboTaxType.SelectedIndex
            If Me.AddInfo <> -1 Then
                ij.PrimaryAddInfo = Me.AddInfo
            End If
            If Me.FacilityStamp <> -1 Then
                ij.PrimaryFacilityStamp = Me.FacilityStamp
            End If
            ij.ShowDialog()
            If ij.Simpan Then
                Me.SPPBDate = ij.dtpSPPBDate.Value
                Me.AddInfo = ij.cboTaxAddInfo.SelectedValue
                Me.FacilityStamp = ij.cboTaxFacilityStamp.SelectedValue
                'Me.TxtSPPBNo.Text = ij.TxtSPPBNo.Text
            End If
            Return ij.Simpan
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cboTaxType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTaxtype.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If cboTaxType.SelectedIndex <> ClsTransaksiSalesInvoice.EnumTaxType.Tax07 And cboTaxType.SelectedIndex <> ClsTransaksiSalesInvoice.EnumTaxType.Tax08 Then
                    Me.LlbViewDetailTax.Enabled = False
                Else
                    Me.LlbViewDetailTax.Enabled = True
                    If Me.objInvoiceDP.SetDontCollect1(Me.cboTaxType.SelectedIndex, Me.chkTax.Checked, Me.cboPPNDipungut.SelectedIndex) = True Then
                        Me.cboPPNDipungut.SelectedIndex = 1
                    End If
                End If
                '2025 Core Tax Tahap 2.2, tambahkan hitung total
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtDPPNilaiLain, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtDPPNilaiLain.Enabled = True Then
                    If Me.txtDPPNilaiLain.ReadOnly = True Then
                        Me.txtDPPNilaiLain.BackColor = Me.TxtReadOnly.BackColor
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
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyDown(Me.txtDPPNilaiLain, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDPPNilaiLain.KeyUp
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyUp(Me.txtDPPNilaiLain, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(Me.txtDPPNilaiLain.Text, Me.DefaultPrecision)
                'Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtDPPNilaiLain.Enabled = True Then
                    If Me.txtDPPNilaiLain.ReadOnly = True Then
                        Me.txtDPPNilaiLain.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtDPPNilaiLain.BackColor = Me.txtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TaxCodeDefaultDariCustomer(ByVal vCode As String)
        Try
            If vCode.Length > 3 Then
                vCode = vCode.Substring(0, 3)
            End If

            Select Case vCode
                Case "01", "010"
                    Me.cboTaxtype.SelectedIndex = 0
                Case "02", "020"
                    Me.cboTaxtype.SelectedIndex = 1
                Case "03", "030"
                    Me.cboTaxtype.SelectedIndex = 2
                Case "04", "040"
                    Me.cboTaxtype.SelectedIndex = 3
                Case "05", "050"
                    Me.cboTaxtype.SelectedIndex = 4
                Case "06", "060"
                    Me.cboTaxtype.SelectedIndex = 5
                Case "07", "070"
                    Me.cboTaxtype.SelectedIndex = 6
                Case "08", "080"
                    Me.cboTaxtype.SelectedIndex = 7
                Case "09", "090"
                    Me.cboTaxtype.SelectedIndex = 8
                Case "10", "100"
                    Me.cboTaxtype.SelectedIndex = 9
                Case Else
                    Me.cboTaxtype.SelectedIndex = 0
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '2025 Core Tax Tahap 2.2 end

    Private Sub txtRevisedInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.Click
        Try '//#2732
            Me.ObjCD.Number_Click(Me.txtRevisedInvoice)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.Leave
        Try '//#2732
            If Me.txtRevisedInvoice.Text = "" Or _
            Me.txtRevisedInvoice.Text = "0" Then
                Me.txtRevisedInvoice.Text = "00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.Enter
        Try '//#2732
            Me.ObjCD.Number_Enter(txtRevisedInvoice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.GotFocus
        Try '//#2732
            Me.txtRevisedInvoice.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.LostFocus
        Try '//#2732
            Me.txtRevisedInvoice.BackColor = Me.txtPrimaryInvoice.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajak_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajak.GotFocus
        Try '//#2732
            Me.txtNoFakturPajak.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajak_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajak.LostFocus
        Try '//#2732
            Me.txtNoFakturPajak.BackColor = Me.TxtPrimaryInvoice.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub txtRevisedInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRevisedInvoice.KeyPress
        Dim txt As TextBox = CType(sender, TextBox)
        Dim hasil As String = txt.Text & e.KeyChar
        Try '//#2732
            ' Izinkan backspace
            If Char.IsControl(e.KeyChar) Then Exit Sub

            ' Hanya angka
            If Not Char.IsDigit(e.KeyChar) Then
                e.Handled = True
                Exit Sub
            End If
            ' Maksimal 2 digit dan hanya boleh 00 atau 01
            If hasil <> "0" AndAlso hasil <> "00" AndAlso hasil <> "01" Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class