Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiAOPDP
    Public Lokal As Boolean = True
    Public UserCode As String = ""
    Public BackColorAll As Color = Me.BackColorAll

    Dim baru As Boolean
    Dim objInvoiceDP As New ClsTransaksiAOP
    Dim ObjBBDP As New Core_Module.ClsTransaksiBalanceDP
    Dim ObjCD As New ClsCekData
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjProduct As New clsMasterProduct
    'Dim ObjCustomer_sProductCode As New clsTransaksiSettingSupplierProductCodeNo
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjSupplier As New clsMasterSupplier
    Dim objInvoiceDPItem As New ClsTransaksiAOPItem
    Dim ObjCL As New ClsTransaksiMonthlyClosing
    Dim ObjLicense As New ClsLicense
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjTaxRate As New Sales_Module.ClsTransaksiSalesInvoiceTaxRate
    Dim ObjAOP As New ClsTransaksiAOP
    Dim ObjClosingAcc As New AccModule.ClsTransaksiClosing

    Dim DefaultPrecision As Integer = 0
    Dim buffer As String = ""
    'Dim DefPrec As Integer = 0

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

    Dim DsAccountDP As New AccModule.DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccountDP As New AccModule.DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccountDP As AccModule.DseMultiAccount.tblisi_cbo_ar_akunRow


    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged
        Try
            If Me.TxtRemark.Enabled = True AndAlso _
            IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso CDec(Me.txtPrimaryCurrency.Text) > -1 Then
                If Me.TxtRemark.Enabled = True Then
                    Me.IsiCBOArAkunByTipePerkiraan1()
                    Me.IsiCBOCurrencyTax()
                    Me.IsiCBOAkunDP()
                    'ini ditambahkan untuk mencegah akun dp yang dipilih ke item pertama
                    'tidak masalah ditaro sini karena, apapun primary currencynya, akun dpnya tetap. kecuali one day berubah
                    Me.txtPrimaryDPAccount.Text = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
                End If
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
                Me.Hitung_Total()
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


    Private Sub cboDPAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDPAccount.SelectedIndexChanged

        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryDPAccount.Text) AndAlso Me.cboPrimaryDPAccount.Text > -1 Then
                Me.txtPrimarydpaccount.Text = Me.cboPrimaryDPAccount.Text
            End If

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

    Private Sub IsiCBOAkunDP()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Other_Receivable
        Dim objMulti As New AccModule.ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboDPAccount.DataSource = Nothing
            Me.cboDPAccount.Items.Clear()
            Me.cboPrimaryDPAccount.DataSource = Nothing
            Me.cboPrimaryDPAccount.Items.Clear()
            Me.tblAccountDP.Rows.Clear()

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccountDP, 1)

            Me.cboDPAccount.DataSource = Me.tblAccountDP.DefaultView
            Me.cboDPAccount.DisplayMember = "Account Code"
            Me.cboPrimaryDPAccount.DataSource = Me.tblAccountDP.DefaultView
            Me.cboPrimaryDPAccount.DisplayMember = "Primary"
            Me.txtPrimaryDPAccount.Text = Me.cboPrimaryDPAccount.Text


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim UseTaxSetting, TaxPercentSetting As String
        Dim Dummy As String = ""
        Try
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP) = "" Then
                MsgBox("Please specify, AOP DP account first", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.objInvoiceDP.ResetObject()
            Me.TxtAOPNo.Text = Me.objInvoiceDP.GetNoBaru(Me.dtpDate.Value, -1)
            '     Me.txtNoFakturPajak.Text = Me.objInvoiceDP.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
            Me.objInvoiceDP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP_DP

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'setting standar untuk currency
            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
            Me.SetDefaultPrecision()
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiCBOCurrencyTax()
            Me.IsiCBOAkunDP()
            Me.txtPrimaryDPAccount.Text = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.chkTax.Checked = False
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
            Else
                Me.chkTax.Checked = True
                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
            End If

            'Me.chkTax.Checked = True
            'Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", Me.ObjCurrency.GetDecimal_DefaultCurrency)
            Me.SetDefaultPrecision()

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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryAOP.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.cboAccount.Enabled = True
            Me.cboDPAccount.Enabled = True
            Me.cboCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
                Me.IsiCBOAkunDP()
                Me.txtPrimaryDPAccount.Text = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
            End If
            Me.SetTaxRate(Me.baru)
            'Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit

            ' Me.llbAddress.Enabled = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.cboAccount.Enabled = False
            Me.cboDPAccount.Enabled = False
            Me.cboCurrTax.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            Me.TxtPrimaryAOP.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimarySupplier.Text = "-1"
            Me.txtPrimaryAccount.Text = "-1"
            Me.txtPrimarydpAccount.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtTax.Text = "0"
            Me.txtTaxExchangeRate.Text = "1"
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.txtDutyStamp.Text = "0"
            Me.txtSubTotal.Text = "0"
            Me.TxtPaymentsTerms.Text = "0"
            Me.txtExchangeRateToUSD.Text = "1"
            Me.txtExhangeRateToIDR.Text = "1"
            Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

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

            Me.chkTax.Checked = True
            ''Me.cboCustomerAddress.DataSource = Nothing
            ''Me.cboCustomerAddress.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim TglPO As Date = Today
        Dim POSubTotal As Decimal = 0
        Dim AOP_DP As Decimal = 0
        Dim Dummy As String = ""
        Dim ObjAkun As New ClsMasterAkun1
        Dim TglClosingAcc As Date = Today

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("po number")
            KoleksiTextBox.Add(Me.txtPONo)
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

            'berkaitan dengan tanggal closing
            'TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            'If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
            '    MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If
            TglClosingAcc = ObjClosingAcc.GetClosingDate
            If Me.dtpDate.Value.Date <= TglClosingAcc.Date Then
                MsgBox("You cannot save this data because transaction date <= accounting closing date(" + Format(TglClosingAcc, "dd MMM yyyy") + ")" _
                    , MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.objInvoiceDP.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objInvoiceDP.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtAOPNo.Text = Me.objInvoiceDP.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objInvoiceDP.TanggalTransaksi
                        Me.txtPONo.Text = Me.objInvoiceDP.NoTransaksi
                        Return False
                    Else
                        Me.TxtAOPNo.Text = Me.objInvoiceDP.GetNoBaru(Me.dtpDate.Value, CDbl(Me.txtPrimarySupplier.Text))
                    End If
                End If
            End If

            'berkaitan dengan tax exchange rate
            If Me.txtPrimaryCurrency.Text = "1" Then
                If CDbl(Me.txtTaxExchangeRate.Text) <> 1 Then
                    MsgBox("For downpayment with " + Me.ObjCurrency.GetCurrencyCode(1) + " currency, the tax exchange rate value must be 1.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'jika bukan IDR dan tax ratenya bernilai 1 maka berikan peringatan
                If CDbl(Me.txtTaxExchangeRate.Text) = 1 Then
                    MsgBox("This downpayment use non-" + Me.ObjCurrency.GetCurrencyCode(1) + " currency and the tax exchange rate value = 1. This is wrong data input.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'cek tanggalnya, minimum sama dengan tanggal PO
            TglPO = Me.ObjPO.GetTanggalPO(Me.txtPrimaryPO.Text)
            If Me.dtpDate.Value.Date < TglPO.Date Then
                MsgBox("Minimum down payment date is " + Format(TglPO, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                Me.dtpDate.Focus()
                Return False
            End If

            'nilai dp maksimum adalah sama dengan nilai po
            POSubTotal = Me.ObjPO.GetPOSubTotal(Me.txtPrimaryPO.Text)
            AOP_DP = Me.objInvoiceDP.GetSubTotal(Me.txtPrimaryPO.Text, Me.TxtPrimaryAOP.Text)
            AOP_DP += Me.ObjBBDP.GetJml(ClsTransaksiBalanceDP.EnumTipeDP.DP_AOP, Me.txtPrimaryPO.Text)

            If POSubTotal - AOP_DP - CDec(Me.txtSubTotal.Text) < 0 Then
                MsgBox("Maximum downpayment amount is " + Me.txtCurrencyCode.Text + " " + _
                Me.ObjCD.DigitGrouping(POSubTotal - AOP_DP, Me.DefaultPrecision), MsgBoxStyle.Information, "Information")

                Return False
            End If

            'nilai dp harus positif
            If Me.ObjCD.CekBilanganPositif(Me.txtSubTotal.Text) = False Then
                MsgBox("Please specify a positive number form down payment value", MsgBoxStyle.Information, "Information")
                Me.txtSubTotal.Focus()
                Return False
            End If

            'kalo pake pajak, maka no seri faktur pajak harus diisikan
            If Me.chkTax.Checked = True Then
                'kalo ada VAT, maka akun VAT out harus ada
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify VAT in account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                TxtTaxInvoiceNumber.Text = TxtTaxInvoiceNumber.Text.Trim
                If Me.TxtTaxInvoiceNumber.Text.StartsWith("010") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("020") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("030") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("040") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("060") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("070") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("080") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("090") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("011") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("021") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("031") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("041") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("061") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("071") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("081") = False And _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("091") = False Then
                    MsgBox("The tax invoice serial number must starts with one of the following alternatives: 010, 011, 020, 021, 030, 031, 040, 041, 060, 061, 070, 071, 080, 081, 091 or 090.", _
                    MsgBoxStyle.Information, "Information")
                    TxtTaxInvoiceNumber.Focus()
                    Return False
                End If

                If Me.TxtTaxInvoiceNumber.Text.StartsWith("070") = True Or _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("080") = True Or _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("071") = True Or _
               Me.TxtTaxInvoiceNumber.Text.StartsWith("081") = True Then
                    Dim ijk As New Sales_Module.FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur
                    ijk.TxtTaxInvoiceSeries.Text = Me.TxtTaxInvoiceNumber.Text.Substring(0, 3)
                    ijk.vIndexAcuan = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
                    ijk.ShowDialog()
                    Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = ijk.cboKeteranganTambahan.SelectedIndex
                End If
            End If

            'cek akun duty stamp
            If CDbl(Me.txtDutyStamp.Text) > 0 Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify dutystamp expense account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If


            'Account payable dan currency tax tidak boleh berbeda dengan data aop
            If IsNumeric(Me.txtPrimaryPO.Text) AndAlso _
               IsNumeric(Me.txtPrimaryAccount.Text) AndAlso _
               IsNumeric(Me.txtPrimaryCurrTax.Text) Then
                If Me.ObjAOP.GetJmlAOPSama(Me.txtPrimaryPO.Text, _
                                               Me.txtPrimaryAccount.Text, _
                                               Me.txtPrimaryCurrTax.Text, ClsTransaksiAOP.EnumTipeAop.AOP) <= 0 Then
                    Dim PrimaryAOPDPFirst As Double = Me.ObjAOP.GetPrimaryAOPwithPOSama(Me.txtPrimaryPO.Text, ClsTransaksiAOP.EnumTipeAop.AOP)
                    If PrimaryAOPDPFirst > -1 Then
                        Dim AOPNO As String = Me.ObjAOP.GetNoTransaksi(PrimaryAOPDPFirst)
                        Dim PrimaryAccountB4 As Double = Me.ObjAOP.GetPrimaryAccountAP(PrimaryAOPDPFirst)
                        Dim PrimaryCurrTaxB4 As Double = Me.ObjAOP.GetPrimaryCurrTax(PrimaryAOPDPFirst)
                        Dim PrimaryAccountDPB4 As Double = Me.ObjAOP.GetPrimaryAccountDP(PrimaryAOPDPFirst)
                        If PrimaryAccountDPB4 <> Me.txtPrimaryDPAccount.Text OrElse _
                           PrimaryCurrTaxB4 <> Me.txtPrimaryCurrTax.Text Then
                            MsgBox("There's some of AOP transaction using the similiar PO, with different account payable or different currency tax. " & _
                                   "Please use the same account payable and currency tax for this transaction. AOP No. " & _
                                   AOPNO & ".", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If


                If Me.ObjAOP.GetJmlAOPwithPOSama(Me.txtPrimaryPO.Text, ClsTransaksiAOP.EnumTipeAop.AOP_DP) > 0 Then
                    Dim PrimaryAOPDPFirst As Double = Me.ObjAOP.GetPrimaryAOPwithPOSama(Me.txtPrimaryPO.Text, ClsTransaksiAOP.EnumTipeAop.AOP_DP)
                    Dim AOPNO As String = Me.ObjAOP.GetNoTransaksi(PrimaryAOPDPFirst)
                    Dim PrimaryAccountB4 As Double = Me.ObjAOP.GetPrimaryAccountAP(PrimaryAOPDPFirst)
                    Dim PrimaryCurrTaxB4 As Double = Me.ObjAOP.GetPrimaryCurrTax(PrimaryAOPDPFirst)
                    Dim PrimaryAccountDPB4 As Double = Me.ObjAOP.GetPrimaryAccountDP(PrimaryAOPDPFirst)

                    If PrimaryAccountDPB4 <> Me.txtPrimaryDPAccount.Text OrElse _
                       PrimaryAccountB4 <> Me.txtPrimaryAccount.Text OrElse _
                       PrimaryCurrTaxB4 <> Me.txtPrimaryCurrTax.Text Then
                        MsgBox("There's some of AOP down payment using the similiar PO, " & _
                        "with different account DP or account payable or different currency tax, " & _
                        "please use the same account DP and account payable and currency tax for this transaction. AOP DP No. " & _
                        AOPNO & ".", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                End If
            End If

            'If Me.chkTax.Checked = True Then
            '    If CDec(Me.txtTax.Text) <= 0 Then
            '        MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
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
                    Me.txtTax.Focus()

                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If

                If CDec(Me.txtTax.Text) <> CDec(Dummy) Then
                    Me.txtTax.Focus()

                    If MsgBox("Tax rate <> " + Me.ObjCD.DigitGrouping(Dummy, 2) + "%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If

            If CDec(Me.txtDutyStamp.Text) > 0 Then
                If Me.txtPrimaryCurrency.Text <> "1" Then
                    MsgBox("Duty stamp is for IDR AOP only.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'po ini adalah po yang diciptakan untuk quick purchase
            If Me.txtPrimaryPO.Text = "" Then
                Me.txtPrimaryPO.Text = -1
            End If
            Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAOPQuick(Me.txtPrimaryPO.Text).HasValue = True Then
                Dim ObjAOP As New ClsTransaksiAOP
                Dim PrimaryAOP As Double = DTA.ScalarQuery_GetPrimaryAOPQuick(Me.txtPrimaryPO.Text).Value
                If PrimaryAOP <> -1 Then
                    MsgBox("This PO is refered to quick purchase: " + _
                    ObjAOP.GetNoTransaksi(PrimaryAOP) + ". You cannot use this data.", MsgBoxStyle.Information, "Information")
                    ObjAOP = Nothing
                    Return False
                End If
                ObjAOP = Nothing
            End If

            'Update NPWP 16 Digit
            If chkTax.Checked = True Then 'hanya dicek ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                '2025 Core Tax Tahap 2.1
                If ObjSetNPWP.SetNPWP_New(Me.dtpDate.Value.Date, Me.txtPrimarySupplier.Text, Me.TxtNPWPNo.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Supplier, _
                 frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionPurchase, Me.UserCode, Me.TxtPrimaryAOP.Text) = False Then
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
            With Me.objInvoiceDP

                If Me.Lokal = True Then
                    .Lokal = ClsTransaksiAOP.EnumLokal.Lokal
                Else
                    .Lokal = ClsTransaksiAOP.EnumLokal.Import
                End If
                'invoice DP: DP=0 dan Calculates by invoice
                .PrimaryPO4DP = CDbl(Me.txtPrimaryPO.Text)
                .CalculatesBy = ClsTransaksiAOP.EnumCalculatesBy.AOP
                .DownPayment = 0
                .DutyStamp = CDbl(Me.txtDutyStamp.Text)
                .ExchangRateToIDR = CDbl(Me.txtExhangeRateToIDR.Text)
                .ExhangeRateToUSD = CDbl(Me.txtExchangeRateToUSD.Text)
                'invoice DP tidak ada revisi
                .Total = CDbl(Me.TxtTotal.Text)
                .TaxExchangeRate = CDbl(Me.txtTaxExchangeRate.Text)
                .EFak_KodeJenisTransaksiPajak = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
                .NoTransaksi = Me.TxtAOPNo.Text
                .PrimaryCurrency = CDbl(Me.txtPrimaryCurrency.Text)
                .PrimarySupplier = CDbl(Me.txtPrimarySupplier.Text)
                .SubtotalBeforeDP = CDbl(Me.txtSubTotal.Text)
                .DownPayment = 0
                .SubTotal = CDbl(Me.txtSubTotal.Text)
                .TaxAmount = CDbl(Me.txtTaxAmount.Text)
                .PaymentTerms = CDbl(Me.TxtPaymentsTerms.Text)
                .Tax = CDbl(Me.txtTax.Text)
                .Total = CDbl(Me.TxtTotal.Text)
                If Me.chkTax.Checked = True Then
                    .TaxExclude = ClsTransaksiAOP.EnumTaxExclude.Exclude
                Else
                    .TaxExclude = ClsTransaksiAOP.EnumTaxExclude.Include
                End If
                .Remarks = Me.TxtRemark.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .SupplierInvoiceNo = Me.txtSupplierInvoiceNo.Text
                .SupplierInvoiceDate = Me.dtpSupplierInvoiceDate.Value.Date
                .NoNPWPSupplier = Me.TxtNPWPNo.Text
                .NoFakturPajak = Me.TxtTaxInvoiceNumber.Text
                If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    .VATDipungut = True
                Else
                    .VATDipungut = False
                End If
                'Me.txtPrimaryAccount.Visible = True
                'Me.txtPrimaryCurrency.Visible = True

                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
                .PrimaryAccountDP = Me.txtPrimaryDPAccount.Text

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objInvoiceDP
                Me.TxtPrimaryAOP.Text = .Primary.ToString
                Me.TxtAOPNo.Text = .NoTransaksi
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency.ToString
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))

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
                    Me.txtPrimarydpAccount.Text = .PrimaryAccountDP
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
                Me.txtPrimaryPO.Text = .PrimaryPO4DP.ToString
                Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(CDbl(Me.txtPrimaryPO.Text))
                Me.txtPrimarySupplier.Text = .PrimarySupplier.ToString
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal.ToString, Me.DefaultPrecision))
                If .TaxExclude = ClsTransaksiAOP.EnumTaxExclude.Exclude Then
                    Me.chkTax.Checked = True
                Else
                    Me.chkTax.Checked = False
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(.Tax.ToString, 2)
                Me.TxtPaymentsTerms.Text = .PaymentTerms.ToString
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(.TaxExchangeRate.ToString, 2)
                Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = .EFak_KodeJenisTransaksiPajak

                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(.DutyStamp, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))


                Me.txtExhangeRateToIDR.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExchangRateToIDR.ToString, Me.DefaultPrecision))
                Me.txtExchangeRateToUSD.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExhangeRateToUSD.ToString, Me.DefaultPrecision))
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total.ToString, Me.DefaultPrecision))
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtSupplierInvoiceNo.Text = .SupplierInvoiceNo
                Me.dtpSupplierInvoiceDate.Value = .SupplierInvoiceDate
                Me.TxtNPWPNo.Text = .NoNPWPSupplier
                Me.TxtTaxInvoiceNumber.Text = .NoFakturPajak
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

                Me.IsiCBOAkunDP()
                If .PrimaryAccount > -1 Then
                    Me.txtPrimarydpAccount.Text = .PrimaryAccountDP
                    Me.cboPrimaryDPAccount.Text = .PrimaryAccountDP
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimarydpAccount.Text = -1
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


                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub Hitung_Total()
        Dim exchangRate As Double = 0
        Dim IDR_DefPrec As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Total_MiddlePoint()
                Exit Sub
            End If
            IDR_DefPrec = Me.ObjCurrency.GetDefaultPrecission(1)
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            'hitung pajak + Total
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * CDec(Me.txtTaxExchangeRate.Text) * _
                                               CDec(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down).ToString, IDR_DefPrec)
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                End If
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtDutyStamp.Text), IDR_DefPrec)

                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((CDec(Me.txtSubTotal.Text) + _
                               CDec(Me.txtTaxAmount.Text) + CDec(Me.txtDutyStamp.Text)).ToString, IDR_DefPrec))

            Else

                If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    If Me.chkTax.Checked = True Then
                        If Me.cboPPNDipungut.SelectedIndex = 0 Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * CDec(Me.txtSubTotal.Text) / 100, _
                                                   DefaultPrecision, ClsCekData.EnumTipeRounding.Down).ToString, DefaultPrecision)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefaultPrecision)
                        End If
                    Else
                        Me.txtTax.Text = "0.00"
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefaultPrecision)
                    End If

                    Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtDutyStamp.Text), DefaultPrecision)

                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((CDec(Me.txtSubTotal.Text) + _
                                             CDec(Me.txtTaxAmount.Text) + CDec(Me.txtDutyStamp.Text)), DefaultPrecision))
                Else
                    If Me.chkTax.Checked = True Then
                        If Me.cboPPNDipungut.SelectedIndex = 0 Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * CDec(Me.txtTaxExchangeRate.Text) _
                                                    * CDec(Me.txtSubTotal.Text) / 100, _
                                                   DefaultPrecision, ClsCekData.EnumTipeRounding.Down).ToString, IDR_DefPrec)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                        End If
                    Else
                        Me.txtTax.Text = "0.00"
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                    End If

                    Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtDutyStamp.Text), IDR_DefPrec)
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
                End If
            End If
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------------------------

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub Hitung_Total_MiddlePoint()
        Dim exchangRate As Double = 0
        Dim IDR_DefPrec As Integer = 0

        Try
            IDR_DefPrec = Me.ObjCurrency.GetDefaultPrecission(1)
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            'hitung pajak + Total
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * CDec(Me.txtTaxExchangeRate.Text) * _
                                               CDec(Me.txtSubTotal.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint).ToString, IDR_DefPrec)
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                End If
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtDutyStamp.Text), IDR_DefPrec)

                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((CDec(Me.txtSubTotal.Text) + _
                               CDec(Me.txtTaxAmount.Text) + CDec(Me.txtDutyStamp.Text)).ToString, IDR_DefPrec))

            Else

                If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    If Me.chkTax.Checked = True Then
                        If Me.cboPPNDipungut.SelectedIndex = 0 Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * CDec(Me.txtSubTotal.Text) / 100, _
                                                   DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint).ToString, DefaultPrecision)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefaultPrecision)
                        End If
                    Else
                        Me.txtTax.Text = "0.00"
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefaultPrecision)
                    End If

                    Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtDutyStamp.Text), DefaultPrecision)

                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((CDec(Me.txtSubTotal.Text) + _
                                             CDec(Me.txtTaxAmount.Text) + CDec(Me.txtDutyStamp.Text)), DefaultPrecision))
                Else
                    If Me.chkTax.Checked = True Then
                        If Me.cboPPNDipungut.SelectedIndex = 0 Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * CDec(Me.txtTaxExchangeRate.Text) _
                            * CDec(Me.txtSubTotal.Text) / 100, DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint).ToString, IDR_DefPrec)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                        End If
                    Else
                        Me.txtTax.Text = "0.00"
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", IDR_DefPrec)
                    End If

                    Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtDutyStamp.Text), IDR_DefPrec)
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
                End If
            End If
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            '-------------------------------------------------------------------------------------------------------------------------------------------------------

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================
    '====================================================================================================================================================================

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub dtpDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.GotFocus
        Me.dtpDate.BackColor = Me.BackColorAll
    End Sub

    Private Sub dtpDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.LostFocus
        Me.dtpDate.BackColor = Me.TxtPrimaryAOP.BackColor
    End Sub

    Private Sub txtPONo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONo.Enter
        Try
            Me.buffer = Me.txtPONo.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPONo.ReadOnly = True Then
                    Me.txtPONo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtPONo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONo.LostFocus
        Me.txtPONo.BackColor = Me.TxtPrimaryAOP.BackColor
    End Sub

    Private Sub txtExchangeRateToUSD_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchangeRateToUSD.GotFocus
        Me.txtExchangeRateToUSD.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtExchangeRateToUSD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchangeRateToUSD.LostFocus
        Me.txtExchangeRateToUSD.BackColor = Me.TxtPrimaryAOP.BackColor
    End Sub

    Private Sub txtExhangeRateToIDR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.GotFocus
        Me.txtExhangeRateToIDR.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtExhangeRateToIDR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.LostFocus
        Me.txtExhangeRateToIDR.BackColor = Me.TxtPrimaryAOP.BackColor
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtCurrencyCode.ReadOnly = False Then
                    Me.txtCurrencyCode.BackColor = Me.BackColorAll
                Else
                    Me.txtCurrencyCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtCurrencyCode.ReadOnly = False Then
                    Me.txtCurrencyCode.BackColor = Me.BackColorAll
                Else
                    Me.txtCurrencyCode.BackColor = Me.TxtPrimaryAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubTotal_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubTotal.Enter
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtSubTotal, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubTotal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubTotal.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtSubTotal.ReadOnly = True Then
                    Me.txtSubTotal.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtSubTotal.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.Enter
        Try
            If Me.TxtRemark.Enabled = True And Me.txtDutyStamp.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtDutyStamp, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtDutyStamp.ReadOnly = True Then
                    Me.txtDutyStamp.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtDutyStamp.BackColor = Me.BackColorAll
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
                    Me.TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtRemark.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.ReadOnly = True Then
                    Me.TxtRemark.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtRemark.BackColor = Me.TxtPrimaryAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Me.TxtRemark.BackColor = Me.TxtPrimaryAOP.BackColor
    End Sub

    Private Sub txtSubTotal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSubTotal.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSubTotal, sender, e)
    End Sub

    Private Sub txtSubTotal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSubTotal.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSubTotal, sender, e)
    End Sub

    Private Sub txtSubTotal_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubTotal.Leave
        Try
            If Me.TxtRemark.Enabled = True And Me.txtSubTotal.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.txtSubTotal, sender, e, Me.DefaultPrecision, Me.ObjCD)

                If Me.txtSubTotal.ReadOnly = True Then
                    Me.txtSubTotal.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtSubTotal.BackColor = Me.TxtPrimaryAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubTotal_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubTotal.LostFocus
        Me.txtSubTotal.BackColor = Me.TxtPrimaryAOP.BackColor
        If Me.TxtRemark.Enabled = True Then
            Me.Hitung_Total()
        End If
    End Sub

    Private Sub txtDutyStamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.LostFocus
        If Me.txtDutyStamp.Enabled = True And Me.txtDutyStamp.ReadOnly = False Then
            Me.txtDutyStamp.BackColor = Me.TxtPrimaryAOP.BackColor
            Me.Hitung_Total()
        End If
    End Sub

    Private Sub FrmTransaksiAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            ''Me.CboFakturPajak.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtExchangeRateToUSD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExchangeRateToUSD.KeyDown
        Me.ObjCD.Number_KeyDown(txtExchangeRateToUSD, sender, e)
    End Sub

    Private Sub txtExchangeRateToUSD_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExchangeRateToUSD.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtExchangeRateToUSD, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ij As New FrmCetakTransaksiAOP
        Try
            Dim obj As New ClsTransaksiAOP
            obj.CetakAop(ij.crv1, CDbl(Me.TxtPrimaryAOP.Text))
            obj = Nothing
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.objInvoiceDP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP_DP
            Me.Text = "Approval of Purchase - Down Payment"
            Me.ShowIcon = False
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
            If Me.TxtRemark.Text = "" And CDbl(Me.TxtTotal.Text) = 0 Then
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
                        Me.objInvoiceDP.CariDetail(CDbl(Me.TxtPrimaryAOP.Text))
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
            ij.TipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP_DP
            ij.Text = "Find Purchase Down Payment"

            ij.ShowDialog()
            If Not ij.PrimaryTransaksi = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.objInvoiceDP.CariDetail(ij.PrimaryTransaksi) _
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
        Try
            If Me.TxtAOPNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOPNo.Text <> "" Then
                If Me.objInvoiceDP.ApakahBolehEdit(Me.objInvoiceDP.Primary, Me.UserCode, True) = True Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.baru = False
                    Me.objInvoiceDP.CariDetail(Me.objInvoiceDP.Primary)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
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
            If Me.TxtAOPNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtAOPNo.Text <> "" Then
                If Me.objInvoiceDP.ApakahBolehHapus(Me.objInvoiceDP.Primary, True, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objInvoiceDP.Hapus(CDbl(Me.TxtPrimaryAOP.Text), Me.UserCode, True) = True Then
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
            If Me.TxtRemark.Enabled = True Then
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


    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text).ToString
                    Me.SetDefaultPrecision()
                    Me.txtSubTotal.Focus()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency.ToString
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                        Me.SetDefaultPrecision()
                        Me.txtSubTotal.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
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
                    Me.txtSubTotal.Focus()
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text).ToString
                    Me.SetDefaultPrecision()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency.ToString
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                        Me.SetDefaultPrecision()
                        Me.txtSubTotal.Focus()
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                Me.txtSubTotal.Focus()
                Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text).ToString
                Me.SetDefaultPrecision()
            Else
                i.BackColorAll = Me.BackColorAll
                i.cboCriteria.SelectedIndex = 0
                If Me.txtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.txtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.txtPrimaryCurrency.Text = i.PrimaryCurrency.ToString
                    Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                    Me.SetDefaultPrecision()
                    Me.txtSubTotal.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetDefaultPrecision()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetDefaultPrecision()
        Try
            Me.DefaultPrecision = CInt(Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
            Me.Hitung_Total()
            'Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            'Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxExchangeRate.Text, DefaultPrecision))
            'Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim Dummy As String = ""
        Try
            If Me.chkTax.Checked = True Then
                Me.txtTax.ReadOnly = False
                Me.txtTax.BackColor = Me.TxtPrimaryAOP.BackColor
                Me.txtTax.TabStop = True
                Me.txtTaxExchangeRate.ReadOnly = False
                Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryAOP.BackColor
                Me.txtTaxExchangeRate.TabStop = True

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Me.txtTax.TabStop = False
                Me.txtTaxExchangeRate.ReadOnly = True
                Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                Me.txtTaxExchangeRate.TabStop = False
                Me.txtTax.Text = "0"
                Me.txtTaxExchangeRate.Text = "1"
            End If


            If Me.TxtRemark.Enabled = True Then
                Me.SetTaxRate()
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTax.ReadOnly = False Then
                    Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
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

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            If Me.TxtRemark.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                Me.ObjCD.Number_Leave(Me.txtTax, sender, e, 2, Me.ObjCD)
                Me.Hitung_Total()

                If Me.txtTax.ReadOnly = True Then
                    Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTax.BackColor = Me.TxtPrimaryAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
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
            If Me.TxtRemark.Enabled = True Then
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

    Private Sub txtTax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTax.TextChanged
        
    End Sub

    Private Sub TxtCustomerName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplierName.TextChanged
        Try
            If Me.txtPrimarySupplier.Text <> "-1" And Me.txtPrimarySupplier.Text <> "" Then
                If Me.TxtSupplierName.Text <> "" Then
                    If Me.txtSupplierCode.Text <> "" Then
                        Me.Text = "Approval of Purchase - Down Payment : " + Me.txtSupplierCode.Text + " - " + Me.TxtSupplierName.Text
                    Else
                        Me.Text = "Approval of Purchase - Down Payment"
                    End If
                Else
                    Me.Text = "Approval of Purchase - Down Payment"
                End If
            Else
                Me.Text = "Approval of Purchase - Down Payment"
            End If
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

    Private Sub txtTaxExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Enter
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = False Then
                    Me.ObjCD.Number_Enter(Me.txtTaxExchangeRate, sender, e)
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

    Private Sub txtTaxExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
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
            If Me.TxtRemark.Enabled = True Then
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

    Private Sub txtSalesOrderNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPONo.KeyDown
        Dim i As New FrmCariTransaksiPO
        Dim Jml As Long = 0
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If e.KeyValue = 13 Then
                If Me.ObjPO.ApakahSudahAda(Me.txtPONo.Text) = True Then
                    Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONo.Text).ToString
                    Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(CDbl(Me.txtPrimaryPO.Text))
                    Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text).ToString
                    Dim SupName, CPName, PhoneNo As String
                    SupName = ""
                    CPName = ""
                    PhoneNo = ""

                    ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(CDbl(Me.txtPrimarySupplier.Text), SupName, CPName, PhoneNo)
                    Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                    Me.TxtSupplierName.Text = SupName
                    Me.SetCurrency()
                    Me.txtExchangeRateToUSD.Focus()

                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtPONo.Text <> "" Then
                        i.txtCari.Text = Me.txtPONo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryPO > 0 Then
                        Me.txtPrimaryPO.Text = i.PrimaryPO.ToString  'Me.ObjPO.GetPrimary(Me.txtPONO.Text)
                        Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(CDbl(Me.txtPrimaryPO.Text))
                        Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(CDbl(Me.txtPrimaryPO.Text)).ToString
                        Dim SupName, CPName, PhoneNo As String
                        SupName = ""
                        CPName = ""
                        PhoneNo = ""

                        ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(CDbl(Me.txtPrimarySupplier.Text), SupName, CPName, PhoneNo)
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                        Me.TxtSupplierName.Text = SupName
                        Me.SetCurrency()
                        Me.txtExchangeRateToUSD.Focus()

                    Else
                        Me.txtPONo.Clear()
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtPONo.Focus()
                    End If
                End If

                If Me.txtPONo.Text <> "" Then
                    If Me.ObjPO.ApakahSudahOtorisasi(CDbl(Me.txtPrimaryPO.Text)) = False Then
                        MsgBox("The purchase order doesn't have sufficient authorization, please authorize first", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtPrimaryPO.Text = -1
                        Me.txtPONo.Clear()
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtPONo.Focus()
                    End If
                End If
            End If

            If Me.txtPONo.Text = "" Or Me.txtPrimaryPO.Text = "" Or Me.txtPrimaryPO.Text = "-1" Then
                Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
            Else
                Me.txtPrimaryCurrency.Text = Me.ObjPO.GetPrimaryCurrency(Me.txtPrimaryPO.Text)
            End If
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            Me.SetTaxRate()
            Me.Hitung_Total()

            If Me.baru = True Then
                Me.GetNoBaru()
            End If
            Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWP(Me.txtPrimarySupplier.Text)
            '2024-08-27 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
            If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                '2025 Core Tax Tahap 2.1
                Me.TxtNPWPNo.Text = ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
            Else
                Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWPOLD(Me.txtPrimarySupplier.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSalesOrderNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONo.Leave
        Dim i As New FrmCariTransaksiPO
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPONo.Text <> "" Then
                    If Me.ObjPO.ApakahSudahAda(Me.txtPONo.Text) = True Then
                        Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONo.Text).ToString
                        Me.txtPrimarySupplier.Text = Me.ObjPO.GetPrimarySupplier(CDbl(Me.txtPrimaryPO.Text)).ToString
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                        Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                        Me.SetCurrency()
                        Me.txtExchangeRateToUSD.Focus()
                    Else
                        i.BackColorAll = Me.BackColorAll
                        If Me.txtPONo.Text <> "" Then
                            i.txtCari.Text = Me.txtPONo.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryPO > 0 Then
                            Me.txtPrimaryPO.Text = i.PrimaryPO.ToString
                            Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(CDbl(Me.txtPrimaryPO.Text))
                            Me.txtPrimarySupplier.Text = Me.ObjPO.GetPrimarySupplier(CDbl(Me.txtPrimaryPO.Text)).ToString
                            Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                            Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                            Me.SetCurrency()
                            Me.txtExchangeRateToUSD.Focus()
                        Else
                            Me.txtPONo.Clear()
                            Me.txtPrimaryPO.Text = "-1"
                            Me.txtSupplierCode.Clear()
                            Me.TxtSupplierName.Clear()
                            Me.txtPrimarySupplier.Text = "-1"
                            Me.txtExchangeRateToUSD.Focus()
                        End If
                    End If
                Else
                    Me.txtPONo.Clear()
                    Me.txtPrimaryPO.Text = "-1"
                    Me.txtSupplierCode.Clear()
                    Me.TxtSupplierName.Clear()
                    Me.txtPrimarySupplier.Text = "-1"
                    'Me.txtPONo.Focus()
                End If

                If Me.txtPONo.Text <> "" Then
                    If Me.ObjPO.ApakahSudahOtorisasi(CDbl(Me.txtPrimaryPO.Text)) = False Then
                        MsgBox("The purchase order doesn't have sufficient authorization, please authorize first", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtPONo.Clear()
                        Me.txtPrimaryPO.Text = "-1"
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtPONo.Focus()
                    End If
                End If

                If Me.txtPONo.Text = "" Then
                    Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
                Else
                    Me.txtPrimaryCurrency.Text = Me.ObjPO.GetPrimaryCurrency(Me.txtPrimaryPO.Text)
                End If
                If Me.txtPONo.Text.ToLower <> Me.buffer.ToLower Then
                    Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                    Me.SetTaxRate()
                    Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                    '2024-08-27 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.TxtNPWPNo.Text = ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
                    Else
                        Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWPOLD(Me.txtPrimarySupplier.Text)
                    End If
                End If
                Me.Hitung_Total()

                If Me.baru = True Then
                    Me.GetNoBaru()
                End If

                If Me.txtPONo.ReadOnly = True Then
                    Me.txtPONo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtPONo.BackColor = Me.TxtPrimaryAOP.BackColor
                End If

                If Me.txtPONo.Text = "" Then
                    Me.TxtPaymentsTerms.Text = "0"
                Else
                    If CDbl(Me.TxtPaymentsTerms.Text) = 0 Then
                        Me.TxtPaymentsTerms.Text = Me.ObjPO.GetPaymentDue(Me.txtPrimaryPO.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetCurrency()
        Try
            If Me.txtPrimaryPO.Text <> "-1" Then
                Me.txtPrimaryCurrency.Text = Me.ObjPO.GetPrimaryCurrency(CDbl(Me.txtPrimaryPO.Text))
                Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))


                Me.SetDefaultPrecision()
                If Me.ObjPO.GetUseTax(CDbl(Me.txtPrimaryPO.Text)) = True Then
                    Me.chkTax.Checked = True
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetTaxPercentage(CDbl(Me.txtPrimaryPO.Text)).ToString, 2)
                    Me.Hitung_Total()
                Else
                    Me.chkTax.Checked = False
                    Me.txtTax.Text = "0"
                    Me.Hitung_Total()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSalesOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPO.Click
        Dim i As New FrmCariTransaksiPO
        'Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        'Dim tbl As New DseTransaksiAOP.tbltrans_aop2DataTable
        'Dim row As DseTransaksiAOP.tbltrans_aop2Row
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            i.BackColorAll = Me.BackColorAll

            If Me.txtPONo.Text <> "" Then
                i.txtCari.Text = Me.txtPONo.Text
            End If

            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                Me.txtPrimaryPO.Text = i.PrimaryPO.ToString
                Me.txtPONo.Text = Me.ObjPO.GetNoTransaksiPO(CDbl(Me.txtPrimaryPO.Text))
                Me.txtPrimarySupplier.Text = Me.ObjPO.GetPrimarySupplier(CDbl(Me.txtPrimaryPO.Text)).ToString
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.txtPrimarySupplier.Text))
                Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.txtPrimarySupplier.Text))
                Me.SetCurrency()
                Me.txtExchangeRateToUSD.Focus()
            Else
                Me.txtPONo.Clear()
                Me.txtPrimaryPO.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtPONo.Focus()
            End If
            Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWP(Me.txtPrimarySupplier.Text)
            '2024-08-27 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
            If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWP(Me.txtPrimarySupplier.Text)
                '2025 Core Tax Tahap 2.1
                Me.TxtNPWPNo.Text = ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
            Else
                Me.TxtNPWPNo.Text = Me.ObjSupplier.GetNPWPOLD(Me.txtPrimarySupplier.Text)
            End If

            If Me.txtPONo.Text = "" Then
                Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
            Else
                Me.txtPrimaryCurrency.Text = Me.ObjPO.GetPrimaryCurrency(Me.txtPrimaryPO.Text)
            End If
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)

            Me.SetTaxRate()
            Me.Hitung_Total()

            If Me.baru = True Then
                Me.GetNoBaru()
            End If

            If Me.txtPONo.Text <> "" Then
                If Me.ObjPO.ApakahSudahOtorisasi(CDbl(Me.txtPrimaryPO.Text)) = False Then
                    MsgBox("The purchase order doesn't have sufficient authorization, please authorize first", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtPrimaryPO.Text = "-1"
                    Me.txtPONo.Clear()
                    Me.txtPrimaryPO.Text = "-1"
                    Me.txtSupplierCode.Clear()
                    Me.TxtSupplierName.Clear()
                    Me.txtPrimarySupplier.Text = "-1"
                    Me.txtPONo.Focus()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetTaxRate(Optional ByVal vSetNewTaxRate As Boolean = True)
        Dim Rate As Decimal = 0

        Try
            If Me.chkTax.Checked = True Then
                If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                    Me.txtTaxExchangeRate.Enabled = False
                    Me.txtTaxExchangeRate.ReadOnly = True
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Me.txtTaxExchangeRate.TabStop = False
                    If vSetNewTaxRate = True Then
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Me.ObjCurrency.GetDefaultPrecission(1))
                    End If
                Else
                    Me.txtTaxExchangeRate.Enabled = True
                    Me.txtTaxExchangeRate.ReadOnly = False
                    Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryAOP.BackColor
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
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Me.txtTaxExchangeRate.TabStop = False
                    If vSetNewTaxRate = True Then
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Me.ObjCurrency.GetDefaultPrecission(1))
                    End If
                Else
                    Me.txtTaxExchangeRate.Enabled = True
                    Me.txtTaxExchangeRate.ReadOnly = False
                    Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryAOP.BackColor
                    Me.txtTaxExchangeRate.TabStop = True
                    If vSetNewTaxRate = True Then
                        Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                        If Rate = 0 Then
                            Rate = 1
                        End If
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Me.ObjTaxRate.DefaultPrec)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCurrency(ByVal vPrimaryPO As Double)
        Dim PrimaryCur As Double = -1

        Try
            PrimaryCur = Me.ObjPO.GetPrimaryCurrency(vPrimaryPO)
            If PrimaryCur <> -1 Then
                Me.txtPrimaryCurrency.Text = PrimaryCur.ToString
                Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                Me.SetDefaultPrecision()
            Else
                Me.txtPrimaryCurrency.Text = -1
                Me.txtCurrencyCode.Text = ""
                Me.SetDefaultPrecision()
            End If
            If Me.ObjPO.GetUseTax(vPrimaryPO) = True Then
                Me.chkTax.Checked = True
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetTaxPercentage(vPrimaryPO).ToString)
                Me.Hitung_Total()
            Else
                Me.chkTax.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentsTerms.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentsTerms.ReadOnly = True Then
                    Me.TxtPaymentsTerms.BackColor = Me.TxtReadOnly.BackColor
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

    Private Sub TxtPaymentsTerms_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentsTerms.Leave
        Try
            If Me.TxtRemark.Enabled = True And Me.TxtPaymentsTerms.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.TxtPaymentsTerms, sender, e, 0, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentsTerms.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentsTerms.ReadOnly = True Then
                    Me.TxtPaymentsTerms.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPaymentsTerms.BackColor = Me.TxtPrimaryAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetNoBaru()
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjPO.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjPO.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtAOPNo.Text = Me.objInvoiceDP.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtAOPNo.Text = Me.objInvoiceDP.GetNoBaru(Me.dtpDate.Value, CDbl(Me.txtPrimarySupplier.Text))
                    End If
                Else
                    Me.TxtAOPNo.Text = Me.objInvoiceDP.GetNoBaru(Me.dtpDate.Value, CDbl(Me.txtPrimarySupplier.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            Me.GetNoBaru()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Leave
        Dim defPrec As Decimal = 0
        Try
            If Me.TxtRemark.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                If CDec(Me.txtTaxExchangeRate.Text) = 0 Then
                    Me.txtTaxExchangeRate.Text = 1
                End If
                defPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Me.ObjCD.Number_Leave(Me.txtTaxExchangeRate, sender, e, defPrec, Me.ObjCD)
                Me.Hitung_Total()

                If Me.txtTaxExchangeRate.ReadOnly = True Then
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryAOP.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.Leave
        Try
            If Me.TxtRemark.Enabled = True And Me.txtDutyStamp.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.txtDutyStamp, sender, e, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text), Me.ObjCD)
                Me.Hitung_Total()
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
                        Me.TxtNPWPNo.BackColor = Me.TxtReadOnly.BackColor
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
                        Me.TxtNPWPNo.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.TxtNPWPNo.BackColor = Me.TxtPrimaryAOP.BackColor
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
                        Me.TxtTaxInvoiceNumber.BackColor = Me.TxtReadOnly.BackColor
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
                        Me.TxtTaxInvoiceNumber.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.TxtTaxInvoiceNumber.BackColor = Me.TxtPrimaryAOP.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierInvoiceNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierInvoiceNo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtSupplierInvoiceNo.Enabled = True Then
                    If Me.txtSupplierInvoiceNo.ReadOnly = True Then
                        Me.txtSupplierInvoiceNo.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtSupplierInvoiceNo.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierInvoiceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierInvoiceNo.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtSupplierInvoiceNo.Enabled = True Then
                    If Me.txtSupplierInvoiceNo.ReadOnly = True Then
                        Me.txtSupplierInvoiceNo.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtSupplierInvoiceNo.BackColor = Me.TxtPrimaryAOP.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPPNDipungut_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPPNDipungut.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.Hitung_Total()
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
                If Me.objInvoiceDP.CariDetail(vPrimary) _
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

    Public Sub FromCutOff_AOPDP_Create(ByVal vPrimaryTransaksi As Double, _
     ByVal vTblPrimaryPO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
     ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core)

        Try
            Me.baru = False
            Me.objInvoiceDP.ResetObject()
            Me.objInvoiceDP.tipeAOP = ClsTransaksiAOP.EnumTipeAop.AOP_DP
            Me.objInvoiceDP.CariDetail(vPrimaryTransaksi)
            Me.objInvoiceDP.PrimaryPO4DP = ObjCutOff_Core.GetPrimaryUpdate(Me.objInvoiceDP.PrimaryPO4DP, vTblPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_AOPDP_Save(ByVal vTransJurnal As Boolean) As Double
        Try
            Me.baru = True
            Me.objInvoiceDP.SimpanBaru(Me.UserCode, vTransJurnal, False, True)
            Return Me.objInvoiceDP.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
End Class