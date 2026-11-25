Option Explicit On
Option Strict Off

Imports AccModule
Public Class FrmSettingAP_AR
    Public BackColorAll As Color = Color.LightSeaGreen
    Public UserCode As String = ""

    Dim Buffer As String = ""
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjClosing As New Purchase_Module.ClsTransaksiMonthlyClosing
    Dim ObjCurrency As New Core_Module.clsMasterCurrency

    Dim SettingInvoiceBerubah As Boolean = False
    Dim SettingAOPBerubah As Boolean = False
    Dim SettingDOBerubah As Boolean = False
    Dim SettingPOBerubah As Boolean = False
    Dim SettingPRBerubah As Boolean = False
    Dim SettingHRBerubah As Boolean = False
    Dim SettingSOBerubah As Boolean = False
    Dim SettingInventoryBerubah As Boolean = False

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiSettingJurnal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.chkDO_AROverDue.Checked = True
            Me.chkDO_AROverDue.Checked = False
            Me.chkDO_ARExceedCreditLimit.Checked = True
            Me.chkDO_ARExceedCreditLimit.Checked = False

            Me.TxtAkunAR.Focus()
            Me.SalinKeFieldSettingJurnal()
            Me.SalinKeFieldSettingInvoice()
            Me.SalinKeFieldSettingAOP()
            Me.SalinKeFieldSettingDO()
            Me.SalinKeFieldSettingPO()
            Me.SalinKeFieldSettingPR()
            Me.SalinKeFieldSettingReport()
            Me.SalinKeFieldSettingHR()
            Me.SalinKeFieldSettingSO()
            Me.SalinKeFieldSettingInventory()
            Me.SalinKeFieldSettingPph21()
            Me.SalinKeFieldSettingLoanPayroll()

            Me.Text = "Global Settings"
            If Me.cboCosting.SelectedIndex = -1 Then
                Me.cboCosting.SelectedIndex = 0
            End If
            If Me.ObjClosing.GetJml > 0 Then
                Me.cboCosting.Enabled = False
            Else
                Me.cboCosting.Enabled = True
            End If

            'untuk sementara setting HR hidden
            Me.TabControl1.TabPages.Remove(Me.TabPageAttendance)
            'Me.TabControl1.TabPages.Remove(Me.TabPageSalary)
            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        'Dim fileExists As Boolean
        Dim DummyTipe As Integer = 0

        Try
            'jika menggunakan custom invoice print maka alamatnya tidak boleh kosong dan filenya harus ada
            If Me.chkUseCustomInvoice.Checked = True Then
                'pastikan ada isinya
                'Me.TxtCustomInvoiceAddress.Text.Trim()
                'If Me.TxtCustomInvoiceAddress.Text = "" Then
                'MsgBox("Please specify valid custom invoice file.", MsgBoxStyle.Information, "Information")
                'Return False
                'End If
                'pastikan file tsb valid
                'fileExists = My.Computer.FileSystem.FileExists(Me.TxtCustomInvoiceAddress.Text)
                'If fileExists = False Then
                'MsgBox("Please specify valid custom invoice file.", MsgBoxStyle.Information, "Information")
                'Return False
                'End If
            End If

            If Me.chkUseCustomTaxInvoice.Checked = True Then
                'pastikan ada isinya
                'Me.TxtCustomTaxInvoiceAddress.Text.Trim()
                'If Me.TxtCustomTaxInvoiceAddress.Text = "" Then
                'MsgBox("Please specify valid custom tax invoice file.", MsgBoxStyle.Information, "Information")
                'Return False
                'End If
                'pastikan file tsb valid
                'fileExists = My.Computer.FileSystem.FileExists(Me.TxtCustomTaxInvoiceAddress.Text)
                'If fileExists = False Then
                'MsgBox("Please specify valid custom tax invoice file.", MsgBoxStyle.Information, "Information")
                'Return False
                'End If
            End If

            If Me.chkUseCustomDO.Checked = True Then
                'pastikan ada isinya
                'Me.TxtCustomeDOAddress.Text.Trim()
                'If Me.TxtCustomeDOAddress.Text = "" Then
                'MsgBox("Please specify valid custom delivery order file.", MsgBoxStyle.Information, "Information")
                'Return False
                'End If
                'pastikan file tsb valid
                'fileExists = My.Computer.FileSystem.FileExists(Me.TxtCustomeDOAddress.Text)
                'If fileExists = False Then
                'MsgBox("Please specify valid custom delivery order file.", MsgBoxStyle.Information, "Information")
                'Return False
                'End If
            End If

            'current year earning harus diisikan + tipenya harus modal
            If Me.TxtCurrentYearEarningAcc.Text = "" Then
                MsgBox("Please specify current year earning account.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                Me.TxtCurrentYearEarningAcc.Focus()
                Return False
            Else
                If Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryCurrencyYearEarningAcc.Text) <> ClsMasterAkun.EnumTipeAkun.Equity Then
                    MsgBox("Please specify equity type ccount for current year earning account", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                    Me.TxtCurrentYearEarningAcc.Focus()
                    Return False
                End If
            End If

            'retained earning harus diisikan + tipenya harus modal
            If Me.TxtRetainedEarningAcc.Text = "" Then
                MsgBox("Please specify retained earning account.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                Me.TxtRetainedEarningAcc.Focus()
                Return False
            Else
                If Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryRetainedEarningAcc.Text) <> ClsMasterAkun.EnumTipeAkun.Equity Then
                    MsgBox("Please specify equity type ccount for retained earning account", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                    Me.TxtRetainedEarningAcc.Focus()
                    Return False
                End If
            End If

            'payment difference
            If Me.TxtPaymentDiffAcc.Text = "" Then
                MsgBox("Please specify payment difference account.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                Me.TxtPaymentDiffAcc.Focus()
                Return False
            End If

            'VAT out, VAT in, Dutystamp expense bisa tidak diisikan maka dikonfirmasi
            If Me.TxtVATIn.Text = "" Then
                If MsgBox("You don't specify VAT In account, continue save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                    Me.TxtVATIn.Focus()
                    Return False
                End If
            End If
            If Me.TxtVATOutAcc.Text = "" Then
                If MsgBox("You don't specify VAT out account, continue save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                    Me.TxtVATOutAcc.Focus()
                    Return False
                End If
            End If
            If Me.TxtDutystampAcc.Text = "" Then
                If MsgBox("You don't specify dutystamp expense account, continue save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                    Me.TxtDutystampAcc.Focus()
                    Return False
                End If
            End If
            If Me.TxtShippingIncomeAcc.Text = "" Then
                If MsgBox("You don't specify shipping income account, continue save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                    Me.TxtShippingIncomeAcc.Focus()
                    Return False
                End If
            End If
            If Me.TxtPurchaseDPAcc.Text = "" Then
                MsgBox("Please specify purchase down payment account.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                Me.TxtPurchaseDPAcc.Focus()
                Return False
            End If
            'pastikan purchase downpayment account currency-nya IDR
            If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryPurchaseDPAcc.Text) <> Core_Module.PrimaryBaseCurrency Then
                MsgBox("Please specify account with " + Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                " currency for purchase down payment account.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'pastikan purchase downpayment account tipenya asset
            If Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryPurchaseDPAcc.Text) <> ClsMasterAkun.EnumTipeAkun.Asset Then
                MsgBox("Please specify asset account type for purchase down payment account.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.TxtSalesDPAcc.Text = "" Then
                MsgBox("Please specify sales down payment account.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedTab = Me.TabPage_Accounting
                Me.TxtSalesDPAcc.Focus()
                Return False
            End If
            'pastikan sales downpayment account currency-nya IDR
            If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimarySalesDPAcc.Text) <> Core_Module.PrimaryBaseCurrency Then
                MsgBox("Please specify account with " + Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                " currency for sales down payment account.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'pastikan sales downpayment account tipenya asset
            If Me.ObjAkun.GetTipeAkun(Me.TxtPrimarySalesDPAcc.Text) <> ClsMasterAkun.EnumTipeAkun.Liabilities Then
                MsgBox("Please specify liability account type for sales down payment account.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'untuk salary payment account, pastikan tipenya cash n bank
            If Me.TxtPrimaryCashAccount.Text <> "-1" And Me.TxtPrimaryCashAccount.Text <> "" Then
                If Me.ObjAkun.GetTipeAkun2ByPrimaryAkun(Me.TxtPrimaryCashAccount.Text) <> ClsMasterAkun.EnumTipeAkun2.Cash_And_Bank Then
                    MsgBox("Please specify cash and bank account type for salary default settings: cash account.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPageSalary
                    Me.TxtCashAkunCode.Focus()
                    Return False
                End If
            End If

            'untuk tax payable account, pastikan tipe-nya payable dan mata uangnya IDR
            If Me.TxtPrimaryTaxPayableAccount.Text <> "-1" And Me.TxtPrimaryTaxPayableAccount.Text <> "" Then
                If Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxPayableAccount.Text) <> ClsMasterAkun.EnumTipeAkun.Liabilities Then
                    MsgBox("Please specify liabilities account type for salary default settings: tax payable account.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPageSalary
                    Me.TxtTaxPayableAccount.Focus()
                    Return False
                End If
                If Me.ObjAkun.GetPrimaryCurrency(Me.TxtPrimaryTaxPayableAccount.Text) <> Core_Module.PrimaryBaseCurrency Then
                    MsgBox("Please specify account with " + Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + _
                    " currency for tax payable account.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPageSalary
                    Me.TxtTaxPayableAccount.Focus()
                    Return False
                End If
            End If

            'untuk tax expense account, pastikan tipe-nya expense/asset/other expense
            If Me.TxtPrimaryTaxAkun.Text <> "-1" And Me.TxtPrimaryTaxAkun.Text <> "" Then
                DummyTipe = Me.ObjAkun.GetTipeAkun(Me.TxtPrimaryTaxAkun.Text)
                If DummyTipe <> ClsMasterAkun.EnumTipeAkun.Expense And _
                DummyTipe <> ClsMasterAkun.EnumTipeAkun.Other_Expense And _
                DummyTipe <> ClsMasterAkun.EnumTipeAkun.Asset Then
                    MsgBox("Please specify expense or other expense account type for salary default settings: tax expense account.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPageSalary
                    Me.TxtTaxAkunCode.Focus()
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'simpan perubahan ke dalam basis data
            If Me.ApakahDataBenar() = True Then
                'konfirmasi disimpan atau tidak
                hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    'simpan data setting invoice
                    Me.SimpanDataSettingInvoice()
                    'simpan data setting aop
                    Me.SimpanDataSettingAOP()
                    'simpan data jurnal
                    Me.SimpanDataSettingJurnal()
                    'simpan data do
                    Me.SimpanDataSettingDO()
                    'simpan data setting po
                    Me.SimpanDataSettingPO()
                    'simpan data setting pr
                    Me.SimpanDataSettingPR()
                    'simpan data report
                    Me.SimpanDataSettingReport()
                    'simpan data HR
                    Me.SimpanDataSettingHR()
                    'simpan data so
                    Me.SimpanDataSettingSO()
                    'simpan data inventory
                    Me.SimpanDataSettingInventory()
                    'simpan data setting pajak + payroll
                    SimpanDataSettingPPH21()
                    Me.SimpanDataSettingLoanPayroll()

                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account AR
    Private Sub TxtAkunAR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtAkunAR.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAkunAR.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda(Me.TxtAkunAR.Text) = True Then
                        Me.TxtPrimaryAkunAR.Text = CStr(ObjAkun.GetPrimary(Me.TxtAkunAR.Text))
                        Me.TxtAkunAR.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                        Me.TxtAkunAP.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAR.PerformClick()
                    End If
                Else
                    Me.btnFindAR.PerformClick()
                End If
            End If
            Me.TxtARDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAR.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAR_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtAkunAR.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAR.Text <> "" Then
                If ObjAkun.ApakahSudahAda(Me.TxtAkunAR.Text) = True Then
                    Me.TxtPrimaryAkunAR.Text = CStr(ObjAkun.GetPrimary(Me.TxtAkunAR.Text))
                    Me.TxtAkunAR.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                Else
                    If Me.TxtAkunAR.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAkunAR.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAkunAR.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAkunAR.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                    Else
                        Me.TxtAkunAR.Clear()
                        Me.TxtPrimaryAkunAR.Text = "-1"
                        Me.TxtAkunAR.Focus()
                    End If
                End If
            Else
                Me.TxtAkunAR.Clear()
                Me.TxtPrimaryAkunAR.Text = "-1"
            End If
            Me.TxtARDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAR.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAR.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAR.Text <> "" Then
                ij.txtCari.Text = Me.TxtAkunAR.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkunAR.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAkunAR.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
                Me.TxtARDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAR.Text)
                Me.TxtAkunAP.Focus()
            Else
                Me.TxtAkunAR.Clear()
                Me.TxtPrimaryAkunAR.Text = "-1"
                Me.TxtARDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAR.Text)
                Me.TxtAkunAR.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAkunAP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
      Handles TxtAkunAP.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAkunAP.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda(Me.TxtAkunAP.Text) = True Then
                        Me.TxtPrimaryAkunAP.Text = CStr(ObjAkun.GetPrimary(Me.TxtAkunAP.Text))
                        Me.TxtAkunAP.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                        Me.TxtAkunAP.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindAkunAP.PerformClick()
                    End If
                Else
                    Me.btnFindAkunAP.PerformClick()
                End If
            End If
            Me.TxtAPDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAP.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAP_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtAkunAP.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAP.Text <> "" Then
                ''If Me.cboTipeTujuan.SelectedIndex = 0 Then
                If ObjAkun.ApakahSudahAda(Me.TxtAkunAP.Text) = True Then
                    Me.TxtPrimaryAkunAP.Text = CStr(ObjAkun.GetPrimary(Me.TxtAkunAP.Text))
                    Me.TxtAkunAP.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                Else
                    If Me.TxtAkunAP.Text <> "" Then
                        ij.txtCari.Text = Me.TxtAkunAP.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAkunAP.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtAkunAP.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                    Else
                        Me.TxtAkunAP.Clear()
                        Me.TxtPrimaryAkunAP.Text = "-1"
                        Me.TxtAkunAP.Focus()
                    End If
                End If
                ''End If
            Else
                Me.TxtAkunAP.Clear()
                Me.TxtPrimaryAkunAP.Text = "-1"
            End If
            Me.TxtAPDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAP.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindAkunAP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindAkunAP.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtAkunAP.Text <> "" Then
                ij.txtCari.Text = Me.TxtAkunAP.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAkunAP.Text = CStr(ij.mPrimaryAkun)
                Me.TxtAkunAP.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
                Me.btnOK.Focus()
            Else
                Me.TxtAkunAP.Clear()
                Me.TxtPrimaryAkunAP.Text = "-1"
                Me.TxtAkunAP.Focus()
            End If
            Me.TxtAPDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAP.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtAkunAR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAR.GotFocus
        Try
            Me.TxtAkunAR.BackColor = Core_Module.GetSMESSColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAR.LostFocus
        Try
            Me.TxtAkunAR.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAP.GotFocus
        Try
            Me.TxtAkunAP.BackColor = Core_Module.GetSMESSColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAkunAP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAkunAP.LostFocus
        Try
            Me.TxtAkunAP.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingInvoice()
        Dim Dummy As String = ""

        Try
            'salin data use tax
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.chkTaxInvoice.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkTaxInvoice.Checked = False
                Else
                    Me.chkTaxInvoice.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy <> "" Then
                        Me.TxtTaxInvoice.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                    Else
                        Me.TxtTaxInvoice.Text = Me.ObjCD.DigitGrouping(0, 2)
                    End If
                End If
            End If

            'salin data custom invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool)
            If Dummy = "" Then
                Me.chkUseCustomInvoice.Checked = True
                Me.chkUseCustomInvoice.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUseCustomInvoice.Checked = True
                    Me.chkUseCustomInvoice.Checked = False
                Else
                    Me.chkUseCustomInvoice.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomInvoiceAddress_Str)
                    Me.TxtCustomInvoiceAddress.Text = Dummy
                End If
            End If
            'salin data custom tax invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool)
            If Dummy = "" Then
                Me.chkUseCustomTaxInvoice.Checked = True
                Me.chkUseCustomTaxInvoice.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUseCustomTaxInvoice.Checked = True
                    Me.chkUseCustomTaxInvoice.Checked = False
                Else
                    Me.chkUseCustomTaxInvoice.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomTaxInvoice_Str)
                    Me.TxtCustomTaxInvoiceAddress.Text = Dummy
                End If
            End If

            Me.TxtSignedByName.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByName)
            Me.TxtSignedByPosition.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByJabatan)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShipmentTax)
            If Dummy = "" Then
                Me.chkShipmentTax.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkShipmentTax.Checked = False
                Else
                    Me.chkShipmentTax.Checked = True
                End If
            End If

            Me.TxtPrimarySalesGlobalDiscAcc.Text = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.GlobalDiscAcc)
            If Me.TxtPrimarySalesGlobalDiscAcc.Text = "" Then
                Me.TxtPrimarySalesGlobalDiscAcc.Text = "-1"
            End If
            Me.TxtSalesGlobalDiscAcc.Text = Me.ObjAkun.GetAccountCode(Me.TxtPrimarySalesGlobalDiscAcc.Text)
            Me.TxtSalesGlobalDiscAccDesc.Text = Me.ObjAkun.GetDescription(Me.TxtSalesGlobalDiscAcc.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.InvoiceIncludeCalculation)
            If Dummy = "" Then
                Me.cboIncludeCalculation.SelectedIndex = 0
            Else
                If Dummy = "0" Then
                    Me.cboIncludeCalculation.SelectedIndex = 0
                Else
                    Me.cboIncludeCalculation.SelectedIndex = 1
                End If
            End If

            Me.SettingInvoiceBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingHR()
        Dim Dummy As String = ""

        Try
            'salin data ke field
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreBreakData)
            If Dummy = "" Then
                Me.chkIgnoreBreak.Checked = True
            Else
                If Dummy = "0" Then
                    Me.chkIgnoreBreak.Checked = False
                Else
                    Me.chkIgnoreBreak.Checked = True
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                       Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyIn)
            If Dummy = "" Then
                Me.chkIgnoreEarlyIn.Checked = True
            Else
                If Dummy = "0" Then
                    Me.chkIgnoreEarlyIn.Checked = False
                Else
                    Me.chkIgnoreEarlyIn.Checked = True
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyOut)
            If Dummy = "" Then
                Me.chkIgnoreEarlyOut.Checked = True
            Else
                If Dummy = "0" Then
                    Me.chkIgnoreEarlyOut.Checked = False
                Else
                    Me.chkIgnoreEarlyOut.Checked = True
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.OT_Duration_RoundingUnit)
            If Dummy = "" Then
                Me.TxtRoundingUnit.Text = "00.01"
            Else
                If IsNumeric(Dummy) = True Then
                    Me.TxtRoundingUnit.Text = Me.ObjCD.Convert_ToString(Dummy)
                Else
                    Me.TxtRoundingUnit.Text = "00.01"
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.OT_Duration_RoundingType)
            If Dummy = "" Then
                Me.cboRoundingType.SelectedIndex = 0
            Else
                Me.cboRoundingType.SelectedIndex = Dummy
            End If

            Me.SettingHRBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingInventory()
        Dim Dummy As String = ""

        Try
            'salin data custom invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan)
            If Dummy = "" Then
                Me.TxtMaterialPlanning_WorkDays.Text = "0"
            Else
                Me.TxtMaterialPlanning_WorkDays.Text = Dummy
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Purchase, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockInPurchase.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockInPurchase.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_SalesReturn, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockInSalesReturn.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockInSalesReturn.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Adjust, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockInAdjust.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockInAdjust.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Adjust, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockOutAdjust.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockOutAdjust.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Purchase_Return, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockOutPurchaseReturn.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockOutPurchaseReturn.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Sales, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockOutSales.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockOutSales.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Transfer, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockTransfer.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockTransfer.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "1" Then
                Me.chkInv_UseCustomPrint_StockBundling.Checked = True
            Else
                Me.chkInv_UseCustomPrint_StockBundling.Checked = False
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
                      Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_CostFaktorCalc)
            If Dummy = "1" Then
                Me.chkUseCustom_CostFactorCalc.Checked = True
            Else
                Me.chkUseCustom_CostFactorCalc.Checked = False
            End If

            Me.SettingInventoryBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingSO()
        Dim Dummy As String = ""

        Try
            'salin data custom invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit)
            If Dummy = "" Then
                Me.ChkSO_AR_ExceedLimit.Checked = False
            Else
                If Dummy = "0" Then
                    Me.ChkSO_AR_ExceedLimit.Checked = False
                Else
                    Me.ChkSO_AR_ExceedLimit.Checked = True
                End If
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue)
            If Dummy = "" Then
                Me.ChkSO_AR_Overdue.Checked = False
            Else
                If Dummy = "0" Then
                    Me.ChkSO_AR_Overdue.Checked = False
                Else
                    Me.ChkSO_AR_Overdue.Checked = True
                End If
            End If

            Me.SettingSOBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingDO()
        Dim Dummy As String = ""

        Try
            'salin data custom invoice print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool)
            If Dummy = "" Then
                Me.chkUseCustomDO.Checked = True
                Me.chkUseCustomDO.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUseCustomDO.Checked = True
                    Me.chkUseCustomDO.Checked = False
                Else
                    Me.chkUseCustomDO.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomDO_Str)
                    Me.TxtCustomeDOAddress.Text = Dummy
                End If
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit_Tolerance)
            If Dummy = "" Then
                Me.TxtDO_ToleranceARCreditLimit.Text = Me.ObjCD.DigitGrouping(0, Me.ObjCurrency.GetDecimal_DefaultCurrency)
            Else
                Me.TxtDO_ToleranceARCreditLimit.Text = Me.ObjCD.DigitGrouping(Dummy, Me.ObjCurrency.GetDecimal_DefaultCurrency)
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit)
            If Dummy = "" Then
                Me.chkDO_ARExceedCreditLimit.Checked = False
            Else
                If Dummy = 0 Then
                    Me.chkDO_ARExceedCreditLimit.Checked = False
                Else
                    Me.chkDO_ARExceedCreditLimit.Checked = True
                End If
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue_Tolerance)
            If Dummy = "" Then
                Me.TxtDO_ToleranceAROverDue.Text = Me.ObjCD.DigitGrouping(0, Me.ObjCurrency.GetDecimal_DefaultCurrency)
            Else
                Me.TxtDO_ToleranceAROverDue.Text = Me.ObjCD.DigitGrouping(Dummy, 0)
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue)
            If Dummy = "" Then
                Me.chkDO_AROverDue.Checked = False
            Else
                If Dummy = 0 Then
                    Me.chkDO_AROverDue.Checked = False
                Else
                    Me.chkDO_AROverDue.Checked = True
                End If
            End If

            Me.SettingDOBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingInvoice()
        Try
            If Me.SettingInvoiceBerubah = True Then
                If Me.chkTaxInvoice.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, Me.TxtTaxInvoice.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, "0")
                End If

                If Me.chkUseCustomInvoice.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomInvoiceAddress_Str, Me.TxtCustomInvoiceAddress.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomInvoiceAddress_Str, "")
                End If

                If Me.chkUseCustomTaxInvoice.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomTaxInvoice_Str, Me.TxtCustomTaxInvoiceAddress.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomTaxInvoice_Str, "")
                End If

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByJabatan, Me.TxtSignedByPosition.Text)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.SignedByName, Me.TxtSignedByName.Text)

                If Me.chkShipmentTax.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShipmentTax, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShipmentTax, "0")
                End If

                If Me.TxtSalesGlobalDiscAcc.Text = "" Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.GlobalDiscAcc, "-1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.GlobalDiscAcc, Me.TxtPrimarySalesGlobalDiscAcc.Text)
                End If

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.InvoiceIncludeCalculation, Me.cboIncludeCalculation.SelectedIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingInventory()
        Try
            If Me.SettingInventoryBerubah = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan, Me.TxtMaterialPlanning_WorkDays.Text)

                If Me.chkInv_UseCustomPrint_StockBundling.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkInv_UseCustomPrint_StockInAdjust.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Adjust, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Adjust, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkInv_UseCustomPrint_StockInPurchase.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Purchase, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Purchase, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkInv_UseCustomPrint_StockInSalesReturn.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_SalesReturn, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_SalesReturn, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkInv_UseCustomPrint_StockOutAdjust.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Adjust, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Adjust, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkInv_UseCustomPrint_StockOutPurchaseReturn.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Purchase_Return, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Purchase_Return, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkInv_UseCustomPrint_StockOutSales.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Sales, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Sales, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkInv_UseCustomPrint_StockTransfer.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Transfer, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Transfer, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, 0)
                End If
                If Me.chkUseCustom_CostFactorCalc.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_CostFaktorCalc, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_CostFaktorCalc, 0)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SimpanDataSettingHR()
        Try
            If Me.SettingHRBerubah = True Then
                If Me.chkIgnoreBreak.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreBreakData, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreBreakData, "0")
                End If

                If Me.chkIgnoreEarlyIn.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyIn, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyIn, "0")
                End If

                If Me.chkIgnoreEarlyOut.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyOut, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreEarlyOut, "0")
                End If

                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.OT_Duration_RoundingType, Me.cboRoundingType.SelectedIndex)
                If Me.ObjCD.Is_Time_String(Me.TxtRoundingUnit.Text) = False Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.OT_Duration_RoundingUnit, 1)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.OT_Duration_RoundingUnit, Me.ObjCD.Convert_ToMinutes(Me.TxtRoundingUnit.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingSO()
        Try
            If Me.SettingSOBerubah = True Then
                If Me.ChkSO_AR_ExceedLimit.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit, "0")
                End If
                If Me.ChkSO_AR_Overdue.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue, "0")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingDO()
        Try
            If Me.SettingDOBerubah = True Then
                If Me.chkUseCustomDO.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomDO_Str, Me.TxtCustomeDOAddress.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomDO_Bool, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomDO_Str, "")
                End If
                If Me.chkDO_ARExceedCreditLimit.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit, 1)
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit_Tolerance, CDbl(Me.TxtDO_ToleranceARCreditLimit.Text))
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit, 0)
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit_Tolerance, 0)
                End If
                If Me.chkDO_AROverDue.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue, 1)
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue_Tolerance, CDbl(Me.TxtDO_ToleranceAROverDue.Text))
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue, 0)
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue_Tolerance, 0)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTaxInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTaxInvoice.CheckedChanged
        Try
            If Me.chkTaxInvoice.Checked = True Then
                Me.TxtTaxInvoice.ReadOnly = False
                Me.TxtTaxInvoice.TabStop = True
            Else
                Me.TxtTaxInvoice.ReadOnly = True
                Me.TxtTaxInvoice.TabStop = False
                Me.TxtTaxInvoice.Text = "0"
            End If
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxInvoice.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtTaxInvoice, sender, e, 100)
        Me.SettingInvoiceBerubah = True
    End Sub

    Private Sub TxtTaxInvoice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxInvoice.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtTaxInvoice, sender, e)
    End Sub

    Private Sub SalinKeFieldSettingJurnal()
        Dim Dummy As String = ""

        Try
            'salin data AP
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AP)
            If Dummy = "" Then
                Me.TxtPrimaryAkunAP.Text = "-1"
                Me.TxtAkunAP.Text = ""
            Else
                Me.TxtPrimaryAkunAP.Text = Dummy
                Me.TxtAkunAP.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAP.Text))
            End If
            Me.TxtAPDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAP.Text)

            'salin data AR
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AR)
            If Dummy = "" Then
                Me.TxtPrimaryAkunAR.Text = "-1"
                Me.TxtAkunAR.Text = ""
            Else
                Me.TxtPrimaryAkunAR.Text = Dummy
                Me.TxtAkunAR.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryAkunAR.Text))
            End If
            Me.TxtARDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryAkunAR.Text)

            'custom print
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint)
            If Dummy = "" Then
                Me.chkCustomJurnal.Checked = False
            Else
                If Dummy = "1" Then
                    Me.chkCustomJurnal.Checked = True
                Else
                    Me.chkCustomJurnal.Checked = False
                End If
            End If

            'current year earning
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning)
            If Dummy = "" Then
                Me.TxtPrimaryCurrencyYearEarningAcc.Text = "-1"
                Me.TxtCurrentYearEarningAcc.Text = ""
            Else
                Me.TxtPrimaryCurrencyYearEarningAcc.Text = Dummy
                Me.TxtCurrentYearEarningAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCurrencyYearEarningAcc.Text))
            End If
            Me.TxtCurrenctYeatEarningAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCurrencyYearEarningAcc.Text)

            'retained earning
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Retained_Earning)
            If Dummy = "" Then
                Me.TxtPrimaryRetainedEarningAcc.Text = "-1"
                Me.TxtRetainedEarningAcc.Text = ""
            Else
                Me.TxtPrimaryRetainedEarningAcc.Text = Dummy
                Me.TxtRetainedEarningAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryRetainedEarningAcc.Text))
            End If
            Me.TxtRetainedEarningDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRetainedEarningAcc.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
            If Dummy = "" Then
                Me.TxtPrimaryVATIn.Text = "-1"
                Me.TxtVATIn.Text = ""
            Else
                Me.TxtPrimaryVATIn.Text = Dummy
                Me.TxtVATIn.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATIn.Text))
            End If
            Me.TxtVATInDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATIn.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
            If Dummy = "" Then
                Me.TxtPrimaryVATOut.Text = "-1"
                Me.TxtVATOutAcc.Text = ""
            Else
                Me.TxtPrimaryVATOut.Text = Dummy
                Me.TxtVATOutAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATOut.Text))
            End If
            Me.TxtVATOutAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATOut.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
            If Dummy = "" Then
                Me.TxtPrimaryDutystamp.Text = "-1"
                Me.TxtDutystampAcc.Text = ""
            Else
                Me.TxtPrimaryDutystamp.Text = Dummy
                Me.TxtDutystampAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryDutystamp.Text))
            End If
            Me.TxtDutystampAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryDutystamp.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_PaymentDifference)
            If Dummy = "" Then
                Me.TxtPrimaryPaymentDiff.Text = "-1"
                Me.TxtPaymentDiffAcc.Text = ""
            Else
                Me.TxtPrimaryPaymentDiff.Text = Dummy
                Me.TxtPaymentDiffAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPaymentDiff.Text))
            End If
            Me.TxtPaymentDiffAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPaymentDiff.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_ShippingIncome)
            If Dummy = "" Then
                Me.TxtPrimaryShippingIncome.Text = "-1"
                Me.TxtShippingIncomeAcc.Text = ""
            Else
                Me.TxtPrimaryShippingIncome.Text = Dummy
                Me.TxtShippingIncomeAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryShippingIncome.Text))
            End If
            Me.TxtShippingIncomeAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryShippingIncome.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                       Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
            If Dummy = "" Then
                Me.cboCosting.SelectedIndex = 0
            Else
                Me.cboCosting.SelectedIndex = Dummy
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                     Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait)
            If Dummy = "" Then
                Me.cboPrint.SelectedIndex = 0
            Else
                Me.cboPrint.SelectedIndex = Dummy
            End If

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
            If Dummy = "" Then
                Me.TxtPrimaryPurchaseDPAcc.Text = "-1"
                Me.TxtPurchaseDPAcc.Text = ""
            Else
                Me.TxtPrimaryPurchaseDPAcc.Text = Dummy
                Me.TxtPurchaseDPAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPurchaseDPAcc.Text))
            End If
            Me.TxtPurchaseDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPurchaseDPAcc.Text)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)
            If Dummy = "" Then
                Me.TxtPrimarySalesDPAcc.Text = "-1"
                Me.TxtSalesDPAcc.Text = ""
            Else
                Me.TxtPrimarySalesDPAcc.Text = Dummy
                Me.TxtSalesDPAcc.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesDPAcc.Text))
            End If
            Me.TxtSalesDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesDPAcc.Text)

            'salin data credit note customer
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Credit_Note_Customer)
            If Dummy = "" Then
                Me.TxtCreditNote_PrimaryCustAccount.Text = "-1"
                Me.TxtCreditNote_CustAccountCode.Text = ""
            Else
                Me.TxtCreditNote_PrimaryCustAccount.Text = Dummy
                Me.TxtCreditNote_CustAccountCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimaryCustAccount.Text))
            End If
            Me.TxtCreditNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimaryCustAccount.Text)

            'salin data credit note supplier
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Credit_Note_Supplier)
            If Dummy = "" Then
                Me.TxtCreditNote_PrimarySupplierAccount.Text = "-1"
                Me.TxtCreditNote_SupplierAccountCode.Text = ""
            Else
                Me.TxtCreditNote_PrimarySupplierAccount.Text = Dummy
                Me.TxtCreditNote_SupplierAccountCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimarySupplierAccount.Text))
            End If
            Me.TxtCreditNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimarySupplierAccount.Text)

            'salin data debit note customer
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Debit_Note_Customer)
            If Dummy = "" Then
                Me.TxtDebitNote_PrimaryCustAccount.Text = "-1"
                Me.TxtDebitNote_CustAccountCode.Text = ""
            Else
                Me.TxtDebitNote_PrimaryCustAccount.Text = Dummy
                Me.TxtDebitNote_CustAccountCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimaryCustAccount.Text))
            End If
            Me.TxtDebitNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimaryCustAccount.Text)

            'salin data debit note supplier
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Debit_Note_Supplier)
            If Dummy = "" Then
                Me.TxtDebitNote_PrimarySupplierAccount.Text = "-1"
                Me.TxtDebitNote_SupplierAccountCode.Text = ""
            Else
                Me.TxtDebitNote_PrimarySupplierAccount.Text = Dummy
                Me.TxtDebitNote_SupplierAccountCode.Text = Me.ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimarySupplierAccount.Text))
            End If
            Me.TxtDebitNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimarySupplierAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingReport()
        Dim Dummy As String = ""

        Try
            'salin data Bracket
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            If Dummy = "" Then
                Me.chkUseBracket.Checked = False
            Else
                If Dummy = "3" Then
                    Me.chkUseBracket.Checked = True
                Else
                    Me.chkUseBracket.Checked = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingAOP()
        Dim Dummy As String = ""

        Try
            'salin data use tax
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.ChkTaxAOP.Checked = False
            Else
                If Dummy = "0" Then
                    Me.ChkTaxAOP.Checked = False
                Else
                    Me.ChkTaxAOP.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy <> "" Then
                        Me.TxtTaxAOP.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                    Else
                        Me.TxtTaxAOP.Text = Me.ObjCD.DigitGrouping("0", 2)
                    End If
                End If
            End If

            'salin data print menggunakan custom logic
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "" Then
                Me.chkUsePrintCustomLogicAOP.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUsePrintCustomLogicAOP.Checked = False
                Else
                    Me.chkUsePrintCustomLogicAOP.Checked = True
                End If
            End If

            Me.SettingAOPBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingPO()
        Dim Dummy As String = ""

        Try
            'salin data print menggunakan custom logic
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "" Then
                Me.chkUseCustomPrintLogicPO.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkUseCustomPrintLogicPO.Checked = False
                Else
                    Me.chkUseCustomPrintLogicPO.Checked = True
                End If
            End If
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PreparationDelay)
            If Dummy = "" Then
                Me.TxtPreparationDelay.Text = "0"
            Else
                Me.TxtPreparationDelay.Text = Dummy
            End If
            Me.SettingPOBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeFieldSettingPR()
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.chkTaxPR.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkTaxPR.Checked = False
                Else
                    Me.chkTaxPR.Checked = True
                    Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If Dummy <> "" Then
                        Me.TxtTaxPR.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
                    Else
                        Me.TxtTaxPR.Text = Me.ObjCD.DigitGrouping(0, 2)
                    End If
                End If
            End If

            'salin data print menggunakan custom logic
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic)
            If Dummy = "" Then
                Me.chkCustomPrintLogic_PR.Checked = False
            Else
                If Dummy = "0" Then
                    Me.chkCustomPrintLogic_PR.Checked = False
                Else
                    Me.chkCustomPrintLogic_PR.Checked = True
                End If
            End If

            Me.SettingPRBerubah = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingAOP()
        Try
            If Me.SettingAOPBerubah = True Then
                If Me.ChkTaxAOP.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, Me.TxtTaxAOP.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, "0")
                End If

                If Me.chkUsePrintCustomLogicAOP.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, "0")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingPO()
        Try
            If Me.SettingPOBerubah = True Then
                If Me.chkUseCustomPrintLogicPO.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, "0")
                End If
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PreparationDelay, Me.TxtPreparationDelay.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingPR()
        Try
            If Me.SettingPRBerubah = True Then
                If Me.chkCustomPrintLogic_PR.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, "1")
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic, "0")
                End If

                If Me.chkTaxPR.Checked = True Then
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "1")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, Me.TxtTaxPR.Text)
                Else
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax, "0")
                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage, "0")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingJurnal()
        Try
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AP, Me.TxtPrimaryAkunAP.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesJurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AR, Me.TxtPrimaryAkunAR.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning, Me.TxtPrimaryCurrencyYearEarningAcc.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Retained_Earning, Me.TxtPrimaryRetainedEarningAcc.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In, Me.TxtPrimaryVATIn.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out, Me.TxtPrimaryVATOut.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense, Me.TxtPrimaryDutystamp.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_PaymentDifference, Me.TxtPrimaryPaymentDiff.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_ShippingIncome, Me.TxtPrimaryShippingIncome.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP, Me.TxtPrimaryPurchaseDPAcc.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP, Me.TxtPrimarySalesDPAcc.Text)

            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod, Me.cboCosting.SelectedIndex)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CetakJurnalPortrait, Me.cboPrint.SelectedIndex)

            If Me.chkCustomJurnal.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint, "1")
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.CustomPrint, "0")
            End If

            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Credit_Note_Customer, Me.TxtCreditNote_PrimaryCustAccount.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Credit_Note_Supplier, Me.TxtCreditNote_PrimarySupplierAccount.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Debit_Note_Customer, Me.TxtDebitNote_PrimaryCustAccount.Text)
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Debit_Note_Supplier, Me.TxtDebitNote_PrimarySupplierAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanDataSettingReport()
        Try
            If Me.chkUseBracket.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets, "3")
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets, "1")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTaxAOP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkTaxAOP.CheckedChanged
        Try
            If Me.ChkTaxAOP.Checked = True Then
                Me.TxtTaxAOP.ReadOnly = False
                Me.TxtTaxAOP.TabStop = True
            Else
                Me.TxtTaxAOP.ReadOnly = True
                Me.TxtTaxAOP.TabStop = False
                Me.TxtTaxAOP.Text = "0"
            End If
            Me.SettingAOPBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAOP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxAOP.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtTaxAOP, sender, e, 100)
        Me.SettingAOPBerubah = True
    End Sub

    Private Sub TxtTaxAOP_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxAOP.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtTaxAOP, sender, e)
    End Sub

    Private Sub chkUseCustomInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustomInvoice.CheckedChanged
        Try
            If Me.chkUseCustomInvoice.Checked = True Then
                Me.TxtCustomInvoiceAddress.Enabled = True
                Me.btnFindCustomInvoiceReport.Enabled = True
            Else
                Me.TxtCustomInvoiceAddress.Enabled = False
                Me.btnFindCustomInvoiceReport.Enabled = False
                Me.TxtCustomInvoiceAddress.Clear()
            End If

            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomInvoiceAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomInvoiceAddress.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomInvoiceReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomInvoiceReport.Click
        Try
            Me.OpenFileDialog.Title = "Find Custom Invoice File"
            If Me.OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtCustomInvoiceAddress.Text = Me.OpenFileDialog.FileName
            Else
                Me.TxtCustomInvoiceAddress.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustomTaxInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustomTaxInvoice.CheckedChanged
        Try
            If Me.chkUseCustomTaxInvoice.Checked = True Then
                Me.TxtCustomTaxInvoiceAddress.Enabled = True
                Me.btnFindCustomTaxInvoice.Enabled = True
            Else
                Me.TxtCustomTaxInvoiceAddress.Enabled = False
                Me.btnFindCustomTaxInvoice.Enabled = False
                Me.TxtCustomTaxInvoiceAddress.Clear()
            End If

            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomTaxInvoiceAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomTaxInvoiceAddress.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomTaxInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomTaxInvoice.Click
        Try
            Me.OpenFileDialog.Title = "Find Custom Tax Invoice File"
            If Me.OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtCustomTaxInvoiceAddress.Text = Me.OpenFileDialog.FileName
            Else
                Me.TxtCustomTaxInvoiceAddress.Text = ""
            End If
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustomDO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustomDO.CheckedChanged
        Try
            If Me.chkUseCustomDO.Checked = True Then
                Me.TxtCustomeDOAddress.Enabled = True
                Me.btnFindCustomDOAddress.Enabled = True
            Else
                Me.TxtCustomeDOAddress.Enabled = False
                Me.btnFindCustomDOAddress.Enabled = False
                Me.TxtCustomeDOAddress.Clear()
            End If

            Me.SettingDOBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomeDOAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomeDOAddress.TextChanged
        Try
            Me.SettingDOBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomDOAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomDOAddress.Click
        Try
            Me.OpenFileDialog.Title = "Find Custom Delivery Order File"
            If Me.OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtCustomeDOAddress.Text = Me.OpenFileDialog.FileName
            Else
                Me.TxtCustomeDOAddress.Text = ""
            End If
            Me.SettingDOBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtSignedByName.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByPosition_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtSignedByPosition.TextChanged
        Try
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUsePrintCustomLogicAOP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUsePrintCustomLogicAOP.CheckedChanged
        Try
            Me.SettingAOPBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustomPrintLogicPO_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseCustomPrintLogicPO.CheckedChanged
        Try
            Me.SettingPOBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCustomPrintLogic_PR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomPrintLogic_PR.CheckedChanged
        Try
            Me.SettingPRBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxPRe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxPR.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtTaxPR, sender, e, 100)
        Me.SettingPRBerubah = True
    End Sub

    Private Sub TxtTaxPRe_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxPR.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtTaxPR, sender, e)
    End Sub

    Private Sub chkUseTaxPR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTaxPR.CheckedChanged
        Try
            If Me.chkTaxPR.Checked = True Then
                Me.TxtTaxPR.ReadOnly = False
                Me.TxtTaxPR.TabStop = True
            Else
                Me.TxtTaxPR.ReadOnly = True
                Me.TxtTaxPR.TabStop = False
                Me.TxtTaxPR.Text = "0"
            End If
            Me.SettingPRBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxPR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxPR.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtTaxPR, sender, e, 100)
        Me.SettingPRBerubah = True
    End Sub

    Private Sub TxtTaxPR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxPR.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtTaxPR, sender, e)
    End Sub

    Private Sub TxtCurrentYearEarningAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrentYearEarningAcc.GotFocus
        Try
            If Me.TxtCurrentYearEarningAcc.Enabled = True Then
                If Me.TxtCurrentYearEarningAcc.ReadOnly = True Then
                    Me.TxtCurrentYearEarningAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCurrentYearEarningAcc.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrentYearEarningAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrentYearEarningAcc.LostFocus
        Try
            If Me.TxtCurrentYearEarningAcc.Enabled = True Then
                If Me.TxtCurrentYearEarningAcc.ReadOnly = True Then
                    Me.TxtCurrentYearEarningAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCurrentYearEarningAcc.BackColor = Me.TxtPrimaryCurrencyYearEarningAcc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRetainedEarningAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
       Handles TxtRetainedEarningAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtRetainedEarningAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRetainedEarningAcc.Text) = True Then
                        Me.TxtPrimaryRetainedEarningAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtRetainedEarningAcc.Text))
                        Me.TxtRetainedEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryRetainedEarningAcc.Text))
                        Me.TxtVATIn.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindRetainedEarning.PerformClick()
                    End If
                Else
                    Me.btnFindRetainedEarning.PerformClick()
                End If
            End If
            Me.TxtRetainedEarningDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRetainedEarningAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRetainedEarningAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtRetainedEarningAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtRetainedEarningAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtRetainedEarningAcc.Text) = True Then
                    Me.TxtPrimaryRetainedEarningAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtRetainedEarningAcc.Text))
                    Me.TxtRetainedEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryRetainedEarningAcc.Text))
                Else
                    If Me.TxtRetainedEarningAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtRetainedEarningAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryRetainedEarningAcc.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtRetainedEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryRetainedEarningAcc.Text))
                    Else
                        Me.TxtRetainedEarningAcc.Clear()
                        Me.TxtPrimaryRetainedEarningAcc.Text = "-1"
                        Me.TxtRetainedEarningAcc.Focus()
                    End If
                End If
            Else
                Me.TxtRetainedEarningAcc.Clear()
                Me.TxtPrimaryRetainedEarningAcc.Text = "-1"
            End If
            Me.TxtRetainedEarningDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRetainedEarningAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindRetainedEarning_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindRetainedEarning.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtRetainedEarningAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtRetainedEarningAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryRetainedEarningAcc.Text = CStr(ij.mPrimaryAkun)
                Me.TxtRetainedEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryRetainedEarningAcc.Text))
                Me.TxtRetainedEarningDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRetainedEarningAcc.Text)
                Me.TxtVATIn.Focus()
            Else
                Me.TxtRetainedEarningAcc.Clear()
                Me.TxtPrimaryRetainedEarningAcc.Text = "-1"
                Me.TxtRetainedEarningDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryRetainedEarningAcc.Text)
                Me.TxtRetainedEarningAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    'Account Retained Earning
    Private Sub TxtCurrentYearEarningAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtCurrentYearEarningAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCurrentYearEarningAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCurrentYearEarningAcc.Text) = True Then
                        Me.TxtPrimaryCurrencyYearEarningAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtCurrentYearEarningAcc.Text))
                        Me.TxtCurrentYearEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCurrencyYearEarningAcc.Text))
                        Me.TxtRetainedEarningAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindCurrenctYearEarning.PerformClick()
                    End If
                Else
                    Me.btnFindCurrenctYearEarning.PerformClick()
                End If
            End If
            Me.TxtCurrenctYeatEarningAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCurrencyYearEarningAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrentYearEarningAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCurrentYearEarningAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCurrentYearEarningAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCurrentYearEarningAcc.Text) = True Then
                    Me.TxtPrimaryCurrencyYearEarningAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtCurrentYearEarningAcc.Text))
                    Me.TxtCurrentYearEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCurrencyYearEarningAcc.Text))
                Else
                    If Me.TxtCurrentYearEarningAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtCurrentYearEarningAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryCurrencyYearEarningAcc.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtCurrentYearEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCurrencyYearEarningAcc.Text))
                    Else
                        Me.TxtCurrentYearEarningAcc.Clear()
                        Me.TxtPrimaryCurrencyYearEarningAcc.Text = "-1"
                        Me.TxtCurrentYearEarningAcc.Focus()
                    End If
                End If
            Else
                Me.TxtCurrentYearEarningAcc.Clear()
                Me.TxtPrimaryCurrencyYearEarningAcc.Text = "-1"
            End If
            Me.TxtCurrenctYeatEarningAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCurrencyYearEarningAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindCurrenctYearEarning_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrenctYearEarning.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCurrentYearEarningAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtCurrentYearEarningAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryCurrencyYearEarningAcc.Text = CStr(ij.mPrimaryAkun)
                Me.TxtCurrentYearEarningAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryCurrencyYearEarningAcc.Text))
                Me.TxtCurrenctYeatEarningAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCurrencyYearEarningAcc.Text)
                Me.TxtRetainedEarningAcc.Focus()
            Else
                Me.TxtCurrentYearEarningAcc.Clear()
                Me.TxtPrimaryCurrencyYearEarningAcc.Text = "-1"
                Me.TxtCurrenctYeatEarningAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryCurrencyYearEarningAcc.Text)
                Me.TxtCurrentYearEarningAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtVATIn_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVATIn.GotFocus
        Try
            If Me.TxtVATIn.Enabled = True Then
                If Me.TxtVATIn.ReadOnly = True Then
                    Me.TxtVATIn.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtVATIn.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVATIn_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVATIn.LostFocus
        Try
            If Me.TxtVATIn.Enabled = True Then
                If Me.TxtVATIn.ReadOnly = True Then
                    Me.TxtVATIn.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtVATIn.BackColor = Me.TxtPrimaryVATIn.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVATIn_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtVATIn.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtVATIn.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtVATIn.Text) = True Then
                        Me.TxtPrimaryVATIn.Text = CStr(ObjAkun.GetPrimary(Me.TxtVATIn.Text))
                        Me.TxtVATIn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATIn.Text))
                        Me.TxtVATOutAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindVATIn.PerformClick()
                    End If
                Else
                    Me.btnFindVATIn.PerformClick()
                End If
            End If
            Me.TxtVATInDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATIn.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVATIn_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtVATIn.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtVATIn.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtVATIn.Text) = True Then
                    Me.TxtPrimaryVATIn.Text = CStr(ObjAkun.GetPrimary(Me.TxtVATIn.Text))
                    Me.TxtVATIn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATIn.Text))
                Else
                    If Me.TxtVATIn.Text <> "" Then
                        ij.txtCari.Text = Me.TxtVATIn.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryVATIn.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtVATIn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATIn.Text))
                    Else
                        Me.TxtVATIn.Clear()
                        Me.TxtPrimaryVATIn.Text = "-1"
                        Me.TxtVATIn.Focus()
                    End If
                End If
            Else
                Me.TxtVATIn.Clear()
                Me.TxtPrimaryVATIn.Text = "-1"
            End If
            Me.TxtVATInDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATIn.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindVATIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindVATIn.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtVATIn.Text <> "" Then
                ij.txtCari.Text = Me.TxtVATIn.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryVATIn.Text = CStr(ij.mPrimaryAkun)
                Me.TxtVATIn.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATIn.Text))
                Me.TxtVATInDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATIn.Text)
                Me.TxtVATOutAcc.Focus()
            Else
                Me.TxtVATIn.Clear()
                Me.TxtPrimaryVATIn.Text = "-1"
                Me.TxtVATInDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATIn.Text)
                Me.TxtVATIn.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtVATOutAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVATOutAcc.GotFocus
        Try
            If Me.TxtVATOutAcc.Enabled = True Then
                If Me.TxtVATOutAcc.ReadOnly = True Then
                    Me.TxtVATOutAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtVATOutAcc.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVATOutAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtVATOutAcc.LostFocus
        Try
            If Me.TxtVATOutAcc.Enabled = True Then
                If Me.TxtVATOutAcc.ReadOnly = True Then
                    Me.TxtVATOutAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtVATOutAcc.BackColor = Me.TxtPrimaryVATOut.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account Retained Earning
    Private Sub TxtVATOutAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtVATOutAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtVATOutAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtVATOutAcc.Text) = True Then
                        Me.TxtPrimaryVATOut.Text = CStr(ObjAkun.GetPrimary(Me.TxtVATOutAcc.Text))
                        Me.TxtVATOutAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATOut.Text))
                        Me.TxtDutystampAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindVATOutAcc.PerformClick()
                    End If
                Else
                    Me.btnFindVATOutAcc.PerformClick()
                End If
            End If
            Me.TxtVATOutAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATOut.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtVATOutAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtVATOutAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtVATOutAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtVATOutAcc.Text) = True Then
                    Me.TxtPrimaryVATOut.Text = CStr(ObjAkun.GetPrimary(Me.TxtVATOutAcc.Text))
                    Me.TxtVATOutAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATOut.Text))
                Else
                    If Me.TxtVATOutAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtVATOutAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryVATOut.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtVATOutAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATOut.Text))
                    Else
                        Me.TxtVATOutAcc.Clear()
                        Me.TxtPrimaryVATOut.Text = "-1"
                        Me.TxtVATOutAcc.Focus()
                    End If
                End If
            Else
                Me.TxtVATOutAcc.Clear()
                Me.TxtPrimaryVATOut.Text = "-1"
            End If
            Me.TxtVATOutAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATOut.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindVATOutAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindVATOutAcc.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtVATOutAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtVATOutAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryVATOut.Text = CStr(ij.mPrimaryAkun)
                Me.TxtVATOutAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryVATOut.Text))
                Me.TxtVATOutAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATOut.Text)
                Me.TxtDutystampAcc.Focus()
            Else
                Me.TxtVATOutAcc.Clear()
                Me.TxtPrimaryVATOut.Text = "-1"
                Me.TxtVATOutAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryVATOut.Text)
                Me.TxtVATOutAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtDutystampAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystampAcc.GotFocus
        Try
            If Me.TxtDutystampAcc.Enabled = True Then
                If Me.TxtDutystampAcc.ReadOnly = True Then
                    Me.TxtDutystampAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDutystampAcc.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystampAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystampAcc.LostFocus
        Try
            If Me.TxtDutystampAcc.Enabled = True Then
                If Me.TxtDutystampAcc.ReadOnly = True Then
                    Me.TxtDutystampAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDutystampAcc.BackColor = Me.TxtPrimaryDutystamp.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account Retained Earning
    Private Sub TxtDutystampAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtDutystampAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtDutystampAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDutystampAcc.Text) = True Then
                        Me.TxtPrimaryDutystamp.Text = CStr(ObjAkun.GetPrimary(Me.TxtDutystampAcc.Text))
                        Me.TxtDutystampAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryDutystamp.Text))
                        Me.TxtPaymentDiffAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindDutystamp.PerformClick()
                    End If
                Else
                    Me.btnFindDutystamp.PerformClick()
                End If
            End If
            Me.TxtDutystampAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryDutystamp.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystampAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtDutystampAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDutystampAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDutystampAcc.Text) = True Then
                    Me.TxtPrimaryDutystamp.Text = CStr(ObjAkun.GetPrimary(Me.TxtDutystampAcc.Text))
                    Me.TxtDutystampAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryDutystamp.Text))
                Else
                    If Me.TxtDutystampAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtDutystampAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryDutystamp.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtDutystampAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryDutystamp.Text))
                    Else
                        Me.TxtDutystampAcc.Clear()
                        Me.TxtPrimaryDutystamp.Text = "-1"
                        Me.TxtDutystampAcc.Focus()
                    End If
                End If
            Else
                Me.TxtDutystampAcc.Clear()
                Me.TxtPrimaryDutystamp.Text = "-1"
            End If
            Me.TxtDutystampAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryDutystamp.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindDutystamp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDutystamp.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDutystampAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtDutystampAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryDutystamp.Text = CStr(ij.mPrimaryAkun)
                Me.TxtDutystampAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryDutystamp.Text))
                Me.TxtDutystampAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryDutystamp.Text)
                Me.TxtPaymentDiffAcc.Focus()
            Else
                Me.TxtDutystampAcc.Clear()
                Me.TxtPrimaryDutystamp.Text = "-1"
                Me.TxtDutystampAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryDutystamp.Text)
                Me.TxtDutystampAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtPaymentDiffAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentDiffAcc.GotFocus
        Try
            If Me.TxtPaymentDiffAcc.Enabled = True Then
                If Me.TxtPaymentDiffAcc.ReadOnly = True Then
                    Me.TxtPaymentDiffAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPaymentDiffAcc.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDiffAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentDiffAcc.LostFocus
        Try
            If Me.TxtPaymentDiffAcc.Enabled = True Then
                If Me.TxtPaymentDiffAcc.ReadOnly = True Then
                    Me.TxtPaymentDiffAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPaymentDiffAcc.BackColor = Me.TxtPrimaryPaymentDiff.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Account Retained Earning
    Private Sub TxtPaymentDiffAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtPaymentDiffAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtPaymentDiffAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtPaymentDiffAcc.Text) = True Then
                        Me.TxtPrimaryPaymentDiff.Text = CStr(ObjAkun.GetPrimary(Me.TxtPaymentDiffAcc.Text))
                        Me.TxtPaymentDiffAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPaymentDiff.Text))
                        Me.TxtShippingIncomeAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindPaymentDiff.PerformClick()
                    End If
                Else
                    Me.btnFindPaymentDiff.PerformClick()
                End If
            End If
            Me.TxtPaymentDiffAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPaymentDiff.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentDiffAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtPaymentDiffAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPaymentDiffAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtPaymentDiffAcc.Text) = True Then
                    Me.TxtPrimaryPaymentDiff.Text = CStr(ObjAkun.GetPrimary(Me.TxtPaymentDiffAcc.Text))
                    Me.TxtPaymentDiffAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPaymentDiff.Text))
                Else
                    If Me.TxtPaymentDiffAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtPaymentDiffAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryPaymentDiff.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtPaymentDiffAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPaymentDiff.Text))
                    Else
                        Me.TxtPaymentDiffAcc.Clear()
                        Me.TxtPrimaryPaymentDiff.Text = "-1"
                        Me.TxtPaymentDiffAcc.Focus()
                    End If
                End If
            Else
                Me.TxtPaymentDiffAcc.Clear()
                Me.TxtPrimaryPaymentDiff.Text = "-1"
            End If
            Me.TxtPaymentDiffAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPaymentDiff.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindPaymentDiff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPaymentDiff.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPaymentDiffAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtPaymentDiffAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryPaymentDiff.Text = CStr(ij.mPrimaryAkun)
                Me.TxtPaymentDiffAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPaymentDiff.Text))
                Me.TxtPaymentDiffAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPaymentDiff.Text)
                Me.TxtShippingIncomeAcc.Focus()
            Else
                Me.TxtPaymentDiffAcc.Clear()
                Me.TxtPrimaryPaymentDiff.Text = "-1"
                Me.TxtPaymentDiffAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPaymentDiff.Text)
                Me.TxtPaymentDiffAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    'Account Shipping Income
    Private Sub TxtShippingIncomeAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShippingIncomeAcc.GotFocus
        Try
            If Me.TxtShippingIncomeAcc.Enabled = True Then
                If Me.TxtShippingIncomeAcc.ReadOnly = True Then
                    Me.TxtShippingIncomeAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtShippingIncomeAcc.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShippingIncomeAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShippingIncomeAcc.LostFocus
        Try
            If Me.TxtShippingIncomeAcc.Enabled = True Then
                If Me.TxtShippingIncomeAcc.ReadOnly = True Then
                    Me.TxtShippingIncomeAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtShippingIncomeAcc.BackColor = Me.TxtPrimaryShippingIncome.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShippingIncomeAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtShippingIncomeAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtShippingIncomeAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtShippingIncomeAcc.Text) = True Then
                        Me.TxtPrimaryShippingIncome.Text = CStr(ObjAkun.GetPrimary(Me.TxtShippingIncomeAcc.Text))
                        Me.TxtShippingIncomeAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryShippingIncome.Text))
                        Me.btnOK.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindShippingIncomeAcc.PerformClick()
                    End If
                Else
                    Me.btnFindShippingIncomeAcc.PerformClick()
                End If
            End If
            Me.TxtShippingIncomeAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryShippingIncome.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShippingIncomeAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtShippingIncomeAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtShippingIncomeAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtShippingIncomeAcc.Text) = True Then
                    Me.TxtPrimaryShippingIncome.Text = CStr(ObjAkun.GetPrimary(Me.TxtShippingIncomeAcc.Text))
                    Me.TxtShippingIncomeAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryShippingIncome.Text))
                Else
                    If Me.TxtShippingIncomeAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtShippingIncomeAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryShippingIncome.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtShippingIncomeAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryShippingIncome.Text))
                    Else
                        Me.TxtShippingIncomeAcc.Clear()
                        Me.TxtPrimaryShippingIncome.Text = "-1"
                        Me.TxtShippingIncomeAcc.Focus()
                    End If
                End If
            Else
                Me.TxtShippingIncomeAcc.Clear()
                Me.TxtPrimaryShippingIncome.Text = "-1"
            End If
            Me.TxtShippingIncomeAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryShippingIncome.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindShippingIncomeAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindShippingIncomeAcc.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtShippingIncomeAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtShippingIncomeAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryShippingIncome.Text = CStr(ij.mPrimaryAkun)
                Me.TxtShippingIncomeAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryShippingIncome.Text))
                Me.TxtShippingIncomeAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryShippingIncome.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtShippingIncomeAcc.Clear()
                Me.TxtPrimaryShippingIncome.Text = "-1"
                Me.TxtShippingIncomeAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryShippingIncome.Text)
                Me.TxtShippingIncomeAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtPurchaseDPAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtPurchaseDPAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtPurchaseDPAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtPurchaseDPAcc.Text) = True Then
                        Me.TxtPrimaryPurchaseDPAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtPurchaseDPAcc.Text))
                        Me.TxtPurchaseDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPurchaseDPAcc.Text))
                        Me.TxtVATIn.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindPurchaseDP.PerformClick()
                    End If
                Else
                    Me.btnFindPurchaseDP.PerformClick()
                End If
            End If
            Me.TxtPurchaseDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPurchaseDPAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseDPAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtPurchaseDPAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPurchaseDPAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtPurchaseDPAcc.Text) = True Then
                    Me.TxtPrimaryPurchaseDPAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtPurchaseDPAcc.Text))
                    Me.TxtPurchaseDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPurchaseDPAcc.Text))
                Else
                    If Me.TxtPurchaseDPAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtPurchaseDPAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryPurchaseDPAcc.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtPurchaseDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPurchaseDPAcc.Text))
                    Else
                        Me.TxtPurchaseDPAcc.Clear()
                        Me.TxtPrimaryPurchaseDPAcc.Text = "-1"
                        Me.TxtPurchaseDPAcc.Focus()
                    End If
                End If
            Else
                Me.TxtPurchaseDPAcc.Clear()
                Me.TxtPrimaryPurchaseDPAcc.Text = "-1"
            End If
            Me.TxtPurchaseDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPurchaseDPAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindPurchaseDP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPurchaseDP.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtPurchaseDPAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtPurchaseDPAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimaryPurchaseDPAcc.Text = CStr(ij.mPrimaryAkun)
                Me.TxtPurchaseDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimaryPurchaseDPAcc.Text))
                Me.TxtPurchaseDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPurchaseDPAcc.Text)
                Me.TxtVATIn.Focus()
            Else
                Me.TxtPurchaseDPAcc.Clear()
                Me.TxtPrimaryPurchaseDPAcc.Text = "-1"
                Me.TxtPurchaseDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimaryPurchaseDPAcc.Text)
                Me.TxtPurchaseDPAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesDPAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
         Handles TxtSalesDPAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtSalesDPAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesDPAcc.Text) = True Then
                        Me.TxtPrimarySalesDPAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtSalesDPAcc.Text))
                        Me.TxtSalesDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesDPAcc.Text))
                        Me.TxtVATIn.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindSalesDP.PerformClick()
                    End If
                Else
                    Me.btnFindSalesDP.PerformClick()
                End If
            End If
            Me.TxtSalesDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesDPAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesDPAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtSalesDPAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSalesDPAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesDPAcc.Text) = True Then
                    Me.TxtPrimarySalesDPAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtSalesDPAcc.Text))
                    Me.TxtSalesDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesDPAcc.Text))
                Else
                    If Me.TxtSalesDPAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtSalesDPAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimarySalesDPAcc.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtSalesDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesDPAcc.Text))
                    Else
                        Me.TxtSalesDPAcc.Clear()
                        Me.TxtPrimarySalesDPAcc.Text = "-1"
                        Me.TxtSalesDPAcc.Focus()
                    End If
                End If
            Else
                Me.TxtSalesDPAcc.Clear()
                Me.TxtPrimarySalesDPAcc.Text = "-1"
            End If
            Me.TxtSalesDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesDPAcc.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesDP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesDP.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSalesDPAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtSalesDPAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimarySalesDPAcc.Text = CStr(ij.mPrimaryAkun)
                Me.TxtSalesDPAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesDPAcc.Text))
                Me.TxtSalesDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesDPAcc.Text)
                Me.TxtVATIn.Focus()
            Else
                Me.TxtSalesDPAcc.Clear()
                Me.TxtPrimarySalesDPAcc.Text = "-1"
                Me.TxtSalesDPAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesDPAcc.Text)
                Me.TxtSalesDPAcc.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub chkShipmentTax_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShipmentTax.CheckedChanged
        Me.SettingInvoiceBerubah = True
    End Sub

    Private Sub TxtSalesGlobalDiscAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesGlobalDiscAcc.GotFocus
        Try
            If Me.TxtSalesGlobalDiscAcc.Enabled = True Then
                If Me.TxtSalesGlobalDiscAcc.ReadOnly = True Then
                    Me.TxtSalesGlobalDiscAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSalesGlobalDiscAcc.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesGlobalDiscAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesGlobalDiscAcc.LostFocus
        Try
            If Me.TxtSalesGlobalDiscAcc.Enabled = True Then
                If Me.TxtSalesGlobalDiscAcc.ReadOnly = True Then
                    Me.TxtSalesGlobalDiscAcc.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSalesGlobalDiscAcc.BackColor = Me.TxtPrimarySalesGlobalDiscAcc.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesGlobalDiscAcc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles TxtSalesGlobalDiscAcc.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtSalesGlobalDiscAcc.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesGlobalDiscAcc.Text) = True Then
                        Me.TxtPrimarySalesGlobalDiscAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtSalesGlobalDiscAcc.Text))
                        Me.TxtSalesGlobalDiscAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesGlobalDiscAcc.Text))
                        Me.TxtVATOutAcc.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindSalesGlobalDisc.PerformClick()
                    End If
                Else
                    Me.btnFindSalesGlobalDisc.PerformClick()
                End If
            End If
            Me.TxtSalesGlobalDiscAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesGlobalDiscAcc.Text)
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesGlobalDiscAcc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtSalesGlobalDiscAcc.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtSalesGlobalDiscAcc.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtSalesGlobalDiscAcc.Text) = True Then
                    Me.TxtPrimarySalesGlobalDiscAcc.Text = CStr(ObjAkun.GetPrimary(Me.TxtSalesGlobalDiscAcc.Text))
                    Me.TxtSalesGlobalDiscAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesGlobalDiscAcc.Text))
                Else
                    If Me.TxtSalesGlobalDiscAcc.Text <> "" Then
                        ij.txtCari.Text = Me.TxtSalesGlobalDiscAcc.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtPrimarySalesGlobalDiscAcc.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtSalesGlobalDiscAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesGlobalDiscAcc.Text))
                    Else
                        Me.TxtSalesGlobalDiscAcc.Clear()
                        Me.TxtPrimarySalesGlobalDiscAcc.Text = "-1"
                        Me.TxtSalesGlobalDiscAcc.Focus()
                    End If
                End If
            Else
                Me.TxtSalesGlobalDiscAcc.Clear()
                Me.TxtPrimarySalesGlobalDiscAcc.Text = "-1"
            End If
            Me.TxtSalesGlobalDiscAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesGlobalDiscAcc.Text)
            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnFindGlobalDisc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesGlobalDisc.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtSalesGlobalDiscAcc.Text <> "" Then
                ij.txtCari.Text = Me.TxtSalesGlobalDiscAcc.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtPrimarySalesGlobalDiscAcc.Text = CStr(ij.mPrimaryAkun)
                Me.TxtSalesGlobalDiscAcc.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtPrimarySalesGlobalDiscAcc.Text))
                Me.TxtSalesGlobalDiscAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesGlobalDiscAcc.Text)
                Me.TxtVATOutAcc.Focus()
            Else
                Me.TxtSalesGlobalDiscAcc.Clear()
                Me.TxtPrimarySalesGlobalDiscAcc.Text = "-1"
                Me.TxtSalesGlobalDiscAccDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtPrimarySalesGlobalDiscAcc.Text)
                Me.TxtSalesGlobalDiscAcc.Focus()
            End If

            Me.SettingInvoiceBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub chkIgnoreEarlyIn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.SettingHRBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkIgnoreEarlyOut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.SettingHRBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkIgnoreBreak_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.SettingHRBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingType_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Me.Buffer = Me.cboRoundingType.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingType_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Me.cboRoundingType.SelectedIndex <> Me.Buffer Then
                Me.SettingHRBerubah = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Me.Buffer = Me.TxtRoundingUnit.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Me.TxtRoundingUnit.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Me.ObjCD.Is_Time_String(Me.TxtRoundingUnit.Text) = False Then
                Me.TxtRoundingUnit.Text = Me.Buffer
            Else
                If Me.ObjCD.Is_Time_String(Me.Buffer) = True Then
                    If Me.ObjCD.Convert_ToMinutes(Me.Buffer) <> Me.ObjCD.Convert_ToMinutes(Me.TxtRoundingUnit.Text) Then
                        Me.SettingHRBerubah = True
                    End If
                Else
                    Me.SettingHRBerubah = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Me.TxtRoundingUnit.BackColor = Me.TxtPrimaryAkunAP.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboIncludeCalculation_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboIncludeCalculation.SelectedIndexChanged
        Me.SettingInvoiceBerubah = True
    End Sub

    Private Sub ChkSO_AR_ExceedLimit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkSO_AR_ExceedLimit.CheckedChanged
        Me.SettingSOBerubah = True
    End Sub

    Private Sub ChkSO_AR_Overdue_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkSO_AR_Overdue.CheckedChanged
        Me.SettingSOBerubah = True
    End Sub

    Private Sub TxtMaterialPlanning_WorkDays_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaterialPlanning_WorkDays.Enter
        Try
            Me.Buffer = Me.TxtMaterialPlanning_WorkDays.Text
            Me.ObjCD.Number_Enter(Me.TxtMaterialPlanning_WorkDays, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaterialPlanning_WorkDays_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaterialPlanning_WorkDays.GotFocus
        Try
            Me.TxtMaterialPlanning_WorkDays.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaterialPlanning_WorkDays_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMaterialPlanning_WorkDays.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtMaterialPlanning_WorkDays, sender, e, 31)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaterialPlanning_WorkDays_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMaterialPlanning_WorkDays.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtMaterialPlanning_WorkDays, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaterialPlanning_WorkDays_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaterialPlanning_WorkDays.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtMaterialPlanning_WorkDays, sender, e, 0, Me.ObjCD)
            If CDbl(Me.TxtMaterialPlanning_WorkDays.Text) <> CDbl(Me.Buffer) Then
                Me.SettingInventoryBerubah = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMaterialPlanning_WorkDays_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMaterialPlanning_WorkDays.LostFocus
        Try
            Me.TxtMaterialPlanning_WorkDays.BackColor = Me.TxtPrimaryAkunAP.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparationDelay_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparationDelay.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtPreparationDelay, sender, e)
            Me.Buffer = Me.TxtPreparationDelay.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparationDelay_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPreparationDelay.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPreparationDelay, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparationDelay_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPreparationDelay.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPreparationDelay, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparationDelay_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparationDelay.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtPreparationDelay, sender, e, 0, Me.ObjCD)
            If CInt(Me.TxtPreparationDelay.Text) <> CInt(Me.Buffer) Then
                Me.SettingPOBerubah = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkDO_AROverDue_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDO_AROverDue.CheckedChanged
        Try
            If Me.chkDO_AROverDue.Checked = True Then
                Me.TxtDO_ToleranceAROverDue.ReadOnly = False
                Me.TxtDO_ToleranceAROverDue.BackColor = Me.TxtPrimaryAkunAP.BackColor
                Me.TxtDO_ToleranceAROverDue.TabStop = True
            Else
                Me.TxtDO_ToleranceAROverDue.ReadOnly = True
                Me.TxtDO_ToleranceAROverDue.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtDO_ToleranceAROverDue.TabStop = False
                Me.TxtDO_ToleranceAROverDue.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkDO_ARExceedCreditLimit_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDO_ARExceedCreditLimit.CheckedChanged
        Try
            If Me.chkDO_ARExceedCreditLimit.Checked = True Then
                Me.TxtDO_ToleranceARCreditLimit.ReadOnly = False
                Me.TxtDO_ToleranceARCreditLimit.BackColor = Me.TxtPrimaryAkunAP.BackColor
                Me.TxtDO_ToleranceARCreditLimit.TabStop = True
            Else
                Me.TxtDO_ToleranceARCreditLimit.ReadOnly = True
                Me.TxtDO_ToleranceARCreditLimit.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtDO_ToleranceARCreditLimit.TabStop = False
                Me.TxtDO_ToleranceARCreditLimit.Text = Me.ObjCD.DigitGrouping(0, Me.ObjCurrency.GetDecimal_DefaultCurrency)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceAROverDue_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceAROverDue.GotFocus
        Try
            If Me.TxtDO_ToleranceAROverDue.ReadOnly = True Then
                Me.TxtDO_ToleranceAROverDue.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtDO_ToleranceAROverDue.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceAROverDue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceAROverDue.LostFocus
        Try
            If Me.TxtDO_ToleranceAROverDue.ReadOnly = True Then
                Me.TxtDO_ToleranceAROverDue.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtDO_ToleranceAROverDue.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceAROverDue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceAROverDue.Enter
        Try
            If Me.TxtDO_ToleranceAROverDue.ReadOnly = False Then
                Me.Buffer = Me.TxtDO_ToleranceAROverDue.Text
                Me.ObjCD.Number_Enter(Me.TxtDO_ToleranceAROverDue, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceAROverDue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDO_ToleranceAROverDue.KeyDown
        Try
            If Me.TxtDO_ToleranceAROverDue.ReadOnly = False Then
                Me.ObjCD.Number_KeyDown(Me.TxtDO_ToleranceAROverDue, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceAROverDue_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDO_ToleranceAROverDue.KeyUp
        Try
            If Me.TxtDO_ToleranceAROverDue.ReadOnly = False Then
                Me.ObjCD.Number_KeyUp(Me.TxtDO_ToleranceAROverDue, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceAROverDue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceAROverDue.Leave
        Try
            If Me.TxtDO_ToleranceAROverDue.ReadOnly = False Then
                Me.ObjCD.Number_Leave(Me.TxtDO_ToleranceAROverDue, sender, e, Me.ObjCurrency.GetDecimal_DefaultCurrency, Me.ObjCD)
                If CDec(Me.Buffer) <> CDec(Me.TxtDO_ToleranceAROverDue.Text) Then
                    Me.SettingDOBerubah = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceARCreditLimit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceARCreditLimit.GotFocus
        Try
            If Me.TxtDO_ToleranceARCreditLimit.ReadOnly = True Then
                Me.TxtDO_ToleranceARCreditLimit.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtDO_ToleranceARCreditLimit.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceARCreditLimit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceARCreditLimit.Enter
        Try
            If Me.TxtDO_ToleranceARCreditLimit.ReadOnly = False Then
                Me.Buffer = Me.TxtDO_ToleranceARCreditLimit.Text
                Me.ObjCD.Number_Enter(Me.TxtDO_ToleranceARCreditLimit, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceARCreditLimit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDO_ToleranceARCreditLimit.KeyDown
        Try
            If Me.TxtDO_ToleranceARCreditLimit.ReadOnly = False Then
                Me.ObjCD.Number_KeyDown(Me.TxtDO_ToleranceARCreditLimit, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceARCreditLimit_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDO_ToleranceARCreditLimit.KeyUp
        Try
            If Me.TxtDO_ToleranceARCreditLimit.ReadOnly = False Then
                Me.ObjCD.Number_KeyUp(Me.TxtDO_ToleranceARCreditLimit, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceARCreditLimit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceARCreditLimit.Leave
        Try
            If Me.TxtDO_ToleranceARCreditLimit.ReadOnly = False Then
                Me.ObjCD.Number_Leave(Me.TxtDO_ToleranceARCreditLimit, sender, e, Me.ObjCurrency.GetDecimal_DefaultCurrency, Me.ObjCD)
                If CDec(Me.TxtDO_ToleranceARCreditLimit.Text) <> CDec(Me.Buffer) Then
                    Me.SettingDOBerubah = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDO_ToleranceARCreditLimit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDO_ToleranceARCreditLimit.LostFocus
        Try
            If Me.TxtDO_ToleranceARCreditLimit.ReadOnly = True Then
                Me.TxtDO_ToleranceARCreditLimit.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtDO_ToleranceARCreditLimit.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSignedByName.GotFocus
        Try
            If Me.TxtSignedByName.ReadOnly = True Then
                Me.TxtSignedByName.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSignedByName.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSignedByName.LostFocus
        Try
            If Me.TxtSignedByName.ReadOnly = True Then
                Me.TxtSignedByName.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSignedByName.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByPosition_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSignedByPosition.GotFocus
        Try
            If Me.TxtSignedByPosition.ReadOnly = True Then
                Me.TxtSignedByPosition.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSignedByPosition.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSignedByPosition_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSignedByPosition.LostFocus
        Try
            If Me.TxtSignedByPosition.ReadOnly = True Then
                Me.TxtSignedByPosition.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSignedByPosition.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoice.GotFocus
        Try
            If Me.TxtTaxInvoice.ReadOnly = True Then
                Me.TxtTaxInvoice.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtTaxInvoice.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxInvoice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxInvoice.LostFocus
        Try
            If Me.TxtTaxInvoice.ReadOnly = True Then
                Me.TxtTaxInvoice.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtTaxInvoice.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseDPAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseDPAcc.GotFocus
        Try
            If Me.TxtPurchaseDPAcc.ReadOnly = True Then
                Me.TxtPurchaseDPAcc.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtPurchaseDPAcc.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseDPAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseDPAcc.LostFocus
        Try
            If Me.TxtPurchaseDPAcc.ReadOnly = True Then
                Me.TxtPurchaseDPAcc.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtPurchaseDPAcc.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesDPAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesDPAcc.GotFocus
        Try
            If Me.TxtSalesDPAcc.ReadOnly = True Then
                Me.TxtSalesDPAcc.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSalesDPAcc.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesDPAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesDPAcc.LostFocus
        Try
            If Me.TxtSalesDPAcc.ReadOnly = True Then
                Me.TxtSalesDPAcc.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtSalesDPAcc.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRetainedEarningAcc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRetainedEarningAcc.GotFocus
        Try
            If Me.TxtRetainedEarningAcc.ReadOnly = True Then
                Me.TxtRetainedEarningAcc.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtRetainedEarningAcc.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRetainedEarningAcc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRetainedEarningAcc.LostFocus
        Try
            If Me.TxtRetainedEarningAcc.ReadOnly = True Then
                Me.TxtRetainedEarningAcc.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtRetainedEarningAcc.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparationDelay_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparationDelay.GotFocus
        Try
            If Me.TxtPreparationDelay.ReadOnly = True Then
                Me.TxtPreparationDelay.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtPreparationDelay.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPreparationDelay_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPreparationDelay.LostFocus
        Try
            If Me.TxtPreparationDelay.ReadOnly = True Then
                Me.TxtPreparationDelay.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtPreparationDelay.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAOP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAOP.GotFocus
        Try
            If Me.TxtTaxAOP.ReadOnly = True Then
                Me.TxtTaxAOP.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtTaxAOP.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxAOP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxAOP.LostFocus
        Try
            If Me.TxtTaxAOP.ReadOnly = True Then
                Me.TxtTaxAOP.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtTaxAOP.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxPR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxPR.GotFocus
        Try
            If Me.TxtTaxPR.ReadOnly = True Then
                Me.TxtTaxPR.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtTaxPR.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxPR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxPR.LostFocus
        Try
            If Me.TxtTaxPR.ReadOnly = True Then
                Me.TxtTaxPR.BackColor = Me.TxtReadOnly.BackColor
            Else
                Me.TxtTaxPR.BackColor = Me.TxtPrimaryAkunAP.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditNote_CustAccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditNote_CustAccountCode.GotFocus
        Try
            If Me.TxtCreditNote_CustAccountCode.Enabled = True Then
                If Me.TxtCreditNote_CustAccountCode.ReadOnly = True Then
                    Me.TxtCreditNote_CustAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCreditNote_CustAccountCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditNote_CustAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditNote_CustAccountCode.LostFocus
        Try
            If Me.TxtCreditNote_CustAccountCode.Enabled = True Then
                If Me.TxtCreditNote_CustAccountCode.ReadOnly = True Then
                    Me.TxtCreditNote_CustAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCreditNote_CustAccountCode.BackColor = Me.TxtCreditNote_PrimaryCustAccount.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditNote_CustAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
           Handles TxtCreditNote_CustAccountCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCreditNote_CustAccountCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCreditNote_CustAccountCode.Text) = True Then
                        Me.TxtCreditNote_PrimaryCustAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtCreditNote_CustAccountCode.Text))
                        Me.TxtCreditNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimaryCustAccount.Text))
                        Me.TxtCreditNote_SupplierAccountCode.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnCreditNote_FindCustAccount.PerformClick()
                    End If
                Else
                    Me.BtnCreditNote_FindCustAccount.PerformClick()
                End If
            End If
            Me.TxtCreditNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimaryCustAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditNote_CustAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCreditNote_CustAccountCode.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCreditNote_CustAccountCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCreditNote_CustAccountCode.Text) = True Then
                    Me.TxtCreditNote_PrimaryCustAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtCreditNote_CustAccountCode.Text))
                    Me.TxtCreditNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimaryCustAccount.Text))
                Else
                    If Me.TxtCreditNote_CustAccountCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtCreditNote_CustAccountCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtCreditNote_PrimaryCustAccount.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtCreditNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimaryCustAccount.Text))
                    Else
                        Me.TxtCreditNote_CustAccountCode.Clear()
                        Me.TxtCreditNote_PrimaryCustAccount.Text = "-1"
                        Me.TxtCreditNote_CustAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtCreditNote_CustAccountCode.Clear()
                Me.TxtCreditNote_PrimaryCustAccount.Text = "-1"
            End If
            Me.TxtCreditNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimaryCustAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnCreditNote_FindIncomeAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCreditNote_FindCustAccount.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCreditNote_CustAccountCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtCreditNote_CustAccountCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtCreditNote_PrimaryCustAccount.Text = CStr(ij.mPrimaryAkun)
                Me.TxtCreditNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimaryCustAccount.Text))
                Me.TxtCreditNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimaryCustAccount.Text)
                Me.TxtCreditNote_SupplierAccountCode.Focus()
            Else
                Me.TxtCreditNote_CustAccountCode.Clear()
                Me.TxtCreditNote_PrimaryCustAccount.Text = "-1"
                Me.TxtCreditNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimaryCustAccount.Text)
                Me.TxtCreditNote_CustAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtCreditNote_SupplierAccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditNote_SupplierAccountCode.GotFocus
        Try
            If Me.TxtCreditNote_SupplierAccountCode.Enabled = True Then
                If Me.TxtCreditNote_SupplierAccountCode.ReadOnly = True Then
                    Me.TxtCreditNote_SupplierAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCreditNote_SupplierAccountCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditNote_SupplierAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCreditNote_SupplierAccountCode.LostFocus
        Try
            If Me.TxtCreditNote_SupplierAccountCode.Enabled = True Then
                If Me.TxtCreditNote_SupplierAccountCode.ReadOnly = True Then
                    Me.TxtCreditNote_SupplierAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtCreditNote_SupplierAccountCode.BackColor = Me.TxtCreditNote_PrimarySupplierAccount.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditNote_SupplierAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
           Handles TxtCreditNote_SupplierAccountCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCreditNote_SupplierAccountCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCreditNote_SupplierAccountCode.Text) = True Then
                        Me.TxtCreditNote_PrimarySupplierAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtCreditNote_SupplierAccountCode.Text))
                        Me.TxtCreditNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimarySupplierAccount.Text))
                        Me.btnOK.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnCreditNote_FindSupplierAccount.PerformClick()
                    End If
                Else
                    Me.BtnCreditNote_FindSupplierAccount.PerformClick()
                End If
            End If
            Me.TxtCreditNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimarySupplierAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCreditNote_SupplierAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCreditNote_SupplierAccountCode.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCreditNote_SupplierAccountCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtCreditNote_SupplierAccountCode.Text) = True Then
                    Me.TxtCreditNote_PrimarySupplierAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtCreditNote_SupplierAccountCode.Text))
                    Me.TxtCreditNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimarySupplierAccount.Text))
                Else
                    If Me.TxtCreditNote_SupplierAccountCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtCreditNote_SupplierAccountCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtCreditNote_PrimarySupplierAccount.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtCreditNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimarySupplierAccount.Text))
                    Else
                        Me.TxtCreditNote_SupplierAccountCode.Clear()
                        Me.TxtCreditNote_PrimarySupplierAccount.Text = "-1"
                        Me.btnOK.Focus()
                    End If
                End If
            Else
                Me.TxtCreditNote_SupplierAccountCode.Clear()
                Me.TxtCreditNote_PrimarySupplierAccount.Text = "-1"
            End If
            Me.TxtCreditNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimarySupplierAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnCreditNote_FindExpenseAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCreditNote_FindSupplierAccount.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCreditNote_SupplierAccountCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtCreditNote_SupplierAccountCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtCreditNote_PrimarySupplierAccount.Text = CStr(ij.mPrimaryAkun)
                Me.TxtCreditNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtCreditNote_PrimarySupplierAccount.Text))
                Me.TxtCreditNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimarySupplierAccount.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtCreditNote_SupplierAccountCode.Clear()
                Me.TxtCreditNote_PrimarySupplierAccount.Text = "-1"
                Me.TxtCreditNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtCreditNote_PrimarySupplierAccount.Text)
                Me.btnOK.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtDebitNote_CustAccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDebitNote_CustAccountCode.GotFocus
        Try
            If Me.TxtDebitNote_CustAccountCode.Enabled = True Then
                If Me.TxtDebitNote_CustAccountCode.ReadOnly = True Then
                    Me.TxtDebitNote_CustAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDebitNote_CustAccountCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDebitNote_CustAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDebitNote_CustAccountCode.LostFocus
        Try
            If Me.TxtDebitNote_CustAccountCode.Enabled = True Then
                If Me.TxtDebitNote_CustAccountCode.ReadOnly = True Then
                    Me.TxtDebitNote_CustAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDebitNote_CustAccountCode.BackColor = Me.TxtDebitNote_PrimaryCustAccount.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDebitNote_CustAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
           Handles TxtDebitNote_CustAccountCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtDebitNote_CustAccountCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDebitNote_CustAccountCode.Text) = True Then
                        Me.TxtDebitNote_PrimaryCustAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtDebitNote_CustAccountCode.Text))
                        Me.TxtDebitNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimaryCustAccount.Text))
                        Me.TxtDebitNote_SupplierAccountCode.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnDebitNote_FindCustAccount.PerformClick()
                    End If
                Else
                    Me.BtnDebitNote_FindCustAccount.PerformClick()
                End If
            End If
            Me.TxtDebitNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimaryCustAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDebitNote_CustAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtDebitNote_CustAccountCode.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDebitNote_CustAccountCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDebitNote_CustAccountCode.Text) = True Then
                    Me.TxtDebitNote_PrimaryCustAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtDebitNote_CustAccountCode.Text))
                    Me.TxtDebitNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimaryCustAccount.Text))
                Else
                    If Me.TxtDebitNote_CustAccountCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtDebitNote_CustAccountCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtDebitNote_PrimaryCustAccount.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtDebitNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimaryCustAccount.Text))
                    Else
                        Me.TxtDebitNote_CustAccountCode.Clear()
                        Me.TxtDebitNote_PrimaryCustAccount.Text = "-1"
                        Me.TxtDebitNote_CustAccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtDebitNote_CustAccountCode.Clear()
                Me.TxtDebitNote_PrimaryCustAccount.Text = "-1"
            End If
            Me.TxtDebitNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimaryCustAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnDebitNote_FindIncomeAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDebitNote_FindCustAccount.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDebitNote_CustAccountCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtDebitNote_CustAccountCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtDebitNote_PrimaryCustAccount.Text = CStr(ij.mPrimaryAkun)
                Me.TxtDebitNote_CustAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimaryCustAccount.Text))
                Me.TxtDebitNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimaryCustAccount.Text)
                Me.TxtDebitNote_SupplierAccountCode.Focus()
            Else
                Me.TxtDebitNote_CustAccountCode.Clear()
                Me.TxtDebitNote_PrimaryCustAccount.Text = "-1"
                Me.TxtDebitNote_CustAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimaryCustAccount.Text)
                Me.TxtDebitNote_CustAccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtDebitNote_SupplierAccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDebitNote_SupplierAccountCode.GotFocus
        Try
            If Me.TxtDebitNote_SupplierAccountCode.Enabled = True Then
                If Me.TxtDebitNote_SupplierAccountCode.ReadOnly = True Then
                    Me.TxtDebitNote_SupplierAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDebitNote_SupplierAccountCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDebitNote_SupplierAccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDebitNote_SupplierAccountCode.LostFocus
        Try
            If Me.TxtDebitNote_SupplierAccountCode.Enabled = True Then
                If Me.TxtDebitNote_SupplierAccountCode.ReadOnly = True Then
                    Me.TxtDebitNote_SupplierAccountCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDebitNote_SupplierAccountCode.BackColor = Me.TxtDebitNote_PrimarySupplierAccount.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDebitNote_SupplierAccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
           Handles TxtDebitNote_SupplierAccountCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtDebitNote_SupplierAccountCode.Text <> "" Then
                    ''pilih cbo tipe tujuan
                    If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDebitNote_SupplierAccountCode.Text) = True Then
                        Me.TxtDebitNote_PrimarySupplierAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtDebitNote_SupplierAccountCode.Text))
                        Me.TxtDebitNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimarySupplierAccount.Text))
                        Me.btnOK.Focus()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.BtnDebitNote_FindSupplierAccount.PerformClick()
                    End If
                Else
                    Me.BtnDebitNote_FindSupplierAccount.PerformClick()
                End If
            End If
            Me.TxtDebitNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimarySupplierAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDebitNote_SupplierAccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtDebitNote_SupplierAccountCode.Leave
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDebitNote_SupplierAccountCode.Text <> "" Then
                If ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.TxtDebitNote_SupplierAccountCode.Text) = True Then
                    Me.TxtDebitNote_PrimarySupplierAccount.Text = CStr(ObjAkun.GetPrimary(Me.TxtDebitNote_SupplierAccountCode.Text))
                    Me.TxtDebitNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimarySupplierAccount.Text))
                Else
                    If Me.TxtDebitNote_SupplierAccountCode.Text <> "" Then
                        ij.txtCari.Text = Me.TxtDebitNote_SupplierAccountCode.Text
                    End If

                    ij.ShowDialog()
                    If ij.mPrimaryAkun > 0 Then
                        Me.TxtDebitNote_PrimarySupplierAccount.Text = CStr(ij.mPrimaryAkun)
                        Me.TxtDebitNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimarySupplierAccount.Text))
                    Else
                        Me.TxtDebitNote_SupplierAccountCode.Clear()
                        Me.TxtDebitNote_PrimarySupplierAccount.Text = "-1"
                        Me.btnOK.Focus()
                    End If
                End If
            Else
                Me.TxtDebitNote_SupplierAccountCode.Clear()
                Me.TxtDebitNote_PrimarySupplierAccount.Text = "-1"
            End If
            Me.TxtDebitNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimarySupplierAccount.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnDebitNote_FindExpenseAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDebitNote_FindSupplierAccount.Click
        Dim ij As New FrmMasterAkun_Cari
        ij.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtDebitNote_SupplierAccountCode.Text <> "" Then
                ij.txtCari.Text = Me.TxtDebitNote_SupplierAccountCode.Text
            End If

            ij.ShowDialog()
            If ij.mPrimaryAkun > 0 Then
                Me.TxtDebitNote_PrimarySupplierAccount.Text = CStr(ij.mPrimaryAkun)
                Me.TxtDebitNote_SupplierAccountCode.Text = ObjAkun.GetAccountCode(CDbl(Me.TxtDebitNote_PrimarySupplierAccount.Text))
                Me.TxtDebitNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimarySupplierAccount.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtDebitNote_SupplierAccountCode.Clear()
                Me.TxtDebitNote_PrimarySupplierAccount.Text = "-1"
                Me.TxtDebitNote_SupplierAccountDesc.Text = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.TxtDebitNote_PrimarySupplierAccount.Text)
                Me.btnOK.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockBundling_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockBundling.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockInAdjust_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockInAdjust.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockInPurchase_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockInPurchase.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockInSalesReturn_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockInSalesReturn.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockOutAdjust_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockOutAdjust.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockOutPurchaseReturn_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockOutPurchaseReturn.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockOutSales_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockOutSales.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkInv_UseCustomPrint_StockTransfer_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInv_UseCustomPrint_StockTransfer.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustom_CostFactorCalc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustom_CostFactorCalc.CheckedChanged
        Try
            Me.SettingInventoryBerubah = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
