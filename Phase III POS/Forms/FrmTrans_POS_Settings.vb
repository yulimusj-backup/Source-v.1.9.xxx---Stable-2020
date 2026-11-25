Option Explicit On
Option Strict Off
Imports AccModule, Core_Module

Public Class FrmTrans_POS_Settings
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris

    Dim Buffer As String
    Dim mrOkCancel As Boolean = False
    Dim mrPengaturanSettingPosisiEdit As Boolean = False
    Dim EraseRowIndex_Branch As Integer = 0
    Dim ObjCD As New ClsCekData
    Dim ObjCurr As New clsMasterCurrency
    Dim ObjPayOpt As New ClsMast_PaymentOption
    Dim Obj As New ClsTrans_POS_GeneralSetting
    Dim ObjItem As New ClsTrans_POS_GeneralSetting_Item
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjUser As New ClsUser
    Dim BufferCurr As String = ""

    Public Enum EnumPengaturanAuthorization As Integer
        Cancel_item
        Cancel_Whole_Transaction
        Modify_Transaction
        Modify_Unit_Price
        Modify_Discount_Rate
        Reprint_Receipt
    End Enum

    Private Sub PengaturanTabDataSetting()
        Try
            'Pengaturan control tab data setting 
            If Me.chk_CancelItem_Allow.Checked = True Then
                Me.chk_CancelItem_Autho.Enabled = True
                Me.chk_CancelItem_Autho.Checked = False
            Else
                Me.chk_CancelItem_Autho.Enabled = False
                Me.chk_CancelItem_Autho.Checked = False
            End If
            If Me.chk_CancelWhole_Allow.Checked = True Then
                Me.chk_CancelWhole_Autho.Enabled = True
                Me.chk_CancelWhole_Autho.Checked = False
            Else
                Me.chk_CancelWhole_Autho.Enabled = False
                Me.chk_CancelWhole_Autho.Checked = False
            End If
            If Me.chk_Modify_Allow.Checked = True Then
                Me.chk_Modify_Autho.Enabled = True
                Me.chk_Modify_Autho.Checked = False
            Else
                Me.chk_Modify_Autho.Enabled = False
                Me.chk_Modify_Autho.Checked = False
            End If
            If Me.chk_ModifyUnitPrice_Allow.Checked = True Then
                Me.chk_ModifyUnitPrice_Autho.Enabled = True
                Me.chk_ModifyUnitPrice_Autho.Checked = False
            Else
                Me.chk_ModifyUnitPrice_Autho.Enabled = False
                Me.chk_ModifyUnitPrice_Autho.Checked = False
            End If
            If Me.chk_ModifyDiscount_Allow.Checked = True Then
                Me.chk_ModifyDiscount_Autho.Enabled = True
                Me.chk_ModifyDiscount_Autho.Checked = False
            Else
                Me.chk_ModifyDiscount_Autho.Enabled = False
                Me.chk_ModifyDiscount_Autho.Checked = False
            End If
            If Me.chk_Reprint_Allow.Checked = True Then
                Me.chk_Reprint_Autho.Enabled = True
                Me.chk_Reprint_Autho.Checked = False
            Else
                Me.chk_Reprint_Autho.Enabled = False
                Me.chk_Reprint_Autho.Checked = False
            End If
            'Tambahan Koding desain (16 Oktober 2012)
            If Me.chkAllowExpiry.Checked = False And Me.chkAllowLocation.Checked = False And Me.chkAllowBatchNumber.Checked = False Then
                Me.rboNo.Checked = True
                Me.rboNo.Enabled = False
                Me.rboYes.Enabled = False
            Else
                Me.rboNo.Checked = True
                Me.rboNo.Enabled = True
                Me.rboYes.Enabled = True
            End If
            If Me.rbo_Transf_Stop.Checked = True Then
                Me.gboTransfer.Enabled = False
                Me.cboExpiry.SelectedIndex = 1
                Me.cboBatch.SelectedIndex = 1
            Else
                Me.gboTransfer.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_Settings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim mPrimary As Double = -1
        Try
            Me.TxtCompanyName.Focus()
            Me.mrOkCancel = False
            Me.KosongkanField()
            If Me.chkUseTaxRate.Checked = False Then
                Me.TxtTaxRate.Enabled = False
            End If
            Me.dgvAutorization.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'Cari Primary
            If Obj.GetJmlComapany() > 0 Then
                mPrimary = Obj.GetPrimaryCompany
            Else
                mPrimary = -1
            End If
            'Cari detail
            If mPrimary <> -1 Then
                If Me.Obj.CariDetail(mPrimary, Me.dgvAutorization, Me.ObjCD, Me) = True Then
                    'Edit data
                    Me.mrPengaturanSettingPosisiEdit = True
                    Me.PengaturanTabDataSetting()
                    Me.SalinKeField()
                    If Me.chkAllowRounding.Checked = False Then
                        Me.NonAktifkanDetailRounding()
                        Me.KosongkanFieldRounding()
                    End If
                    Me.mrPengaturanSettingPosisiEdit = False
                End If
            Else
                'New data
                Me.mrPengaturanSettingPosisiEdit = False
                Me.PengaturanTabDataSetting()
                Me.chkAllowRounding.Checked = False
                Me.NonAktifkanDetailRounding()
                Me.KosongkanFieldRounding()
                Me.chkAllowLocation.Checked = False
                Me.chkAllowBatchNumber.Checked = False
                Me.chkAllowExpiry.Checked = False
                Me.rboYes.Checked = False
                Me.rboNo.Checked = True
                Me.rbo_Transf_Stop.Checked = True
                Me.rbo_Transf_Automatic.Checked = False
                Me.gboTransfer.Enabled = False
                Me.cboExpiry.SelectedIndex = 1
                Me.cboBatch.SelectedIndex = 1
            End If
            If Me.TxtPrimaryPayment.Text = "-1" Then
                Me.TxtPaymentOption.Text = "Cash"
            End If
            Me.BufferCurr = Me.TxtPaymentOption.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.Obj
                'Header
                .CompanyName = Me.TxtCompanyName.Text
                .CompanyLogo = Me.TxtCompanyLogoFiles.Text
                If Me.chkShowCompanyName.Checked = True Then
                    .ShowCompanyName = 1
                Else
                    .ShowCompanyName = 0
                End If
                If Me.chkShowLogo.Checked = True Then
                    .ShowCompanyLogo = 1
                Else
                    .ShowCompanyLogo = 0
                End If
                .PrimaryCurr = Me.TxtPrimaryCurrency.Text
                .PrimaryPayment = Me.TxtPrimaryPayment.Text
                If Me.chk_Allow_CreditCard.Checked = True Then
                    .AllowCreditCardPayment = 1
                Else
                    .AllowCreditCardPayment = 0
                End If
                If Me.chk_Allow_DebitCard.Checked = True Then
                    .AllowDebitCardPayment = 1
                Else
                    .AllowDebitCardPayment = 0
                End If
                If Me.chk_AlowCashWithdrawal.Checked = True Then
                    .AllowCashPayment = 1
                Else
                    .AllowCashPayment = 0
                End If
                If Me.chk_Allow_VouvherPayment.Checked = True Then
                    .AllowVoucherPayment = 1
                Else
                    .AllowVoucherPayment = 0
                End If
                If Me.chkUseTaxRate.Checked = True Then
                    .UseTaxRate = 1
                    .PersenTaxRate = ObjCD.DigitGrouping(Me.TxtTaxRate.Text, 2)
                Else
                    .UseTaxRate = 0
                    .PersenTaxRate = 0
                End If
                'Rounding 
                If Me.chkAllowRounding.Checked = True Then
                    .AllowRounding = 1
                Else
                    .AllowRounding = 0
                End If
                .RoundingTipe = Me.cboRoundingTipe.SelectedIndex
                .RoundingValue = Me.cboRoundingValue.SelectedIndex
                .RoundingOn = Me.cboRoundingOn.SelectedIndex
                .PrimaryAcc = Me.TxtRounding_PrimaryAccount.Text
                'Detail setting
                If Me.chk_CancelItem_Allow.Checked = True Then
                    .AllowToCancel = 1
                Else
                    .AllowToCancel = 0
                End If
                If Me.chk_CancelItem_Autho.Checked = True Then
                    .UseAuthorToCancel = 1
                Else
                    .UseAuthorToCancel = 0
                End If
                If Me.chk_CancelWhole_Allow.Checked = True Then
                    .AllowUserToCancelWholeTrans = 1
                Else
                    .AllowUserToCancelWholeTrans = 0
                End If
                If Me.chk_CancelWhole_Autho.Checked = True Then
                    .UseAuthorToCancelWholeTrans = 1
                Else
                    .UseAuthorToCancelWholeTrans = 0
                End If
                If Me.chk_Modify_Allow.Checked = True Then
                    .AllowUserToModifyTrans = 1
                Else
                    .AllowUserToModifyTrans = 0
                End If
                If Me.chk_Modify_Autho.Checked = True Then
                    .UseAuthorToModifyTrans = 1
                Else
                    .UseAuthorToModifyTrans = 0
                End If
                If Me.chk_ModifyUnitPrice_Allow.Checked = True Then
                    .AllowUserToModifyUnitPrice = 1
                Else
                    .AllowUserToModifyUnitPrice = 0
                End If
                If Me.chk_ModifyUnitPrice_Autho.Checked = True Then
                    .UseAuthorToModifyUnitPrice = 1
                Else
                    .UseAuthorToModifyUnitPrice = 0
                End If
                If Me.chk_ModifyDiscount_Allow.Checked = True Then
                    .AllowUserToModifyDiscountRate = 1
                Else
                    .AllowUserToModifyDiscountRate = 0
                End If
                If Me.chk_ModifyDiscount_Autho.Checked = True Then
                    .UseAuthorToModifyDiscountRate = 1
                Else
                    .UseAuthorToModifyDiscountRate = 0
                End If
                If Me.chk_Reprint_Allow.Checked = True Then
                    .AllowUserToReprintReceipt = 1
                Else
                    .AllowUserToReprintReceipt = 0
                End If
                If Me.chk_Reprint_Autho.Checked = True Then
                    .UseAuthorToReprintReceipt = 1
                Else
                    .UseAuthorToReprintReceipt = 0
                End If
                If Me.rbo_Transf_Stop.Checked = True Then
                    .AutomaticTransfer = ClsTrans_POS_GeneralSetting.EnumAutomaticTransfer.Stop_Transfer
                Else
                    .AutomaticTransfer = ClsTrans_POS_GeneralSetting.EnumAutomaticTransfer.Automatically
                End If
                If Me.cboExpiry.SelectedIndex = 1 Then
                    .ExpiryASC = ClsTrans_POS_GeneralSetting.EnumExpiryASC.Descending
                Else
                    .ExpiryASC = ClsTrans_POS_GeneralSetting.EnumExpiryASC.Acscending
                End If
                If Me.cboBatch.SelectedIndex = 1 Then
                    .BatchASC = ClsTrans_POS_GeneralSetting.EnumExpiryASC.Descending
                Else
                    .BatchASC = ClsTrans_POS_GeneralSetting.EnumExpiryASC.Acscending
                End If
                If Me.chkAllowLocation.Checked = True Then
                    .AllowLocationControl = 1
                Else
                    .AllowLocationControl = 0
                End If
                If Me.chkAllowBatchNumber.Checked = True Then
                    .AllowBatchNumberControl = 1
                Else
                    .AllowBatchNumberControl = 0
                End If
                If Me.chkAllowExpiry.Checked = True Then
                    .AllowExpiryControl = 1
                Else
                    .AllowExpiryControl = 0
                End If
                If Me.rboNo.Checked = True Then
                    .InputLocationOnCashierForm = ClsTrans_POS_GeneralSetting.EnumInputLocationOnCashierForm.No
                Else
                    .InputLocationOnCashierForm = ClsTrans_POS_GeneralSetting.EnumInputLocationOnCashierForm.Yes
                End If
                If Me.ChkUsePaymentDiscount.Checked = True Then
                    .UsePaymentDiskon = 1
                Else
                    .UsePaymentDiskon = 0
                End If
                .DefPaymentDiskon = Me.TxtDefPaymentDisc.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.Obj
                Me.TxtPrimary.Text = .Primary
                Me.TxtCompanyName.Text = .CompanyName
                Me.TxtCompanyLogoFiles.Text = .CompanyLogo
                If String.IsNullOrEmpty(Me.TxtCompanyLogoFiles.Text) = False Then
                    Me.PictureBox_Logo.Load(Me.TxtCompanyLogoFiles.Text)
                End If
                If .ShowCompanyName = 1 Then
                    Me.chkShowCompanyName.Checked = True
                Else
                    Me.chkShowCompanyName.Checked = False
                End If
                If .ShowCompanyLogo = 1 Then
                    Me.chkShowLogo.Checked = True
                Else
                    Me.chkShowLogo.Checked = False
                End If
                Me.TxtPrimaryCurrency.Text = .PrimaryCurr
                Me.txtCurrencyCode.Text = ObjCurr.GetCurrencyCode(.PrimaryCurr)
                Me.TxtPrimaryPayment.Text = .PrimaryPayment
                Me.TxtPaymentOption.Text = ObjPayOpt.GetCode(.PrimaryPayment)
                If .AllowCreditCardPayment = 1 Then
                    Me.chk_Allow_CreditCard.Checked = True
                Else
                    Me.chk_Allow_CreditCard.Checked = False
                End If
                If .AllowDebitCardPayment = 1 Then
                    Me.chk_Allow_DebitCard.Checked = True
                Else
                    Me.chk_Allow_DebitCard.Checked = False
                End If
                If .AllowCashPayment = 1 Then
                    Me.chk_AlowCashWithdrawal.Checked = True
                Else
                    Me.chk_AlowCashWithdrawal.Checked = False
                End If
                If .AllowVoucherPayment = 1 Then
                    Me.chk_Allow_VouvherPayment.Checked = True
                Else
                    Me.chk_Allow_VouvherPayment.Checked = False
                End If
                If .UseTaxRate = 1 Then
                    Me.chkUseTaxRate.Checked = True
                    Me.TxtTaxRate.Text = ObjCD.DigitGrouping(.PersenTaxRate, 2)
                Else
                    Me.chkUseTaxRate.Checked = False
                    Me.TxtTaxRate.Text = "0.00"
                End If
                If .AllowRounding = 1 Then
                    Me.chkAllowRounding.Checked = True
                Else
                    Me.chkAllowRounding.Checked = False
                End If
                Me.cboRoundingTipe.SelectedIndex = .RoundingTipe
                Me.cboRoundingValue.SelectedIndex = .RoundingValue
                Me.cboRoundingOn.SelectedIndex = .RoundingOn
                Me.TxtRounding_PrimaryAccount.Text = .PrimaryAcc
                Me.TxtRounding_AccountCode.Text = ObjAkun.GetAccountCode(.PrimaryAcc)
                If .AllowToCancel = 1 Then
                    Me.chk_CancelItem_Allow.Checked = True
                Else
                    Me.chk_CancelItem_Allow.Checked = False
                End If
                If .UseAuthorToCancel = 1 Then
                    Me.chk_CancelItem_Autho.Checked = True
                Else
                    Me.chk_CancelItem_Autho.Checked = False
                End If
                If .AllowUserToCancelWholeTrans = 1 Then
                    Me.chk_CancelWhole_Allow.Checked = True
                Else
                    Me.chk_CancelWhole_Allow.Checked = False
                End If
                If .UseAuthorToCancelWholeTrans = 1 Then
                    Me.chk_CancelWhole_Autho.Checked = True
                Else
                    Me.chk_CancelWhole_Autho.Checked = False
                End If
                If .AllowUserToModifyTrans = 1 Then
                    Me.chk_Modify_Allow.Checked = True
                Else
                    Me.chk_Modify_Allow.Checked = False
                End If
                If .UseAuthorToModifyTrans = 1 Then
                    Me.chk_Modify_Autho.Checked = True
                Else
                    Me.chk_Modify_Autho.Checked = False
                End If
                If .AllowUserToModifyUnitPrice = 1 Then
                    Me.chk_ModifyUnitPrice_Allow.Checked = True
                Else
                    Me.chk_ModifyUnitPrice_Allow.Checked = False
                End If
                If .UseAuthorToModifyUnitPrice = 1 Then
                    Me.chk_ModifyUnitPrice_Autho.Checked = True
                Else
                    Me.chk_ModifyUnitPrice_Autho.Checked = False
                End If
                If .AllowUserToModifyDiscountRate = 1 Then
                    Me.chk_ModifyDiscount_Allow.Checked = True
                Else
                    Me.chk_ModifyDiscount_Allow.Checked = False
                End If
                If .UseAuthorToModifyDiscountRate = 1 Then
                    Me.chk_ModifyDiscount_Autho.Checked = True
                Else
                    Me.chk_ModifyDiscount_Autho.Checked = False
                End If
                If .AllowUserToReprintReceipt = 1 Then
                    Me.chk_Reprint_Allow.Checked = True
                Else
                    Me.chk_Reprint_Allow.Checked = False
                End If
                If .UseAuthorToReprintReceipt = 1 Then
                    Me.chk_Reprint_Autho.Checked = True
                Else
                    Me.chk_Reprint_Autho.Checked = False
                End If
                If .AutomaticTransfer = ClsTrans_POS_GeneralSetting.EnumAutomaticTransfer.Stop_Transfer Then
                    Me.rbo_Transf_Stop.Checked = True
                Else
                    Me.rbo_Transf_Automatic.Checked = True
                End If
                If .ExpiryASC = ClsTrans_POS_GeneralSetting.EnumExpiryASC.Descending Then
                    Me.cboExpiry.SelectedIndex = 1
                Else
                    Me.cboExpiry.SelectedIndex = 0
                End If
                If .BatchASC = ClsTrans_POS_GeneralSetting.EnumExpiryASC.Descending Then
                    Me.cboBatch.SelectedIndex = 1
                Else
                    Me.cboBatch.SelectedIndex = 0
                End If
                If .AllowLocationControl = 1 Then
                    Me.chkAllowLocation.Checked = True
                Else
                    Me.chkAllowLocation.Checked = False
                End If
                If .AllowBatchNumberControl = 1 Then
                    Me.chkAllowBatchNumber.Checked = True
                Else
                    Me.chkAllowBatchNumber.Checked = False
                End If
                If .AllowExpiryControl = 1 Then
                    Me.chkAllowExpiry.Checked = True
                Else
                    Me.chkAllowExpiry.Checked = False
                End If
                If .InputLocationOnCashierForm = ClsTrans_POS_GeneralSetting.EnumInputLocationOnCashierForm.No Then
                    Me.rboNo.Checked = True
                Else
                    Me.rboYes.Checked = True
                End If
                Me.ChkUsePaymentDiscount.Checked = .UsePaymentDiskon
                Me.TxtDefPaymentDisc.Text = .DefPaymentDiskon

                'Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                'Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                'Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                'Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.chkShowCompanyName.Checked = True
            Me.chkShowLogo.Checked = True
            Me.TxtCompanyName.Text = ""
            Me.TxtPrimary.Text = "-1"
            Me.TxtCompanyLogoFiles.Text = ""
            Me.txtCurrencyCode.Text = ""
            Me.TxtPrimaryCurrency.Text = "-1"
            Me.TxtPaymentOption.Text = ""
            Me.TxtPrimaryPayment.Text = "-1"
            Me.TxtTaxRate.Text = "0"
            Me.chk_Allow_CreditCard.Checked = False
            Me.chk_Allow_DebitCard.Checked = False
            Me.chk_AlowCashWithdrawal.Checked = False
            Me.chk_Allow_VouvherPayment.Checked = False
            Me.chkUseTaxRate.Checked = False
            Me.TxtTaxRate.Text = "0.00"
            Me.dgvAutorization.Rows.Clear()

            Me.chk_CancelItem_Allow.Checked = False
            Me.chk_CancelItem_Autho.Checked = False

            Me.chk_CancelWhole_Allow.Checked = False
            Me.chk_CancelWhole_Autho.Checked = False

            Me.chk_Modify_Allow.Checked = False
            Me.chk_Modify_Autho.Checked = False

            Me.chk_ModifyUnitPrice_Allow.Checked = False
            Me.chk_ModifyUnitPrice_Autho.Checked = False

            Me.chk_ModifyDiscount_Allow.Checked = False
            Me.chk_ModifyDiscount_Autho.Checked = False

            Me.chk_Reprint_Allow.Checked = False
            Me.chk_Reprint_Autho.Checked = False

            Me.KosongkanFieldRounding()
            Me.PictureBox_Logo.Image = Nothing

            Me.rbo_Transf_Stop.Checked = True
            Me.rbo_Transf_Automatic.Checked = False
            Me.cboExpiry.SelectedIndex = 1
            Me.cboBatch.SelectedIndex = 1

            Me.chkAllowLocation.Checked = False
            Me.chkAllowBatchNumber.Checked = False
            Me.chkAllowExpiry.Checked = False
            Me.rboYes.Checked = False
            Me.rboNo.Checked = True

            Me.ChkUsePaymentDiscount.Checked = False
            Me.TxtDefPaymentDisc.Text = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanFieldRounding()
        Try
            Me.cboRoundingTipe.SelectedIndex = 0
            Me.cboRoundingValue.SelectedIndex = 6
            Me.cboRoundingOn.SelectedIndex = 0
            Me.TxtRounding_AccountCode.Text = ""
            Me.TxtRounding_PrimaryAccount.Text = "-1"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetailRounding()
        Try
            Me.cboRoundingTipe.Enabled = True
            Me.cboRoundingValue.Enabled = True
            Me.cboRoundingOn.Enabled = True
            Me.TxtRounding_AccountCode.Enabled = True
            Me.btnRounding_Find.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetailRounding()
        Try
            Me.cboRoundingTipe.Enabled = False
            Me.cboRoundingValue.Enabled = False
            Me.cboRoundingOn.Enabled = False
            Me.TxtRounding_AccountCode.Enabled = False
            Me.btnRounding_Find.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllowRounding_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllowRounding.CheckedChanged
        Try
            If Me.chkAllowRounding.Checked = True Then
                Me.cboRoundingTipe.Enabled = True
                Me.cboRoundingTipe.TabStop = True
                Me.cboRoundingValue.Enabled = True
                Me.cboRoundingValue.TabStop = True
                Me.cboRoundingOn.Enabled = True
                Me.cboRoundingOn.TabStop = True
                Me.TxtRounding_AccountCode.Enabled = True
                Me.TxtRounding_AccountCode.TabStop = True
                Me.btnRounding_Find.Enabled = True
            Else
                Me.cboRoundingTipe.Enabled = False
                Me.cboRoundingTipe.TabStop = False
                Me.cboRoundingValue.Enabled = False
                Me.cboRoundingValue.TabStop = False
                Me.cboRoundingOn.Enabled = False
                Me.cboRoundingOn.TabStop = False
                Me.TxtRounding_AccountCode.Enabled = False
                Me.TxtRounding_AccountCode.TabStop = False
                Me.btnRounding_Find.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UncheckedAllDGVAuthor(ByVal DGV As DataGridView, ByVal Tipe As EnumPengaturanAuthorization)
        Dim Row As DataGridViewRow
        Try
            If mrPengaturanSettingPosisiEdit = False Then
                DGV.AllowUserToAddRows = False
                If DGV.Rows.Count > 0 Then
                    If Tipe = EnumPengaturanAuthorization.Cancel_item Then
                        For Each Row In DGV.Rows
                            If Row.Cells(Me.Col_Autho_AllowDeleteItem.Index).Value = True Then
                                Row.Cells(Me.Col_Autho_AllowDeleteItem.Index).Value = False
                            End If
                        Next
                    End If
                    If Tipe = EnumPengaturanAuthorization.Cancel_Whole_Transaction Then
                        For Each Row In DGV.Rows
                            If Row.Cells(Me.Col_Autho_WholeTransDelete.Index).Value = True Then
                                Row.Cells(Me.Col_Autho_WholeTransDelete.Index).Value = False
                            End If
                        Next
                    End If
                    If Tipe = EnumPengaturanAuthorization.Modify_Transaction Then
                        For Each Row In DGV.Rows
                            If Row.Cells(Me.Col_Autho_AllowModify.Index).Value = True Then
                                Row.Cells(Me.Col_Autho_AllowModify.Index).Value = False
                            End If
                        Next
                    End If
                    If Tipe = EnumPengaturanAuthorization.Modify_Unit_Price Then
                        For Each Row In DGV.Rows
                            If Row.Cells(Me.Col_Autho_ModifyUnitPrice.Index).Value = True Then
                                Row.Cells(Me.Col_Autho_ModifyUnitPrice.Index).Value = False
                            End If
                        Next
                    End If
                    If Tipe = EnumPengaturanAuthorization.Modify_Discount_Rate Then
                        For Each Row In DGV.Rows
                            If Row.Cells(Me.Col_Autho_ModifyDiscount.Index).Value = True Then
                                Row.Cells(Me.Col_Autho_ModifyDiscount.Index).Value = False
                            End If
                        Next
                    End If
                    If Tipe = EnumPengaturanAuthorization.Reprint_Receipt Then
                        For Each Row In DGV.Rows
                            If Row.Cells(Me.Col_Autho_AllowReprint.Index).Value = True Then
                                Row.Cells(Me.Col_Autho_AllowReprint.Index).Value = False
                            End If
                        Next
                    End If
                End If
                DGV.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Control Chk Allow dan use
    Private Sub chk_CancelItem_Allow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_CancelItem_Allow.CheckedChanged
        Try
            If Me.chk_CancelItem_Allow.Checked = True Then
                Me.chk_CancelItem_Autho.Enabled = True
                Me.chk_CancelItem_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Cancel_item)
            Else
                Me.chk_CancelItem_Autho.Enabled = False
                Me.chk_CancelItem_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Cancel_item)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_CancelWhole_Allow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_CancelWhole_Allow.CheckedChanged
        Try
            If Me.chk_CancelWhole_Allow.Checked = True Then
                Me.chk_CancelWhole_Autho.Enabled = True
                Me.chk_CancelWhole_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Cancel_Whole_Transaction)
            Else
                Me.chk_CancelWhole_Autho.Enabled = False
                Me.chk_CancelWhole_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Cancel_Whole_Transaction)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_Modify_Allow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Modify_Allow.CheckedChanged
        Try
            If Me.chk_Modify_Allow.Checked = True Then
                Me.chk_Modify_Autho.Enabled = True
                Me.chk_Modify_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Transaction)
            Else
                Me.chk_Modify_Autho.Enabled = False
                Me.chk_Modify_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Transaction)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_ModifyUnitPrice_Allow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_ModifyUnitPrice_Allow.CheckedChanged
        Try
            If Me.chk_ModifyUnitPrice_Allow.Checked = True Then
                Me.chk_ModifyUnitPrice_Autho.Enabled = True
                Me.chk_ModifyUnitPrice_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Unit_Price)
            Else
                Me.chk_ModifyUnitPrice_Autho.Enabled = False
                Me.chk_ModifyUnitPrice_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Unit_Price)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_ModifyDiscount_Allow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_ModifyDiscount_Allow.CheckedChanged
        Try
            If Me.chk_ModifyDiscount_Allow.Checked = True Then
                Me.chk_ModifyDiscount_Autho.Enabled = True
                Me.chk_ModifyDiscount_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Discount_Rate)
            Else
                Me.chk_ModifyDiscount_Autho.Enabled = False
                Me.chk_ModifyDiscount_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Discount_Rate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_Reprint_Allow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Reprint_Allow.CheckedChanged
        Try
            If Me.chk_Reprint_Allow.Checked = True Then
                Me.chk_Reprint_Autho.Enabled = True
                Me.chk_Reprint_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Reprint_Receipt)
            Else
                Me.chk_Reprint_Autho.Enabled = False
                Me.chk_Reprint_Autho.Checked = False
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Reprint_Receipt)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxRate.Enter
        Try
            ObjCD.Number_Enter(Me.TxtTaxRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxRate.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtTaxRate, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtTaxRate.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtTaxRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtTaxRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTaxRate.Leave
        Try
            ObjCD.Number_Leave(Me.TxtTaxRate, sender, e, 2, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Back Color
    Private Sub TxtCompanyName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCompanyName.GotFocus
        Try
            If Me.TxtCompanyName.Enabled = True Then
                If Me.TxtCompanyName.ReadOnly = False Then
                    Me.TxtCompanyName.BackColor = Me.BackColorAll
                Else
                    Me.TxtCompanyName.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCompanyName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCompanyName.LostFocus
        Try
            If Me.TxtCompanyName.Enabled = True Then
                If Me.TxtCompanyName.ReadOnly = False Then
                    Me.TxtCompanyName.BackColor = Me.TxtPrimary.BackColor
                Else
                    Me.TxtCompanyName.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrency_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        Try
            If Me.txtCurrencyCode.Enabled = True Then
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

    Private Sub TxtCurrency_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        Try
            If Me.txtCurrencyCode.Enabled = True Then
                If Me.txtCurrencyCode.ReadOnly = False Then
                    Me.txtCurrencyCode.BackColor = Me.TxtPrimary.BackColor
                Else
                    Me.txtCurrencyCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentOption_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentOption.GotFocus
        Try
            If Me.TxtPaymentOption.Enabled = True Then
                If Me.TxtPaymentOption.ReadOnly = False Then
                    Me.TxtPaymentOption.BackColor = Me.BackColorAll
                Else
                    Me.TxtPaymentOption.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentOption_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentOption.LostFocus
        Try
            If Me.TxtPaymentOption.Enabled = True Then
                If Me.TxtPaymentOption.ReadOnly = False Then
                    Me.TxtPaymentOption.BackColor = Me.TxtPrimary.BackColor
                Else
                    Me.TxtPaymentOption.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRounding_AccountCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRounding_AccountCode.GotFocus
        Try
            If Me.TxtRounding_AccountCode.Enabled = True Then
                If Me.TxtRounding_AccountCode.ReadOnly = False Then
                    Me.TxtRounding_AccountCode.BackColor = Me.BackColorAll
                Else
                    Me.TxtRounding_AccountCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRounding_AccountCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRounding_AccountCode.LostFocus
        Try
            If Me.TxtRounding_AccountCode.Enabled = True Then
                If Me.TxtRounding_AccountCode.ReadOnly = False Then
                    Me.TxtRounding_AccountCode.BackColor = Me.TxtPrimary.BackColor
                Else
                    Me.TxtRounding_AccountCode.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseTaxRate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseTaxRate.CheckedChanged
        Try
            If Me.chkUseTaxRate.Checked = True Then
                Me.TxtTaxRate.Enabled = True
                Me.TxtTaxRate.Text = "0.00"
            Else
                Me.TxtTaxRate.Enabled = False
                Me.TxtTaxRate.Text = "0.00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            '1. cek data yang harus diisi
            KoleksiNamaData.Add("Company Name")
            KoleksiTextBox.Add(Me.TxtCompanyName)

            If Me.ObjCD.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. Default Currency
            If Me.TxtPrimaryCurrency.Text = "-1" Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input default currency.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan default currency.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.txtCurrencyCode.Focus()
                Return False
            End If

            '3 Payment Option
            If String.IsNullOrEmpty(Me.TxtPaymentOption.Text) = True Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input payment option.", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan payment option.", MsgBoxStyle.Information, "Informasi")
                End If
                Me.TxtPaymentOption.Focus()
                Return False
            End If

            '4. Allow Rounding 
            If Me.chkAllowRounding.Checked = True Then
                If Me.TxtRounding_PrimaryAccount.Text = "-1" Then
                    If Bahasa = EnumBahasa.Inggris Then
                        MsgBox("Please input account code.", MsgBoxStyle.Information, "Information")
                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                        MsgBox("Masukan account code.", MsgBoxStyle.Information, "Informasi")
                    End If
                    Me.TabControl2.SelectedIndex = 2
                    Return False
                End If
            End If

            'Item tidak boleh kosong 
            Me.dgvAutorization.AllowUserToAddRows = False
            If Me.dgvAutorization.Rows.Count <= 0 Then
                If Bahasa = EnumBahasa.Inggris Then
                    MsgBox("Please input item data", MsgBoxStyle.Information, "Information")
                ElseIf Bahasa = EnumBahasa.Indonesia Then
                    MsgBox("Masukan data item", MsgBoxStyle.Information, "Informasi")
                End If
                Me.dgvAutorization.AllowUserToAddRows = True
                Me.TabControl2.SelectedIndex = 1
                Return False
            End If
            Me.dgvAutorization.AllowUserToAddRows = True

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Obj.GetJmlComapany() <= 0 Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.mrOkCancel = True
                            Me.Close()
                            'Me.SalinKeField()
                            'Me.NonAktifkanDetail()
                            'Me.btnFind.Focus()
                        End If
                    Else
                        Me.mrOkCancel = False
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.Bahasa, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.mrOkCancel = True
                            Me.Close()
                            'Me.SalinKeField()
                            'Me.NonAktifkanDetail()
                            'Me.btnEdit.Focus()
                        End If
                    Else
                        Me.mrOkCancel = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtCompanyName.Enabled = True Then
                If Me.mrOkCancel = False Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinKeObjek()
                            If Obj.GetJmlComapany() <= 0 Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.Bahasa, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    'Me.SalinKeField()
                                    'Me.NonAktifkanDetail()
                                    'Me.btnFind.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            Else
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.Bahasa, Me) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    'Me.SalinKeField()
                                    'Me.NonAktifkanDetail()
                                    'Me.btnEdit.Focus()
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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Me.TxtCompanyName.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Obj.GetJmlComapany() <= 0 Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanBaru(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.Bahasa, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.mrOkCancel = True
                                Me.Close()
                                'Me.SalinKeField()
                                'Me.NonAktifkanDetail()
                                'Me.btnFind.Focus()
                            Else
                                'e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.Obj.SimpanPerubahan(Me.UserCode, Me.ObjCD, Me.TxtPrimary.Text, Me.Bahasa, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.mrOkCancel = True
                                Me.Close()
                                'Me.SalinKeField()
                                'Me.NonAktifkanDetail()
                                'Me.btnEdit.Focus()
                            Else
                                'e.Cancel = True
                            End If
                        End If
                    Else
                        'e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    mrOkCancel = False
                    'e.Cancel = True
                ElseIf hasil = MsgBoxResult.No Then
                    Me.mrOkCancel = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItemInputBranch(ByVal vUserCode As String) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvAutorization.Rows
                If rowDGV.Cells(Me.ColUserKode.Index).Value = vUserCode Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvAutorization_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAutorization.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColUserKode.Index Or _
            e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                If Me.chk_CancelItem_Autho.Checked = True And String.IsNullOrEmpty(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.Col_Autho_AllowDeleteItem.Index Then
                        Me.Buffer = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If e.ColumnIndex = Me.Col_Autho_AllowDeleteItem.Index Then
                        e.Cancel = True
                    End If
                End If
                If Me.chk_CancelWhole_Autho.Checked = True And String.IsNullOrEmpty(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.Col_Autho_WholeTransDelete.Index Then
                        Me.Buffer = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If e.ColumnIndex = Me.Col_Autho_WholeTransDelete.Index Then
                        e.Cancel = True
                    End If
                End If
                If Me.chk_Modify_Autho.Checked = True And String.IsNullOrEmpty(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.Col_Autho_AllowModify.Index Then
                        Me.Buffer = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If e.ColumnIndex = Me.Col_Autho_AllowModify.Index Then
                        e.Cancel = True
                    End If
                End If
                If Me.chk_ModifyUnitPrice_Autho.Checked = True And String.IsNullOrEmpty(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.Col_Autho_ModifyUnitPrice.Index Then
                        Me.Buffer = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If e.ColumnIndex = Me.Col_Autho_ModifyUnitPrice.Index Then
                        e.Cancel = True
                    End If
                End If
                If Me.chk_ModifyDiscount_Autho.Checked = True And String.IsNullOrEmpty(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.Col_Autho_ModifyDiscount.Index Then
                        Me.Buffer = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If e.ColumnIndex = Me.Col_Autho_ModifyDiscount.Index Then
                        e.Cancel = True
                    End If
                End If
                If Me.chk_Reprint_Autho.Checked = True And String.IsNullOrEmpty(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.Col_Autho_AllowReprint.Index Then
                        Me.Buffer = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    If e.ColumnIndex = Me.Col_Autho_AllowReprint.Index Then
                        e.Cancel = True
                    End If
                End If
                If e.ColumnIndex = Me.ColUserName.Index Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAutorization_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAutorization.CellEndEdit
        Dim Row As DataGridViewRow = Me.dgvAutorization.Rows(e.RowIndex)
        Dim Index As Integer = -1
        Try
            'pastikan ada isinya
            With Row
                If IsNothing(Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    'Cek string keterangan
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If ObjCD.CekDataKosong(Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            If Bahasa = EnumBahasa.Inggris Then
                                MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                            ElseIf Bahasa = EnumBahasa.Indonesia Then
                                MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                            End If
                            Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    ''Cell kode 
                    If e.ColumnIndex = Me.ColUserKode.Index Then
                        Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            If ObjUser.ApakahSudahAda(Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Index = Me.GetIndexItemInputBranch(Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value)
                            Else
                                Index = -1
                            End If
                            If Index < 0 Then
                                ''Me.dgvAutorization.Item(Me.ColBranch_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvAutorization.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> Index Then
                                    If Bahasa = EnumBahasa.Inggris Then
                                        MsgBox("You cannot input the same user code", MsgBoxStyle.Information, "Information")
                                    ElseIf Bahasa = EnumBahasa.Indonesia Then
                                        MsgBox("Anda tidak dapat menginput user code yang sama", MsgBoxStyle.Information, "Informasi")
                                    End If
                                    Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex_Branch = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvAutorization)
                                    Exit Sub
                                End If
                            End If
                        Else
                            '2. kode akun yang diketik tidak ada --> tampilkan form cari
                            Dim ij As New FrmCariMasterUser
                            ij.TxtCari.Text = Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                                Index = Me.GetIndexItemInputBranch(ij.Kodeuser)
                                If Index < 0 Then
                                    Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.dgvAutorization.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    ''Me.dgvAutorization.Item(Me.ColBranch_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                Else
                                    If e.RowIndex <> Index Then
                                        If Bahasa = EnumBahasa.Inggris Then
                                            MsgBox("You cannot input the same user code", MsgBoxStyle.Information, "Information")
                                        ElseIf Bahasa = EnumBahasa.Indonesia Then
                                            MsgBox("Anda tidak dapat menginput user code yang sama", MsgBoxStyle.Information, "Informasi")
                                        End If
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex_Branch = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvAutorization)
                                        Exit Sub
                                    Else
                                        Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.dgvAutorization.Item(Me.ColUserName.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        ''Me.dgvAutorization.Item(Me.ColBranch_Primary.Index, e.RowIndex).Value = ij.mPrimary
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If Me.dgvAutorization.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Branch = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvAutorization)
                    End If
                    'Hapus Baris Jika Code No DownGrade Kosong
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvAutorization.Item(Me.ColUserKode.Index, e.RowIndex).Value) = True Then
                            If Me.dgvAutorization.Rows.Count - 1 <> e.RowIndex Then
                                Me.EraseRowIndex_Branch = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        If e.ColumnIndex = Me.ColUserKode.Index Then
                            If String.IsNullOrEmpty(Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If Me.dgvAutorization.Rows.Count - 1 <> e.RowIndex Then
                                    Me.EraseRowIndex_Branch = e.RowIndex
                                    Me.BackgroundWorker2.RunWorkerAsync()
                                End If
                            End If
                        Else
                            Me.dgvAutorization.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAutorization_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvAutorization.UserDeletedRow
        Try
            ObjCD.SetNumber(Me.dgvAutorization, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.dgvAutorization.Rows.RemoveAt(Me.EraseRowIndex_Branch)
            Me.ObjCD.SetNumber(Me.dgvAutorization)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Fitur Picture
    Private Sub BtnFindLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLogo.Click
        Try
            If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TxtCompanyLogoFiles.Text = Me.OpenFileDialog1.FileName
                Me.PictureBox_Logo.Load(Me.TxtCompanyLogoFiles.Text)
                Me.txtCurrencyCode.Focus()
            Else
                Me.PictureBox_Logo.Image = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClear.Click
        Try
            Me.PictureBox_Logo.Image = Nothing
            Me.TxtCompanyLogoFiles.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency
        Dim Obj As New clsMasterCurrency
        Try
            i.BackColorAll = Me.BackColorAll
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.TxtPrimaryCurrency.Text = Obj.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.TxtPaymentOption.Focus()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = Obj.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                        Me.TxtPaymentOption.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
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

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim i As New FrmCariMasterCurrency
        Dim obj As New clsMasterCurrency

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtCurrencyCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.TxtPaymentOption.Focus()
                    Me.TxtPrimaryCurrency.Text = obj.GetPrimary(Me.txtCurrencyCode.Text)
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = obj.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                        Me.TxtPaymentOption.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCurrency.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        Dim obj As New clsMasterCurrency
        Try
            i.BackColorAll = Me.BackColorAll
            'If obj.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
            'Me.TxtPaymentOption.Focus()
            'Me.TxtPrimaryCurrency.Text = obj.GetPrimary(Me.txtCurrencyCode.Text)
            'Else
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCurrencyCode.Text <> "" Then
                i.txtCari.Text = Me.txtCurrencyCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCurrency > 0 Then
                Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                Me.txtCurrencyCode.Text = obj.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                Me.TxtPaymentOption.Focus()
            Else
                Me.txtCurrencyCode.Clear()
                Me.txtCurrencyCode.Focus()
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtRounding_AccountCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtRounding_AccountCode.KeyDown
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        Try
            If e.KeyValue = 13 Then
                If Me.TxtRounding_AccountCode.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtRounding_AccountCode.Text) = True Then
                        Me.TxtRounding_PrimaryAccount.Text = objCode.GetPrimary(Me.TxtRounding_AccountCode.Text)
                        Me.TxtRounding_AccountCode.Text = objCode.GetAccountCode(Me.TxtRounding_PrimaryAccount.Text)
                        Me.btnOK.Focus()
                    Else
                        Me.btnRounding_Find.PerformClick()
                    End If
                Else
                    Me.btnRounding_Find.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtRounding_AccountCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtRounding_AccountCode.Leave
        Dim objCode As New ClsMasterAkun
        Dim i As New FrmMasterAkun_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtRounding_AccountCode.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtRounding_AccountCode.Text) = True Then
                    Me.TxtRounding_PrimaryAccount.Text = objCode.GetPrimary(Me.TxtRounding_AccountCode.Text)
                    Me.TxtRounding_AccountCode.Text = objCode.GetAccountCode(Me.TxtRounding_PrimaryAccount.Text)
                Else
                    If Me.TxtRounding_AccountCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtRounding_AccountCode.Text
                    End If

                    i.ShowDialog()
                    If i.mPrimaryAkun > 0 Then
                        Me.TxtRounding_PrimaryAccount.Text = i.mPrimaryAkun
                        Me.TxtRounding_AccountCode.Text = objCode.GetAccountCode(Me.TxtRounding_PrimaryAccount.Text)
                    Else
                        Me.TxtRounding_AccountCode.Clear()
                        Me.TxtRounding_PrimaryAccount.Text = "-1"
                        Me.TxtRounding_AccountCode.Focus()
                    End If
                End If
            Else
                Me.TxtRounding_AccountCode.Clear()
                Me.TxtRounding_PrimaryAccount.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnRounding_Find_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles btnRounding_Find.Click
        Dim i As New FrmMasterAkun_Cari
        Dim ObjCode As New ClsMasterAkun
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtRounding_AccountCode.Text <> "" Then
                i.txtCari.Text = Me.TxtRounding_AccountCode.Text
            End If

            i.ShowDialog()
            If i.mPrimaryAkun > 0 Then
                Me.TxtRounding_PrimaryAccount.Text = i.mPrimaryAkun
                Me.TxtRounding_AccountCode.Text = ObjCode.GetAccountCode(Me.TxtRounding_PrimaryAccount.Text)
                Me.btnOK.Focus()
            Else
                Me.TxtRounding_AccountCode.Clear()
                Me.TxtRounding_PrimaryAccount.Text = "-1"
                Me.TxtRounding_AccountCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtPaymentOption_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
    Handles TxtPaymentOption.KeyDown
        Dim objCode As New ClsMast_PaymentOption
        Dim i As New FrmMast_PaymentOption_Cari
        i.BwtGeneralSetting = True
        Try
            If e.KeyValue = 13 Then
                If Me.TxtPaymentOption.Text <> "" And Me.TxtPaymentOption.Text <> "Cash" Then
                    If objCode.ApakahSudahAda(Me.TxtPaymentOption.Text) = True Then
                        Me.TxtPrimaryPayment.Text = objCode.GetPrimary(Me.TxtPaymentOption.Text)
                        Me.TxtPaymentOption.Text = objCode.GetCode(Me.TxtPrimaryPayment.Text)
                        Me.BufferCurr = Me.TxtPaymentOption.Text
                        Me.TxtDefPaymentDisc.Focus()
                    Else
                        Me.BtnFindPaymentOption.PerformClick()
                    End If
                Else
                    Me.BtnFindPaymentOption.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtPaymentOption_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtPaymentOption.Leave
        Dim objCode As New ClsMast_PaymentOption
        Dim i As New FrmMast_PaymentOption_Cari
        i.BackColorAll = Me.BackColorAll
        i.BwtGeneralSetting = True
        Try
            If Me.TxtPaymentOption.Text <> "" And Me.TxtPaymentOption.Text <> "Cash" Then
                If objCode.ApakahSudahAda(Me.TxtPaymentOption.Text) = True Then
                    Me.TxtPrimaryPayment.Text = objCode.GetPrimary(Me.TxtPaymentOption.Text)
                    Me.TxtPaymentOption.Text = objCode.GetCode(Me.TxtPrimaryPayment.Text)
                    Me.BufferCurr = Me.TxtPaymentOption.Text
                Else
                    If Me.BufferCurr <> Me.TxtPaymentOption.Text Then
                        If Me.TxtPaymentOption.Text <> "" Then
                            i.txtCari.Text = Me.TxtPaymentOption.Text
                        End If

                        i.ShowDialog()
                        If i.mPrimary > 0 Then
                            Me.TxtPrimaryPayment.Text = i.mPrimary
                            Me.TxtPaymentOption.Text = objCode.GetCode(Me.TxtPrimaryPayment.Text)
                            Me.BufferCurr = Me.TxtPaymentOption.Text
                        Else
                            Me.TxtPaymentOption.Clear()
                            Me.TxtPrimaryPayment.Text = "-1"
                            Me.TxtPaymentOption.Text = "Cash"
                            Me.BufferCurr = "Cash"
                            Me.BufferCurr = Me.TxtPaymentOption.Text
                        End If
                    End If
                End If
            Else
                Me.TxtPaymentOption.Clear()
                Me.TxtPrimaryPayment.Text = "-1"
                Me.TxtPaymentOption.Text = "Cash"
                Me.BufferCurr = "Cash"
                Me.BufferCurr = Me.TxtPaymentOption.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub BtnFindPaymentOption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles BtnFindPaymentOption.Click
        Dim i As New FrmMast_PaymentOption_Cari
        Dim ObjCode As New ClsMast_PaymentOption
        i.BackColorAll = Me.BackColorAll
        i.BwtGeneralSetting = True
        Try
            If Me.BufferCurr <> Me.TxtPaymentOption.Text Then
                If Me.TxtPaymentOption.Text <> "" And Me.TxtPaymentOption.Text <> "Cash" Then
                    i.txtCari.Text = Me.TxtPaymentOption.Text
                End If

                i.ShowDialog()
                If i.mPrimary > 0 Then
                    Me.TxtPrimaryPayment.Text = i.mPrimary
                    Me.TxtPaymentOption.Text = ObjCode.GetCode(Me.TxtPrimaryPayment.Text)
                    Me.BufferCurr = Me.TxtPaymentOption.Text
                    Me.chk_Allow_CreditCard.Focus()
                Else
                    Me.TxtPaymentOption.Clear()
                    Me.TxtPrimaryPayment.Text = "-1"
                    Me.TxtPaymentOption.Text = "Cash"
                    Me.BufferCurr = "Cash"
                    Me.TxtPaymentOption.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub chk_CancelItem_Autho_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_CancelItem_Autho.CheckedChanged
        Try
            If Me.chk_CancelItem_Autho.Checked = False Then
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Cancel_item)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_CancelWhole_Autho_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_CancelWhole_Autho.CheckedChanged
        Try
            If Me.chk_CancelWhole_Allow.Checked = False Then
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Cancel_Whole_Transaction)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_Modify_Autho_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Modify_Autho.CheckedChanged
        Try
            If Me.chk_Modify_Autho.Checked = False Then
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Transaction)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_ModifyUnitPrice_Autho_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_ModifyUnitPrice_Autho.CheckedChanged
        Try
            If Me.chk_ModifyUnitPrice_Autho.Checked = False Then
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Unit_Price)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_ModifyDiscount_Autho_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_ModifyDiscount_Autho.CheckedChanged
        Try
            If Me.chk_ModifyDiscount_Autho.Checked = False Then
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Modify_Discount_Rate)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chk_Reprint_Autho_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk_Reprint_Autho.CheckedChanged
        Try
            If Me.chk_Reprint_Autho.Checked = False Then
                Me.UncheckedAllDGVAuthor(Me.dgvAutorization, EnumPengaturanAuthorization.Reprint_Receipt)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Tambahan Desain (16 Oktober 2012)
    Private Sub rbo_Transf_Stop_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbo_Transf_Stop.CheckedChanged
        Try
            If Me.rbo_Transf_Stop.Checked = True Then
                Me.gboTransfer.Enabled = False
                Me.cboExpiry.SelectedIndex = 1
                Me.cboBatch.SelectedIndex = 1
            Else
                Me.gboTransfer.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllowBatchNumber_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllowBatchNumber.CheckedChanged
        Try
            If Me.chkAllowBatchNumber.Checked = True Then
                Me.chkAllowExpiry.Enabled = True
            Else
                Me.chkAllowExpiry.Enabled = False
                Me.chkAllowExpiry.Checked = False
            End If
            If Me.chkAllowExpiry.Checked = False And Me.chkAllowLocation.Checked = False And Me.chkAllowBatchNumber.Checked = False Then
                Me.rboNo.Checked = True
                Me.rboNo.Enabled = False
                Me.rboYes.Enabled = False
            Else
                Me.rboNo.Enabled = True
                Me.rboYes.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllowLocation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllowLocation.CheckedChanged
        Try
            If Me.chkAllowExpiry.Checked = False And Me.chkAllowLocation.Checked = False And Me.chkAllowBatchNumber.Checked = False Then
                Me.rboNo.Checked = True
                Me.rboNo.Enabled = False
                Me.rboYes.Enabled = False
            Else
                Me.rboNo.Enabled = True
                Me.rboYes.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllowExpiry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAllowExpiry.CheckedChanged
        Try
            If Me.chkAllowExpiry.Checked = False And Me.chkAllowLocation.Checked = False And Me.chkAllowBatchNumber.Checked = False Then
                Me.rboNo.Checked = True
                Me.rboNo.Enabled = False
                Me.rboYes.Enabled = False
            Else
                Me.rboNo.Enabled = True
                Me.rboYes.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefPaymentDisc_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefPaymentDisc.Enter
        Try
            Me.Buffer = Me.TxtDefPaymentDisc.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefPaymentDisc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefPaymentDisc.KeyDown
        Try
            ObjCD.Disc_KeyDown(Me.TxtDefPaymentDisc, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefPaymentDisc_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDefPaymentDisc.KeyUp
        Try
            ObjCD.Disc_KeyUp(Me.TxtDefPaymentDisc, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDefPaymentDisc_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDefPaymentDisc.Leave
        Try
            ObjCD.Disc_Leave(Me.TxtDefPaymentDisc, Me.TxtDefPaymentDisc.Text, sender, e, 2, ObjCD, Me.Buffer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class