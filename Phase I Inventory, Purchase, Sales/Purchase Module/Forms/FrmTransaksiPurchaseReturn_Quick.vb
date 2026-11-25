Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiPurchaseReturn_Quick

    Dim tblCboUnit As New DseTransaksiPO.CboUnitDataTable
    Dim rowCboUnit As DseTransaksiPO.CboUnitRow
    Dim tblMultiUnit As New DseTransaksiPO.tblmaster_product_item_multiunitDataTable
    Dim rowMultiUnit As DseTransaksiPO.tblmaster_product_item_multiunitRow
    Dim objUnit As New Core_Module.ClsMasterUnit
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ObjCurrency As New Core_Module.clsMasterCurrency

    Dim baru As Boolean
    Dim objPurchaseReturn As New ClsTransaksiPurchaseReturn
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjPurchaseReturnItem As New ClsTransaksiPurchaseReturnItem
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjLS As New clsMasterListedSoftware
    Dim objGRN As New ClsTransaksiGRN
    Dim objGRNItem As New ClsTransaksiGRNItem
    Dim objInOut As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjCL As New ClsTransaksiMonthlyClosing
    Dim ObjAOP As New ClsTransaksiAOP
    Dim ObjLicense As New ClsLicense
    Dim ObjPOItem As New ClsTransaksiPOItem
    Dim objLoc As New clsMasterStorageArea
    Dim ObjSetting As New ClsTransaksiSetting
    Dim ObjSup As New clsMasterSupplier
    Dim ObjTaxRate As New Sales_Module.ClsTransaksiSalesInvoiceTaxRate
    Dim buffer As String = ""
    Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard
    Dim mFromApakahDataBenar As Boolean = False

    'dibawah ini merupakan elemen yang berkaitan dengan cbo po, cboproduct, dst
    Dim TBL_PO As New DseTransaksiPurchaseReturn.PR_OBDataTable
    Dim TBL_PR As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2DataTable
    Dim TBL_PR2 As New DseTransaksiPurchaseReturn.PR_OBDataTable
    Dim TBL_PR3 As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2DataTable

    'dibawah ini merupakan element yang berkaitan dengan stock keluar
    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL4 As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item5DataTable
    Dim TBL5 As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item6DataTable


    '===================================================================================================
    'Code Erna mulai
    '===================================================================================================
    Dim DseAccount As New AccModule.DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccount As New AccModule.DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccount As AccModule.DseMultiAccount.tblisi_cbo_ar_akunRow

    Dim DseCurrTax As New AccModule.DseMultiAccountTableAdapters.tblCurrencyTaxTableAdapter
    Dim tblCurrTax As New AccModule.DseMultiAccount.tblCurrencyTaxDataTable
    Dim rowCurrTax As AccModule.DseMultiAccount.tblCurrencyTaxRow

    Dim EraseRowIndex As Integer = 0
    Dim DefaultPrecision As Integer = 0
    Dim GetHistoricalPrice As Boolean = True

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
            If Me.tblAccount.Rows.Count > 0 Then
                Me.cboPrimaryAccount.SelectedIndex = -1
                Me.cboPrimaryAccount.SelectedIndex = 0
            End If
            Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text


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
                If Me.objPurchaseReturn.CariDetail(vPrimary, Me.dgvItem, Me, Me.TblCustomDummy, Me.TblItemInOutCustom) = True Then
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
            If Me.tblCurrTax.Rows.Count > 0 Then
                Me.cboPrimaryCurrTax.SelectedIndex = -1
                Me.cboPrimaryCurrTax.SelectedIndex = 0
            End If
            Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text

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
                Me.HitungTotal_Multi()
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
                Me.HitungTotal_Multi()
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

    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCurrency.TextChanged

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

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, True, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.BackColor, False)

            Me.KosongkanCBO()
            Me.btnFindCurrency.Enabled = False
            Me.btnFindSupplier.Enabled = False
            Me.chkTax.Enabled = False
            Me.btnFindStockInPurchase.Enabled = False

            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)

            Me.TxtPrimaryPurchaseReturn.Text = "-1"
            Me.txtPrimaryGRN.Text = "-1"
            Me.TxtPrimaryPurchaseReturn.Text = "-1"
            Me.txtPrimaryStockInPurchase.Text = "-1"
            Me.txtPrimarySupplier.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()

            Me.TxtPaymentTerm.Text = "0"
            Me.txtSubTotalAfterDP.Text = "0"
            Me.txtTax.Text = "0"
            Me.txtTaxExchangeRate.Text = "1"
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.TxtDutystamp.Text = "0"
            Me.TxtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtPrimaryAccount.Text = "-1"
            Me.chkTax.Checked = False

            Me.cboCodeNo.DataSource = Nothing
            Me.cboCodeNo.Items.Clear()
            Me.cboPO.DataSource = Nothing
            Me.cboPO.Items.Clear()
            Me.KosongkanItem()
            Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
            Me.TxtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency


            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.cboCurrTax.DataSource = Nothing

            Me.cboPrimaryAccount.Text = ""
            Me.cboAccount.Text = ""
            Me.cboCurrTax.Text = ""

            Me.cboPrimaryAccount.Items.Clear()
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.cboAccount.Items.Clear()
            Me.cboCurrTax.Items.Clear()


            'TR
            Me.TblCustomDummy.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SalinKeField()
        Dim rounding As Double = 0

        Try
            With Me.objPurchaseReturn
                Me.TxtPrimaryPurchaseReturn.Text = .Primary
                Me.TxtPurchaseReturnNo.Text = .NoTransaksi
                Me.txtPrimaryStockInPurchase.Text = Me.objGRN.GetPrimaryStockIn(.PrimaryGRN)
                Me.txtPrimaryGRN.Text = .PrimaryGRN
                Me.dtpDate.Text = .TanggalTransaksi
                Me.TxtReasonOfReturn.Text = .ReasonOfReturn
                Me.TxtRemark.Text = .Remarks
                Me.txtSupplierDONumber.Text = .SupplierDONumber
                Me.TxtPrimaryCurrency.Text = .PrimaryCurrency

                If Me.TxtPrimaryCurrency.Text <> "-1" And Me.TxtPrimaryCurrency.Text <> "" Then
                    Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    rounding = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                End If


                Me.TxtPaymentTerm.Text = Me.ObjCD.DigitGrouping(.PaymentTerm)
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(.TaxExchangeRate, 2)
                Me.cboTipePurchaseReturn.SelectedIndex = .TipeReturn
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, rounding))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, rounding))

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString


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

                Me.txtTax.Text = Me.ObjCD.DigitGrouping(.Tax, 2)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))
                Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(.Dutystamp, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text))

                'Update NPWP 16 Digit
                Me.TxtNPWPNo.Text = .NoNPWPSupplier
            End With

            Me.FormatDGV()

            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value)
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    row.Cells(Me.ColSecQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecQty.Index).Value)
                End If
            Next
            Me.IsiDataStockIn()

            If Me.txtSupplierCode.Text = "" Then
                Me.txtPrimarySupplier.Text = Me.objPurchaseReturn.PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSup.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSup.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.chkTax.Checked = Me.objPurchaseReturn.TaxExclude
                Me.txtTax.Text = Me.objPurchaseReturn.Tax
            End If

            Me.HitungTotal_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objPurchaseReturn
                .Primary = Me.TxtPrimaryPurchaseReturn.Text
                .NoTransaksi = Me.TxtPurchaseReturnNo.Text
                .PrimaryGRN = Me.txtPrimaryGRN.Text
                .TanggalTransaksi = Me.dtpDate.Value
                If Me.chkTax.Checked = True Then
                    .TaxExclude = True
                Else
                    .TaxExclude = False
                End If
                .ReasonOfReturn = Me.TxtReasonOfReturn.Text
                .Remarks = Me.TxtRemark.Text
                .SupplierDONumber = Me.txtSupplierDONumber.Text
                .PrimaryCurrency = Me.TxtPrimaryCurrency.Text
                .Dutystamp = Me.TxtDutystamp.Text
                .PaymentTerm = Me.TxtPaymentTerm.Text
                .TaxExchangeRate = Me.txtTaxExchangeRate.Text
                .TipeReturn = Me.cboTipePurchaseReturn.SelectedIndex
                .Total = Me.TxtTotal.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .SubTotal = Me.txtSubTotalAfterDP.Text
                .Tax = Me.txtTax.Text
                .PrimarySupplier = Me.txtPrimarySupplier.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text
                'Update NPWP 16 Digit
                .NoNPWPSupplier = Me.TxtNPWPNo.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '==================================================================================================
    'Code Erna selesai
    '==================================================================================================


    Private Sub TxtStockOutNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True And Me.TxtStockInNo.Enabled = True Then
                If Me.TxtStockInNo.ReadOnly = True Then
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtStockInNo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockOutNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.LostFocus
        Try
            If Me.TxtStockInNo.Enabled = True Or Me.TxtRemark.Enabled = True Then
                If Me.TxtStockInNo.ReadOnly = True Then
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReasonOfReturn_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReasonOfReturn.GotFocus
        TxtReasonOfReturn.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtReasonOfReturn_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReasonOfReturn.Leave

    End Sub

    Private Sub TxtReasonOfReturn_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReasonOfReturn.LostFocus
        TxtReasonOfReturn.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel1.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer7.Panel2.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.TxtPrimaryPurchaseReturn.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer7.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, True, True)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.FormatTxtStockIN()
            Me.DtpExpired.Enabled = False
            btnFindCurrency.Enabled = True
            Me.btnFindSupplier.Enabled = True
            Me.chkTax.Enabled = True
            Me.SetTax(Me.baru)

            Me.cboAccount.Enabled = True
            Me.cboCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
            End If

            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtCurrency.ReadOnly = True
                Me.TxtCurrency.BackColor = Me.txtReadOnly.BackColor
                Me.TxtCurrency.TabStop = False
                Me.btnFindCurrency.Enabled = False
            Else
                Me.TxtCurrency.ReadOnly = False
                Me.TxtCurrency.BackColor = Me.TxtPrimaryCurrency.BackColor
                Me.TxtCurrency.TabStop = True
                Me.btnFindCurrency.Enabled = True
            End If

            'TR
            Me.LblAdvance.Enabled = True
            Me.Atur1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)

            Me.TxtUnitPrice.Text = "0"
            Me.TxtNettPrice.Text = "0"
            Me.TxtDiscount.Text = "0"
            Me.TxtAmount.Text = "0"
            Me.TxtDiscAmount.Text = "0"

            Me.txtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryPO.Text = "-1"
            Me.TxtPrimaryUnitsDefault.Text = "-1"
            Me.TxtPrimaryUnitsMulti.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyMulti.Text = "0"
            Me.txtSecondaryQty.Text = "0"
            Me.TxtRatioReference.Text = "1"
            Me.TxtRatioDefault.Text = "1"

            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.txtSecondaryQty.Visible = False
            Me.lblSecondaryUnits.Visible = False
            Me.lblSecQty.Visible = False
            Me.DtpExpired.Value = Today
            Me.DtpExpired.Visible = False
            Me.txtDesc.Text = ""
            Me.TxtQtyOnHand.Text = "0"
            Me.TxtDiscFromPerct.Text = "1"

            ''@rosadinnn
            ''@MUaT
            Me.LblUnitPrice.Text = "Unit Price"
            Me.CboUnitMulti.Visible = False
            Me.LblQtyOnHand.Text = "Qty On Hand"
            Me.LblUnitsMulti.Text = "Units"
            Me.LblUnitsMulti.Visible = True

            Me.cboPriceByMultiUnit.Visible = False
            Me.cboPriceByMultiUnit.TabStop = False
            Me.txtPriceByMultiUnit.Text = "1"
            Me.cboPriceByMultiUnit.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim QtyAcuan As Decimal = 0

        Try
            'data item salah bila:
            '0. tidak ada stock in number yang diberikan
            'If Me.TxtStockInNo.Text = "" Then
            'MsgBox("Please specify stock in purchase number first", MsgBoxStyle.Information, "Information")
            'Me.TxtStockInNo.Focus()
            'Return False
            'End If
            '0. jika tidak merujuk pada po, maka harus ada data currency
            If Me.cboPO.Visible = False Then
                If Me.TxtCurrency.Text = "" Then
                    MsgBox("Please enter currency code first", MsgBoxStyle.Information, "Information")
                    Me.mFromApakahDataBenar = False
                    Me.TxtCurrency.Focus()
                    Return False
                End If
            End If
            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.mFromApakahDataBenar = False
                Me.TxtCodeNo.Focus()
                Return False
            End If

            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.mFromApakahDataBenar = False
                Me.TxtQtyMulti.Focus()
                Return False
            End If



            '2. nilai TxtQtyDefault <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyDefault.Text) = False Then
                MsgBox("Qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyMulti.Focus()
                Return False
            End If



            '3. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnit.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.txtSecondaryQty.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.mFromApakahDataBenar = False
                    Me.txtSecondaryQty.Focus()
                    Return False
                End If
            End If

            '4. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlled.Checked = True Then
                If Me.cboBatchNo.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.mFromApakahDataBenar = False
                    Me.cboBatchNo.Focus()
                    Return False
                End If
            End If

            '5. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlled.Checked = True Then
                If Me.DtpExpired.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.mFromApakahDataBenar = False
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
            End If

            '6. cocokan expiry dengan batch number
            If Me.chkBatchNoControlled.Checked = True And Me.chkExpiryDateControlled.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.cboBatchNo.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.mFromApakahDataBenar = False
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatchNo.Index).Value = Me.cboBatchNo.Text And Me.cboBatchNo.Text <> "" And row.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                    Me.txtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColExpiryDate.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.mFromApakahDataBenar = False
                            Me.DtpExpired.Focus()
                            Return False
                        End If
                    End If
                Next
            End If

            '7. pastikan unit price positive
            If CDec(Me.TxtUnitPrice.Text) <= 0 Then
                MsgBox("Please enter a positive number for unit price.", MsgBoxStyle.Information, "Information")
                Me.mFromApakahDataBenar = False
                Me.TxtUnitPrice.Focus()
                Return False
            End If

            '9. Bila setelah diinput jumlah product menjadi minus
            If Me.ObjProduct.ApakahGoods(Me.txtPrimaryCodeNo.Text) = True Then
                Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
                JmlGudang = Me.ObjProduct.GetJml3(Me.txtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlObjLama = Me.ObjPurchaseReturnItem.GetJml2(Me.TxtPrimaryPurchaseReturn.Text, Me.txtPrimaryCodeNo.Text, _
                Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                JmlDGV = Me.GetJmlDGV2(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)

                ''@rosadinnn
                ''@MUaT
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyDefault.Text)

                If JmlSetelahTambah < 0 Then
                    If Me.chkUseSecondaryUnit.Checked = True Then
                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping((JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)) + " " + _
                        Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.txtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, _
                        "Information")
                    Else
                        Dim Max_ As String = ((JmlGudang + JmlObjLama - JmlDGV) * CDec(Me.TxtRatioReference.Text)) / CDec(Me.TxtRatioDefault.Text)
                        Max_ = Me.ObjCD.RoundDown1(Max_, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))

                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(Max_, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)) + " " + _
                        Me.CboUnitMulti.Text + ".", MsgBoxStyle.Information, _
                        "Information")
                    End If
                    Me.mFromApakahDataBenar = False
                    Me.TxtQtyMulti.Focus()
                    Return False
                End If

                '9.b. Cek stock pada tanggal yang bersangkutan
                JmlGudang = ObjSC.GetStock(Me.txtPrimaryCodeNo.Text, Me.dtpDate.Value)
                JmlObjLama = Me.ObjPurchaseReturnItem.GetJml2(Me.TxtPrimaryPurchaseReturn.Text, Me.txtPrimaryCodeNo.Text, _
                Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                JmlDGV = Me.GetJmlDGV2(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, _
                Me.cboBatchNo.Text)

                ''@rosadinnn
                ''@MUaT
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyDefault.Text)

                If JmlSetelahTambah < 0 Then
                    If JmlGudang + JmlObjLama - JmlDGV > 0 Then
                        If Me.chkUseSecondaryUnit.Checked = True Then
                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + ObjCD.DigitGrouping((JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)) + " " + _
                            Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                            Me.ObjProduct.GetSecondaryUnitCode(Me.txtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, _
                            "Information")
                        Else
                            Dim Max_ As String = ((JmlGudang + JmlObjLama - JmlDGV) * CDec(Me.TxtRatioReference.Text)) / CDec(Me.TxtRatioDefault.Text)
                            Max_ = Me.ObjCD.RoundDown1(Max_, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))

                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + ObjCD.DigitGrouping(Max_, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)) + " " + _
                            Me.CboUnitMulti.Text + ".", MsgBoxStyle.Information, _
                            "Information")
                        End If
                    Else
                        MsgBox("Stock level on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + _
                        " is 0. Entering this data will cause negative value in your stock.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.mFromApakahDataBenar = False
                    Me.TxtQtyMulti.Focus()
                    Return False
                End If
            End If

            Me.mFromApakahDataBenar = False
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Double
        Dim rowdgv As DataGridViewRow
        Dim TotalQty As Double = 0

        Try
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    TotalQty += rowdgv.Cells(Me.ColTotalQty.Index).Value
                End If
            Next

            Return TotalQty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV2(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNumber As String) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                rowDGV.Cells(Me.ColBatchNo.Index).Value = vBatchNumber Then
                    RetVal += CDec(rowDGV.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim NoAOP As String = ""
        Dim NoStockOut As String = ""
        Dim DTA As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
        Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
        Dim Jurnal As String = ""
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim DTA_Convt As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_accounting_closingTableAdapter
        Dim ListPrimaryProduct As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim ObjPRCustom As New CustomSMESS.ClsTransaksiPurchaseReturnQuick_AddByCustomProduct
        Dim Dummy As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            If DTA_Convt.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA_Convt.ScalarQuery_GetConvertionDate.Value.Date > Me.dtpDate.Value Then
                    MsgBox("Please specify transaction date after convertion date", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox("Please specify transaction date after convertion date", MsgBoxStyle.Information, "Information")
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

            'cek data yang harus diisi
            If Me.cboPO.Visible = True Then
                KoleksiNamaData.Add("stock in number")
                KoleksiTextBox.Add(Me.TxtStockInNo)
            End If
            KoleksiNamaData.Add("Supplier code")
            KoleksiTextBox.Add(Me.txtSupplierCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek supplier jangan sampai tidak diisikan
            If Me.TxtPurchaseReturnNo.Text = "-1" Then
                MsgBox("Please specify stock in purchase number", MsgBoxStyle.Information, "Information")
                Me.TxtPurchaseReturnNo.Focus()
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
                If Me.objPurchaseReturn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objPurchaseReturn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objPurchaseReturn.TanggalTransaksi
                        Me.TxtStockInNo.Text = Me.objPurchaseReturn.NoTransaksi
                        Return False
                    End If
                End If
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.objPurchaseReturn.TimeStamp <> Me.objPurchaseReturn.GetTimestamp(Me.objPurchaseReturn.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objPurchaseReturn.CariDetail(Me.objPurchaseReturn.Primary, Me.dgvItem, Me, Me.TblCustomDummy, Me.TblItemInOutCustom)
                        Me.SalinKeField()
                    End If
                    Return False
                End If

                '1. Jika sudah ada AOP yang merujuk ke purchase return ini
                'untuk AOP, berikan pesan no AOP dan no stock out purchase return
                NoAOP = ObjAOP.GetNoAOPBy_PrimaryPurchaseReturn(Me.objPurchaseReturn.Primary)
                If NoAOP <> "" Then
                    NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.objPurchaseReturn.Primary)
                    MsgBox("This purchase return is refer by AOP #" + NoAOP + _
                    " and stock out #" + NoStockOut + ". Please delete the transaction to edit this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If

                '3. sudah ada stock out purchase return yang merujuk ke purchase return ini-karena quick pr, maka batasan ini dihapuskan

                '4. jika sudah ada purchase return receivement journal
                Jurnal = ObjJurnal.Get_PurchaseReturn_Jurnal(Me.objPurchaseReturn.Primary)
                If Jurnal <> "" Then
                    MsgBox("This data is refered by purchase return receivement jounal: " + Jurnal + _
                        ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'berkaitan dengan tax exchange rate
            If Me.TxtPrimaryCurrency.Text = "1" Then
                If CDbl(Me.txtTaxExchangeRate.Text) <> 1 Then
                    MsgBox("For data with " + Me.ObjCurrency.GetCurrencyCode(1) + " currency, the tax exchange rate value must be 1.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtTaxExchangeRate.Focus()
                    Return False
                End If
            Else
                'jika bukan IDR dan tax ratenya bernilai 1 maka berikan peringatan
                If CDbl(Me.txtTaxExchangeRate.Text) = 1 Then
                    MsgBox("This purchase return use non-" + Me.ObjCurrency.GetCurrencyCode(1) + " currency and the tax exchange rate value = 1. This is wrong data input.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtTaxExchangeRate.Focus()
                    Return False
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
                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Me.txtTax.Focus()
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If

                If CDec(Me.txtTax.Text) <> CDec(Dummy) Then
                    If MsgBox("Tax rate <> " + Me.ObjCD.DigitGrouping(Dummy, 2) + "%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Me.txtTax.Focus()
                        Return False
                    End If
                End If
            End If

            Dim ObjCust As New CustomSMESS.ClsTransaksiStockInOut
            Dim rd As DataGridViewRow
            Dim TblDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
            TblDummy = Me.TblItemInOutCustom.Copy
            If TblDummy.Rows.Count > 0 Then
                For ij As Integer = 0 To Me.dgvItem.RowCount - 1
                    rd = Me.dgvItem.Rows(ij)
                    With rd
                        Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                        Dim PrimaryPOSO As Double = .Cells(Me.ColPrimaryPO.Index).Value
                        Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                        Dim Batch As String = .Cells(Me.ColBatchNo.Index).Value
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
            End If
            'cek batasan custom
            If ObjPRCustom.ApakahDataBenar(Me.dgvItem, Me.ColPrimaryCodeNo.Index, Me.colPriceByMultiUnit.Index, Me.ColPrimaryUnitMulti.Index, _
            Me.ColPrimaryUnit.Index) = False Then
                Return False
            End If

            'Update NPWP 16 Digit
            If chkTax.Checked = True Then 'hanya dicek ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                '2025 Core Tax Tahap 2.1
                If ObjSetNPWP.SetNPWP_New(Me.dtpDate.Value.Date, Me.txtPrimarySupplier.Text, Me.TxtNPWPNo.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Supplier, _
                 frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionPurchaseReturn, Me.UserCode, Me.TxtPrimaryPurchaseReturn.Text) = False Then
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
            DTA.Dispose()
            ObjJurnal = Nothing
            DTA_Convt.Dispose()
            ListPrimaryProduct = Nothing
        End Try
    End Function


    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.chkTax.Checked = False
            Me.chkTax.Checked = True
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.cboTipePurchaseReturn.SelectedIndex = 1
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.btnUP.BackgroundImage = My.Resources.arrow_up
            Me.btnDown.BackgroundImage = My.Resources.arrow_down
            Me.CboUnitMulti.TabStop = True
            Me.CboUnitMulti.Visible = False
            'TR
            Me.Atur1()
            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.SplitContainer6.TabStop = False
            Me.SplitContainer7.TabStop = False
            Me.Text = "Quick Purchase Return"
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
                        If Me.objPurchaseReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.TblCustomDummy) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.SalinKeField()
                            Me.btnFind.Focus()
                        End If

                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objPurchaseReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.TblCustomDummy) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.SalinKeField()
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
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.dtpDate.Value.Date = Today And _
            Me.TxtCodeNo.Text = "" And Me.cboCodeNo.SelectedIndex = -1 And _
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                        Me.NonAktifkanDetail()
                        Me.btnNew.Focus()
                    Else
                        Me.NonAktifkanDetail()
                        Me.btnNew.Focus()
                        Me.objPurchaseReturn.CariDetail(Me.TxtPrimaryPurchaseReturn.Text, Me.dgvItem, Me, Me.TblCustomDummy, Me.TblItemInOutCustom)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim UseTaxSetting As String = ""
        Dim TaxPercentSetting As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.objPurchaseReturn.ResetObject(Me.dgvItem)
            Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.GetNoBaru(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'yang quick purchase return hanya yang nota return pajak saja
            Me.cboTipePurchaseReturn.SelectedIndex = 1

            'load posisi awal
            Me.TxtPrimaryCurrency.Text = 1
            Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.AOP, ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.chkTax.Checked = True
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
            Else
                If UseTaxSetting = "0" Then
                    Me.chkTax.Checked = False
                Else
                    Me.chkTax.Checked = True
                    TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(TaxPercentSetting) = True Then
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping(TaxPercentSetting, 2)
                    End If
                End If
            End If

            Me.SetTaxRate(True)
            Me.HitungTotal_Multi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiPurchaseReturn
        Try
            ij.IsQuick = 1
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryPurchaseReturn = -1 Then
                If Me.objPurchaseReturn.CariDetail(ij.PrimaryPurchaseReturn, Me.dgvItem, Me, Me.TblCustomDummy, Me.TblItemInOutCustom) = True Then
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
            If Me.TxtPurchaseReturnNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPurchaseReturnNo.Text <> "" Then
                'jika quick purchase return maka tidak boleh diedit.
                If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = False Then
                    MsgBox("Please edit this data from purchase return form.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objPurchaseReturn.ApakahBolehEdit(Me.UserCode, True, True, True) = True Then
                    '1 baris ini ditambahkan, karena setiap kali print pr, ada perubahan timestamp
                    Me.objPurchaseReturn.TimeStamp = Me.objPurchaseReturn.GetTimestamp(Me.objPurchaseReturn.Primary)

                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dtpDate.Focus()
                    Me.SetTaxRate(False)
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
            If Me.TxtPurchaseReturnNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPurchaseReturnNo.Text <> "" Then
                'jika quick purchase return maka tidak boleh diedit.
                If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = False Then
                    MsgBox("Please delete this data from purchase return form.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objPurchaseReturn.ApakahBolehHapus(Me.UserCode, True, True, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objPurchaseReturn.Hapus(Me.TxtPrimaryPurchaseReturn.Text, Me.UserCode, True) = True Then
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemark.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objPurchaseReturn.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.TblCustomDummy) = True Then
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
                            If Me.objPurchaseReturn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.TblCustomDummy) = True Then
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

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.txtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
            Me.chkPdoControlled, Me.TxtPrimaryUnitsDefault, Me.txtPrimarySecondaryUnits)

            Me.cboBatchNo.DataSource = Nothing
            Me.cboLocation.DataSource = Nothing

            If Me.chkExpiryDateControlled.Checked = True Then
                Me.DtpExpired.Visible = True
                Me.DtpExpired.Value = Today
                Me.lblExpiryDate.Visible = False
            Else
                Me.lblExpiryDate.Visible = True
                Me.DtpExpired.Visible = False
                Me.DtpExpired.Value = #1/1/1950#
            End If

            If Me.chkBatchNoControlled.Checked = True Then
                Me.cboBatchNo.Visible = True
                Me.cboBatchNo.TabStop = True
                'isikan data batch no dari basis data
                Me.IsiCboBatch(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryPurchaseReturn.Text)
            Else
                Me.cboBatchNo.Visible = False
                Me.cboBatchNo.TabStop = False
                Me.cboBatchNo.Items.Clear()
            End If
            If Me.chkLocationControlled.Checked = True Then
                Me.cboLocation.Visible = True
                Me.cboLocation.TabStop = True
                Me.IsiCboLokasi(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryPurchaseReturn.Text, Me.cboBatchNo.Text)
            Else
                Me.cboLocation.Visible = False
                Me.cboLocation.TabStop = False
            End If

            If Me.chkUseSecondaryUnit.Checked = False Then
                Me.txtSecondaryQty.Visible = False
                Me.lblSecondaryUnits.Visible = False
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
            Else
                Me.txtSecondaryQty.Visible = True
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                Me.txtSecondaryQty.Text = "0"
            End If

            If Me.txtPrimaryCodeNo.Text <> "-1" Then
                Me.LblUnitsMulti.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.txtPrimaryCodeNo.Text)
                End If
            Else
                Me.LblUnitsMulti.Text = "Units"
                Me.lblSecondaryUnits.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboBatch(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Try
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNo, Me.TBL)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item5Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean

            If Me.baru = False Then
                Me.ObjPurchaseReturnItem.IsikanTBLBatchNo(Me.TBL4, vPrimaryProduct, vPrimaryTransaksi)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    For Each row2 In Me.TBL.Rows
                        If row1.Lot_No = row2.Lot_No Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL.Rows.Add(row1.Lot_No, row1.Expiry_Date, row1.Qty)
                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If

            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLBatchNo(vPrimaryProduct)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                For Each row2 In Me.TBL.Rows
                    If row1.Lot_No = row2.Lot_No Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4.Rows.Remove(row1)
            End While

            If Me.chkExpiryDateControlled.Checked = False Then
                Me.TBL.DefaultView.Sort = "[Lot No] ASC"
            Else
                Me.TBL.DefaultView.Sort = "[Expiry Date] ASC"
            End If
            Me.cboBatchNo.DataSource = Me.TBL.DefaultView
            Me.cboBatchNo.ValueMember = "Lot No"
            Me.cboBatchNo.DisplayMember = "Lot No"
            Me.cboBatchNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Double)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item5Row

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColBatchNo.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = True Then
                            ED = row.Cells(Me.ColExpiryDate.Index).Value
                        End If

                        row1 = Me.TBL4.Newtbltrans_purchasereturn_item5Row

                        With row1
                            .Expiry_Date = ED
                            .Lot_No = row.Cells(Me.ColBatchNo.Index).Value
                            .Qty = row.Cells(Me.ColTotalQty.Index).Value
                        End With

                        Me.TBL4.Rows.Add(row1)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.txtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    Me.SetQtyOnHand()
                    If Me.cboBatchNo.Visible = True Then
                        Me.cboBatchNo.Focus()
                    ElseIf Me.cboLocation.Visible = True Then
                        Me.cboLocation.Focus()
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        Me.SetQtyOnHand()
                        If Me.cboBatchNo.Visible = True Then
                            Me.cboBatchNo.Focus()
                        ElseIf Me.cboLocation.Visible = True Then
                            Me.cboLocation.Focus()
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.txtDesc.Text = ""
                        Me.TxtQtyMulti.Text = "0"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            ElseIf e.KeyData = Keys.Tab Then
                If Me.cboBatchNo.Visible = True Then
                    Me.cboBatchNo.Focus()
                ElseIf Me.cboLocation.Visible = True Then
                    Me.cboLocation.Focus()
                Else
                    Me.TxtQtyMulti.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()

            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.txtPrimaryCodeNo.Text)

        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.txtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    Me.SetQtyOnHand()

                    If Me.TxtCodeNo.Focused = True Then
                        If My.Computer.Keyboard.ShiftKeyDown = False Then
                            If Me.cboBatchNo.Visible = True Then
                                Me.cboBatchNo.Focus()
                            ElseIf Me.cboLocation.Visible = True Then
                                Me.cboLocation.Focus()
                            Else
                                Me.TxtQtyMulti.Focus()
                            End If
                        Else
                            Me.TxtDutystamp.Focus()
                        End If
                    End If
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        Me.SetQtyOnHand()
                        If My.Computer.Keyboard.ShiftKeyDown = False Then
                            If Me.cboBatchNo.Visible = True Then
                                Me.cboBatchNo.Focus()
                            ElseIf Me.cboLocation.Visible = True Then
                                Me.cboLocation.Focus()
                            Else
                                Me.TxtQtyMulti.Focus()
                            End If
                        Else
                            Me.TxtDutystamp.Focus()
                        End If
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.txtDesc.Text = ""
                        Me.TxtQtyMulti.Text = "0"
                        Me.SetGboItem()
                    End If
                End If
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.txtDesc.Text = ""
                Me.TxtQtyMulti.Text = "0"
                Me.SetGboItem()

                If Me.TxtCodeNo.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = False Then
                        If Me.cboBatchNo.Visible = True Then
                            Me.cboBatchNo.Focus()
                        ElseIf Me.cboLocation.Visible = True Then
                            Me.cboLocation.Focus()
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    Else
                        Me.TxtDutystamp.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.txtPrimaryCodeNo.Text)

        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Me.SetGboItem()
                Me.SetQtyOnHand()
                If Me.cboBatchNo.Visible = True Then
                    Me.cboBatchNo.Focus()
                ElseIf Me.cboLocation.Visible = True Then
                    Me.cboLocation.Focus()
                Else
                    Me.TxtQtyMulti.Focus()
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            ''@rosadinnn
            ''@MUaT
            Me.PrepareMultiUnit(Me.txtPrimaryCodeNo.Text)
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan index baris dalam dgv yang sesuai dengan parameter yang diberikan
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="BatchNo"></param>
    ''' <param name="SecondaryQty"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Double, ByVal ExpiryDate As Date, ByVal vPrimaryPO As Double, _
    ByVal vPrimaryLocation As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = #1/1/1950# _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
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

    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal BatchNo As String, _
   ByVal SecondaryQty As Double, ByVal ExpiryDate As Date, ByVal vPrimaryPO As Double, _
   ByVal vPrimaryLocation As Double, ByVal vUnitPrice As Decimal, ByVal vDiscount As Decimal, ByVal vPrimaryUnitMulti As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And _
                        row.Cells(Me.ColDiscount.Index).Value = vDiscount And _
                        row.Cells(Me.ColPrimaryUnitMulti.Index).Value = vPrimaryUnitMulti Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = #1/1/1950# _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And _
                        row.Cells(Me.ColDiscount.Index).Value = vDiscount And _
                        row.Cells(Me.ColPrimaryUnitMulti.Index).Value = vPrimaryUnitMulti Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And _
                        row.Cells(Me.ColDiscount.Index).Value = vDiscount And _
                        row.Cells(Me.ColPrimaryUnitMulti.Index).Value = vPrimaryUnitMulti Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) _
                        And row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                        row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And _
                        row.Cells(Me.ColDiscount.Index).Value = vDiscount And _
                        row.Cells(Me.ColPrimaryUnitMulti.Index).Value = vPrimaryUnitMulti Then
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

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim rowDGV As DataGridViewRow
        Dim ExpiredDateStr As String = ""
        Dim rounding As Integer = 0

        Try
            If Me.ApakahDataItemBenar = True Then
                Dim Index As Integer = 0

                If Me.cboPO.Visible = True Then
                    Index = Me.GetIndex(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text, _
                    Me.txtSecondaryQty.Text, Me.DtpExpired.Value, Me.txtPrimaryPO.Text, Me.txtPrimaryLocation.Text)
                Else
                    Index = Me.GetIndex(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text, _
                    Me.txtSecondaryQty.Text, Me.DtpExpired.Value, Me.txtPrimaryPO.Text, Me.txtPrimaryLocation.Text, _
                    Me.TxtUnitPrice.Text, Me.TxtDiscount.Text, Me.TxtPrimaryUnitsMulti.Text)
                End If

                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)

                    ''@rosadinnn
                    ''@MUaT
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(Me.TxtQtyMulti.Text)
                    Me.SetQtyDefaultOnDgv(row.Index)

                    rounding = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    If row.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                                CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), rounding))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
        CDec(row.Cells(Me.ColTotalQty.Index).Value), rounding))
                    End If

                    Me.HitungTotal_Multi()
                Else
                    'tambahkan baris baru
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count.ToString
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColPONO.Index).Value = Me.cboPO.Text
                        .Cells(Me.ColBatchNo.Index).Value = Me.cboBatchNo.Text
                        If Me.DtpExpired.Value.Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                        End If
                        .Cells(Me.ColExpiryDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColTotalQty.Index).Value = Me.TxtQtyMulti.Text
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyMulti.Text
                        .Cells(Me.ColUnit.Index).Value = Me.LblUnitsMulti.Text
                        .Cells(Me.ColSecQty.Index).Value = -1
                        .Cells(Me.ColSecUnit.Index).Value = ""
                        .Cells(Me.colRemarks.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryPO.Index).Value = Me.txtPrimaryPO.Text
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryUnit.Index).Value = Me.TxtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColLocation.Index).Value = Me.cboLocation.Text

                        If Me.txtPrimaryPO.Text = "" Or Me.txtPrimaryPO.Text = "-1" Then
                            .Cells(Me.ColUnitPrice.Index).Value = Me.TxtUnitPrice.Text
                            .Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)
                            .Cells(Me.ColDiscAmount.Index).Value = Me.TxtDiscAmount.Text
                            .Cells(Me.ColDiscFromPerct.Index).Value = Me.TxtDiscFromPerct.Text
                            .Cells(Me.ColNettPrice.Index).Value = Me.TxtNettPrice.Text
                            .Cells(Me.ColAmount.Index).Value = Me.TxtAmount.Text
                            .Cells(Me.ColExchangeRate.Index).Value = 1
                            .Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency.Text
                        End If

                        ''@rosadinnn
                        ''@MUaT
                        .Cells(Me.ColTotalQty.Index).Value = Me.TxtQtyDefault.Text
                        .Cells(Me.ColUnit.Index).Value = Me.LblUnitsDefault.Text
                        .Cells(Me.ColTotalQtyMulti.Index).Value = Me.TxtQtyMulti.Text
                        .Cells(Me.ColUnitsCodeMulti.Index).Value = Me.LblUnitsMulti.Text
                        .Cells(Me.ColPrimaryUnitMulti.Index).Value = Me.TxtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                        If Me.TxtPrimaryUnitsMulti.Text <> Me.TxtPrimaryUnitsDefault.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                                    ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)) _
                                    & " " & Me.LblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                                    Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)) & " " & Me.LblUnitsDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If

                        .Cells(Me.colPriceByMultiUnit.Index).Value = Me.txtPriceByMultiUnit.Text
                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            .Cells(Me.colPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnitsCodeMulti.Index).Value
                        Else
                            .Cells(Me.colPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnit.Index).Value
                        End If
                    End With

                    'isikan data harga+hitung total
                    Me.HitungTotal_Multi()
                End If

                'format dgv
                Me.FormatDGV()
                Me.FormatTxtStockIN()
                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                'kosongkan gbo item
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
                Me.cboBatchNo.DataSource = Nothing
                Me.cboBatchNo.Items.Clear()
                Me.cboBatchNo.Visible = False
                Me.cboBatchNo.TabStop = False
                Me.DtpExpired.Visible = False
                Me.DtpExpired.TabStop = False
                Me.cboLocation.Visible = False
                Me.cboLocation.DataSource = Nothing
                Me.cboLocation.Items.Clear()
                Me.cboLocation.TabStop = False
            End If

            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtCurrency.ReadOnly = True
                Me.TxtCurrency.BackColor = Me.txtReadOnly.BackColor
                Me.TxtCurrency.TabStop = False
                Me.btnFindCurrency.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanCBO()
        Try
            Me.cboPO.DataSource = Nothing
            Me.cboPO.Items.Clear()
            Me.cboCodeNo.DataSource = Nothing
            Me.cboCodeNo.Items.Clear()
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Items.Clear()
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Items.Clear()
            Me.cboLocation.Visible = False
            Me.cboBatchNo.Visible = False
            Me.DtpExpired.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim TidakAdaLocation As Boolean = True
        Dim TidakAdaPO As Boolean = True
        Dim TidakAdaMultiUnit As Boolean = True

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    row.Cells(Me.ColExpiryDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiryDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiryDate.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiryDate.Index).Value = Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.ColSecQty.Index).Value = "-1" Then
                    row.Cells(Me.ColSecQty.Index).Value = "-"
                    row.Cells(Me.ColSecUnit.Index).Value = "-"
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
                '1.e tandai bila ada location
                If row.Cells(Me.ColPrimaryLocation.Index).Value <> -1 Then
                    TidakAdaLocation = False
                End If
                If row.Cells(Me.ColPrimaryPO.Index).Value <> -1 Then
                    TidakAdaPO = False
                End If
                '1.f multi unit
                If row.Cells(Me.ColPrimaryUnitMulti.Index).Value <> row.Cells(Me.ColPrimaryUnit.Index).Value Then
                    TidakAdaMultiUnit = False
                End If
            Next

            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaMultiUnit = True Then
                Me.ColRatioDesc.Visible = False
                Me.ColTotalQty.Visible = False
                Me.ColUnit.Visible = False
            Else
                Me.ColRatioDesc.Visible = True
                Me.ColTotalQty.Visible = True
                Me.ColUnit.Visible = True
            End If
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = True
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColBatchNo.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColBatchNo.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiryDate.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiryDate.Index).Visible = True
            End If
            '6. jika semua data location tidak ada maka tidak perlu tampilkan kolom location
            If TidakAdaLocation = True Then
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = True
            End If
            'jika ada po yang dirujuk, maka stock in number tidak bisa diedit
            If Me.TxtRemark.Enabled = True Then
                If TidakAdaPO = False Then
                    Me.TxtStockInNo.ReadOnly = True
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                    Me.btnFindStockInPurchase.Enabled = False
                    Me.TxtStockInNo.TabStop = False
                Else
                    Me.TxtStockInNo.ReadOnly = False
                    Me.btnFindStockInPurchase.Enabled = True
                    Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                    Me.TxtStockInNo.TabStop = True
                End If
            End If

            'untuk exchange rate tidak diperhitungkan terlebih dahulu
            Me.ColExchangeRate.Visible = False
            Me.ColCurrency.Visible = False

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Enter
        Me.ObjCD.Number_Enter(Me.TxtQtyMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyMulti, sender, e, False)
    End Sub

    Private Sub txtSecondaryQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQty, sender, e)
    End Sub

    Private Sub txtSecondaryQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQty, sender, e)
    End Sub

    'Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
    '    Try
    '        'yang bisa diedit adalah:
    '        '1. Qty
    '        '2. Unit Price
    '        '3. Discount
    '        '4. Remarks
    '        If Me.TxtRemark.Enabled = True Then
    '            If e.ColumnIndex = Me.ColTotalQty.Index Or _
    '            e.ColumnIndex = Me.colRemarks.Index Then
    '                If e.ColumnIndex = Me.ColSecQty.Index Then
    '                    'cek secondary qty
    '                    If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
    '                        e.Cancel = True
    '                        Exit Sub
    '                    End If
    '                End If
    '                'kalo sampe sini maka boleh edit
    '                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
    '            Else
    '                'jika tanpa po, maka unit price dan decimal juga bisa diedit.
    '                If Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.RowIndex).Value = -1 Then
    '                    If e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
    '                        'kalo sampe sini maka boleh edit
    '                        Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
    '                    Else
    '                        e.Cancel = True
    '                    End If
    '                Else
    '                    e.Cancel = True
    '                End If
    '            End If
    '        Else
    '            'hanya bisa edit pada saat detail aktif
    '            e.Cancel = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Unit Price
            '3. Discount
            '4. Remarks
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or _
                e.ColumnIndex = Me.colRemarks.Index Then
                    If e.ColumnIndex = Me.ColSecQty.Index Then
                        'cek secondary qty
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    'jika tanpa po, maka unit price dan decimal juga bisa diedit.
                    If Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.RowIndex).Value = -1 Then
                        If e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColDiscAmount.Index Then
                            'kalo sampe sini maka boleh edit
                            Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        e.Cancel = True
                    End If
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
   ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
   ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Double, _
   Optional ByVal QtyCheck As Boolean = False, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain, JmlPR, JmlPR_Transaksi As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard
        Dim PrimaryPO As Double = -1

        Try
            If Me.ObjProduct.ApakahGoods(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlObjLama = Me.ObjPurchaseReturnItem.GetJml2(Me.TxtPrimaryPurchaseReturn.Text, PrimaryProduct, _
                BatchNo, PrimaryLocation)
                Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
                JmlDGVTotal = Me.GetJmlDGV2(PrimaryProduct, PrimaryLocation, BatchNo)
                JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal

                If JmlSetelahEdit < 0 Then
                    JmlDGVLain = JmlDGVTotal - JmlDGV
                    If QtyCheck = False Then
                        If vShowMsg = True Then
                            MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                    Else
                        PrimaryPO = row.Cells(Me.ColPrimaryPO.Index).Value
                        JmlPR = Me.ObjPurchaseReturnItem.GetReturnBalance(PrimaryPO, PrimaryProduct, _
                               Me.txtPrimaryStockInPurchase.Text, Me.txtPrimaryGRN.Text)
                        If Me.baru = False Then
                            JmlPR_Transaksi = Me.ObjPurchaseReturnItem.GetJml(Me.TxtPrimaryPurchaseReturn.Text, PrimaryPO, PrimaryProduct, _
                            Me.txtPrimaryGRN.Text)
                        Else
                            JmlPR_Transaksi = 0
                        End If

                        If Me.dgvItem.Item(Me.ColPrimaryPO.Index, row.Index).Value <> -1 Then
                            If JmlGudang + JmlObjLama < JmlPR - JmlPR_Transaksi Then
                                If vShowMsg = True Then
                                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGVLain), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
                                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                                End If
                            Else
                                If vShowMsg = True Then
                                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlPR - JmlPR_Transaksi - JmlDGVLain), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
                                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                                End If
                            End If
                        Else
                            Dim Max_ As String = ((JmlGudang + JmlObjLama - JmlDGVLain) * CDec(row.Cells(Me.ColRatioReference.Index).Value)) / CDec(row.Cells(Me.ColRatioDefault.Index).Value)
                            Max_ = Me.ObjCD.RoundDown1(Max_, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitMulti.Index).Value))

                            If vShowMsg = True Then
                                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(Max_), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct, row.Cells(Me.ColPrimaryUnitMulti.Index).Value)) + " " + _
                                row.Cells(Me.ColUnitsCodeMulti.Index).Value.ToString + ".", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                    Return False
                End If

                'cek stock untuk tanggal yang bersangkutan
                JmlGudang = ObjSC.GetStock(PrimaryProduct, Me.dtpDate.Value)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlObjLama = Me.ObjPurchaseReturnItem.GetJml2(Me.TxtPrimaryPurchaseReturn.Text, PrimaryProduct, _
                BatchNo, PrimaryLocation)
                Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
                JmlDGVTotal = Me.GetJmlDGV2(PrimaryProduct, PrimaryLocation, BatchNo)
                JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
                If JmlSetelahEdit < 0 Then
                    JmlDGVLain = JmlDGVTotal - JmlDGV
                    Dim Max_ As String = ((JmlGudang + JmlObjLama - JmlDGVLain) * CDec(row.Cells(Me.ColRatioReference.Index).Value)) / CDec(row.Cells(Me.ColRatioDefault.Index).Value)
                    Max_ = Me.ObjCD.RoundDown1(Max_, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitMulti.Index).Value))

                    If QtyCheck = False Then
                        If vShowMsg = True Then
                            MsgBox("Your data change was cancelled, because it will make you stock quantity on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + "  minus.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                    Else
                        If vShowMsg = True Then
                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(Max_), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct, row.Cells(Me.ColPrimaryUnitMulti.Index).Value)) + " " + _
                            row.Cells(Me.ColUnitsCodeMulti.Index).Value + ".", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                    Return False
                End If
            End If

            'sampe sini 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function

    'Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
    '    Try
    '        Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
    '        Dim rounding As Integer = 0
    '        Dim Expiry As Date = #1/1/1950#

    '        'pastikan ada isinya
    '        With row
    '            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
    '                If e.ColumnIndex = Me.ColTotalQty.Index Then
    '                    'pastikan bilangan positif
    '                    If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                        MsgBox("Qty data must be a positive number", MsgBoxStyle.Information, "Information")
    '                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                        Exit Sub
    '                    End If

    '                    'digit grouoping
    '                    Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping( _
    '                    Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

    '                    'pastikan stock tidak menjadi minus
    '                    If IsDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value) = True Then
    '                        Expiry = Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value
    '                    Else
    '                        Expiry = #1/1/1950#
    '                    End If
    '                    If Me.ApakahEditSah(Me.dgvItem.Rows(e.RowIndex), Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value, -1, _
    '                    Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, _
    '                    Expiry, -1, Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value, True) = False Then
    '                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                        Exit Sub
    '                    End If

    '                    rounding = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
    '                    Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = _
    '                    Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value) * _
    '                    CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value), rounding))

    '                    Me.GabungkanBaris()
    '                    Me.HitungTotal_Multi()
    '                ElseIf e.ColumnIndex = Me.colRemarks.Index Then
    '                    'kalo remarks lebih besar dari 150 maka akan di potong
    '                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 150 Then
    '                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 150)
    '                    End If
    '                ElseIf e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
    '                    'pastikan bilangan positif
    '                    If e.ColumnIndex = Me.ColUnitPrice.Index Then
    '                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                            MsgBox("Qty data must be a positive number", MsgBoxStyle.Information, "Information")
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                            Exit Sub
    '                        End If
    '                    End If

    '                    If e.ColumnIndex = Me.ColDiscount.Index Then
    '                        rounding = 2
    '                    Else
    '                        rounding = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
    '                    End If

    '                    'digit grouoping
    '                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping( _
    '                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, rounding)

    '                    'jika discount maka tidak boleh lebih besar dari 100
    '                    If e.ColumnIndex = Me.ColDiscount.Index Then
    '                        If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
    '                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                            Exit Sub
    '                        End If
    '                    End If

    '                    'hitung sub total yang baru
    '                    Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value = _
    '                    Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) * _
    '                    (100 - CDec(Me.dgvItem.Item(Me.ColDiscount.Index, e.RowIndex).Value)) / 100, rounding))

    '                    Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value) * _
    '                    CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value), rounding)

    '                    Me.GabungkanBaris()
    '                    'hitung total yang baru
    '                    Me.HitungTotal_Multi()
    '                End If
    '            Else
    '                If e.ColumnIndex = Me.colRemarks.Index Then
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
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            Dim rounding As Integer = 0
            Dim Expiry As Date = #1/1/1950#

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                        'pastikan bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Qty data must be a positive number", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If

                        'digit grouoping
                        Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping( _
                        Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                        ''@rosadinnn
                        ''@MUaT
                        row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Cells(Me.ColTotalQtyMulti.Index).Value, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitMulti.Index).Value))

                        'pastikan stock tidak menjadi minus
                        If IsDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value) = True Then
                            Expiry = Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value
                        Else
                            Expiry = #1/1/1950#
                        End If

                        'hitung qty default
                        Me.SetQtyDefaultOnDgv(row.Index)

                        If Me.ApakahEditSah(Me.dgvItem.Rows(e.RowIndex), Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value, -1, _
                        Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, _
                        Expiry, -1, Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value, True) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.SetQtyDefaultOnDgv(row.Index)
                            Exit Sub
                        End If

                        rounding = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)

                        ''@rosadinnn
                        ''@MUaT
                        If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) = 0 Then
                            MsgBox("Qty restored to previous value, because qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.SetQtyDefaultOnDgv(row.Index)
                        End If

                        If Me.dgvItem.Item(Me.colPriceByMultiUnit.Index, e.RowIndex).Value = 1 Then
                            Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value) * _
                            CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, e.RowIndex).Value), rounding))
                        Else
                            Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value) * _
                            CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value), rounding))
                        End If


                        Me.GabungkanBaris()
                        Me.HitungTotal_Multi()
                    ElseIf e.ColumnIndex = Me.colRemarks.Index Then
                        'kalo remarks lebih besar dari 150 maka akan di potong
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 150 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 150)
                        End If
                    ElseIf e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColDiscAmount.Index Then
                        'pastikan bilangan positif
                        If e.ColumnIndex = Me.ColUnitPrice.Index Then
                            If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Qty data must be a positive number", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If

                            If Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = 0 Then
                                'pastikan nilainya lebih besar dari disc amount
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If

                        If e.ColumnIndex = Me.ColDiscount.Index Then
                            rounding = 2
                        Else
                            rounding = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                        End If

                        'jika discount maka tidak boleh lebih besar dari 100
                        If e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColDiscAmount.Index Then
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for discount.", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If

                            If e.ColumnIndex = Me.ColDiscount.Index Then
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                                    MsgBox("Maximum value is 100%.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                                row.Cells(Me.ColDiscFromPerct.Index).Value = "1"
                            Else
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                                    MsgBox("Maximum value is " & Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value), rounding) & ".", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                                row.Cells(Me.ColDiscFromPerct.Index).Value = "0"
                            End If
                        End If

                        'digit grouoping
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping( _
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, rounding)

                        'hitung sub total yang baru
                        Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) * _
                        (100 - CDec(Me.dgvItem.Item(Me.ColDiscount.Index, e.RowIndex).Value)) / 100, rounding))

                        ''@MUaT
                        'Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value) * _
                        'CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, e.RowIndex).Value), rounding)
                        Me.HitungItem(e.RowIndex, rounding)
                        Me.GabungkanBaris()
                        Me.HitungTotal_Multi()
                    End If
                Else
                    If e.ColumnIndex = Me.colRemarks.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

            ''@MUaT
            Me.SetQtyDefaultOnDgv(e.RowIndex)

        End Try
    End Sub

    Private Sub HitungItem(ByVal vRowIndex As Integer, ByVal rounding As Integer)
        Dim rowDgv As DataGridViewRow
        'Dim DefaultPrecisionRow As Integer = 0
        Dim DiscPerct As Integer = 0

        Try
            rowDgv = Me.dgvItem.Rows(vRowIndex)
            'DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(rowDgv.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            DiscPerct = rowDgv.Cells(Me.ColDiscFromPerct.Index).Value

            If rowDgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                rowDgv.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) * CDec(rowDgv.Cells(Me.ColDiscount.Index).Value))) / 100, rounding))
            Else
                rowDgv.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value)), 2))
            End If
            rowDgv.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(rowDgv.Cells(Me.ColUnitPrice.Index).Value) - CDec(rowDgv.Cells(Me.ColDiscAmount.Index).Value) _
            , rounding))


            If rowDgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
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

    Private Function ApakahAdaBarisYangSama() As Boolean
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= Me.dgvItem.RowCount - 2
                row = dgvItem.Rows(i)
                j = i + 1
                While j <= Me.dgvItem.RowCount - 1
                    row1 = dgvItem.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse

                    If row.Cells(Me.ColPONO.Index).Value = row1.Cells(Me.ColPONO.Index).Value And _
                    row.Cells(Me.ColCodeNo.Index).Value = row1.Cells(Me.ColCodeNo.Index).Value And _
                    row.Cells(Me.ColDescription.Index).Value = row1.Cells(Me.ColDescription.Index).Value And _
                    CStr(row.Cells(Me.ColBatchNo.Index).Value) = CStr(row1.Cells(Me.ColBatchNo.Index).Value) And _
                    row.Cells(Me.ColPrimaryLocation.Index).Value = row1.Cells(Me.ColPrimaryLocation.Index).Value And _
                    CDec(row1.Cells(Me.ColUnitPrice.Index).Value) = CDec(row.Cells(Me.ColUnitPrice.Index).Value) And _
                    CDec(row1.Cells(Me.ColDiscount.Index).Value) = CDec(row.Cells(Me.ColDiscount.Index).Value) Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = True Then
                            If row.Cells(Me.ColSecQty.Index).Value = row1.Cells(Me.ColSecQty.Index).Value Then
                                Return True
                            Else
                                Return False
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = False Then
                            Return True
                        End If
                    End If
                    j += 1
                End While
                i += 1
            End While

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' menggabungkan 2 baris yang sama dalam dgv menjadi 1 dengan menambahkan qty-nya
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GabungkanBaris()
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= Me.dgvItem.RowCount - 2
                row = dgvItem.Rows(i)
                j = i + 1
                While j <= Me.dgvItem.RowCount - 1
                    row1 = dgvItem.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse
                    If CStr(row.Cells(Me.ColPONO.Index).Value) = CStr(row1.Cells(Me.ColPONO.Index).Value) And _
                    CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
                    CStr(row.Cells(Me.ColBatchNo.Index).Value) = CStr(row1.Cells(Me.ColBatchNo.Index).Value) And _
                    CDec(row.Cells(Me.ColUnitPrice.Index).Value) = CDec(row1.Cells(Me.ColUnitPrice.Index).Value) And _
                    CDec(row.Cells(Me.ColDiscount.Index).Value) = CDec(row1.Cells(Me.ColDiscount.Index).Value) And _
                    row.Cells(Me.ColPrimaryLocation.Index).Value = row1.Cells(Me.ColPrimaryLocation.Index).Value And _
                    row.Cells(Me.ColPrimaryUnitMulti.Index).Value = row1.Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = True Then
                            If CDbl(row.Cells(Me.ColSecQty.Index).Value) = CDec(row1.Cells(Me.ColSecQty.Index).Value) Then
                                Me.GabungkanBaris_Multi(row, row1)
                                Me.ObjCD.SetNumber(Me.dgvItem)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = False Then
                            Me.GabungkanBaris_Multi(row, row1)
                            Me.ObjCD.SetNumber(Me.dgvItem)
                            Exit Sub
                        End If
                    Else
                        j += 1
                    End If
                End While
                i += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub GabungkanBaris_Default1(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow)
    '    Dim rounding As Integer = 0

    '    Try
    '        rounding = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)

    '        'gabungkan data qty.
    '        row.Cells(Me.ColTotalQty.Index).Value = CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(row1.Cells(Me.ColTotalQty.Index).Value)
    '        'gabungkan data remarks
    '        If CStr(row.Cells(Me.colRemarks.Index).Value).EndsWith(".") = False Then
    '            'tambahkan titik bila tidak diakhiri dengan titik
    '            row.Cells(Me.colRemarks.Index).Value = CStr(row.Cells(Me.colRemarks.Index).Value) + ". "
    '        End If
    '        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
    '                  CDec(row.Cells(Me.ColTotalQty.Index).Value), rounding)

    '        row.Cells(Me.colRemarks.Index).Value = CStr(row.Cells(Me.colRemarks.Index).Value) + CStr(row1.Cells(Me.colRemarks.Index).Value)
    '        'hitung ulang amount

    '        'hapus baris ybs
    '        Me.EraseRowIndex = row1.Index
    '        Me.BackgroundWorker1.RunWorkerAsync()
    '        'Me.dgvItem.Rows.RemoveAt(row1.Index)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub GabungkanBaris_Multi(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow)
        Dim rounding As Integer = 0

        Try
            rounding = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)

            'gabungkan data qty.
            row.Cells(Me.ColTotalQtyMulti.Index).Value = CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(row1.Cells(Me.ColTotalQtyMulti.Index).Value)
            Me.SetQtyDefaultOnDgv(row.Index)

            'gabungkan data remarks
            If CStr(row.Cells(Me.colRemarks.Index).Value).EndsWith(".") = False And CStr(row.Cells(Me.colRemarks.Index).Value) <> "" Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(Me.colRemarks.Index).Value = CStr(row.Cells(Me.colRemarks.Index).Value) + ". "
            Else
                If (row.Cells(Me.colRemarks.Index).Value) <> "" Then
                    row.Cells(Me.colRemarks.Index).Value = row.Cells(Me.colRemarks.Index).Value + " "
                End If
            End If

            If row.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                          CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), rounding)
            Else
                row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                          CDec(row.Cells(Me.ColTotalQty.Index).Value), rounding)
            End If

            If CStr(row1.Cells(Me.colRemarks.Index).Value) <> "" Then
                row.Cells(Me.colRemarks.Index).Value = CStr(row.Cells(Me.colRemarks.Index).Value) + CStr(row1.Cells(Me.colRemarks.Index).Value)
            End If
            row.Cells(Me.colRemarks.Index).Value = CStr(row.Cells(Me.colRemarks.Index).Value).Trim
            'hitung ulang amount

            'hapus baris ybs
            Me.EraseRowIndex = row1.Index
            Me.BackgroundWorker1.RunWorkerAsync()
            'Me.dgvItem.Rows.RemoveAt(row1.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SetQtyDefaultOnDgv(row.Index)
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Me.dgvItem.TabStop = True
        If Me.TxtRemark.Enabled = True Then
            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True
        End If
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        If Me.dgvItem.Rows.Count = 0 Then
            Me.dgvItem.TabStop = False
            If Me.TxtRemark.Enabled = True Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        End If
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
            Me.HitungTotal_Multi()
            Me.FormatTxtStockIN()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            Me.TxtRemark.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtRemark.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        Me.TxtCurrency.Focus()
                    Else
                        Me.cboAccount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            Me.TxtRemark.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            Me.TxtCodeNo.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            Me.TxtCodeNo.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.GotFocus
        Try
            Me.TxtQtyMulti.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.LostFocus
        Try
            Me.TxtQtyMulti.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.GotFocus
        Try
            Me.txtSecondaryQty.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.LostFocus
        Try
            Me.txtSecondaryQty.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        Try
            Me.TxtRemarkItem.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If My.Computer.Keyboard.ShiftKeyDown = True Then
                    If Me.TxtDiscAmount.ReadOnly = False Then
                        Me.TxtDiscAmount.Focus()
                    Else
                        Me.TxtQtyMulti.Focus()
                    End If
                Else
                    Me.BtnInput.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        Try
            Me.TxtRemarkItem.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.objPurchaseReturn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objPurchaseReturn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtPurchaseReturnNo.Text = Me.objPurchaseReturn.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindStockInPurchase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindStockInPurchase.Click
        Dim i As New Inventory_Module.FrmCariTransaksiStockInOut

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll
            i.JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            If Me.TxtStockInNo.Text <> "" Then
                i.txtCari.Text = Me.TxtStockInNo.Text
            End If
            i.ShowDialog()

            If i.PrimaryTransaksiStockInOut > 0 Then
                'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                Me.txtPrimaryStockInPurchase.Text = i.PrimaryTransaksiStockInOut
                Me.IsiDataStockIn()
                'kalo ada item maka stock in no tidak bisa diubah
                Me.FormatTxtStockIN()
                Me.TxtReasonOfReturn.Focus()
            Else
                Me.TxtStockInNo.Clear()
                Me.txtPrimaryStockInPurchase.Text = "-1"
                Me.IsiDataStockIn()
                Me.TxtStockInNo.Focus()
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FormatTxtStockIN()
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.TxtStockInNo.ReadOnly = True
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtStockInNo.TabStop = False
                    Me.btnFindStockInPurchase.Enabled = False
                    Me.TxtCurrency.ReadOnly = True
                    Me.TxtCurrency.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtCurrency.TabStop = False
                    Me.btnFindCurrency.Enabled = False
                Else
                    Me.TxtStockInNo.ReadOnly = False
                    Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                    Me.btnFindStockInPurchase.Enabled = True
                    Me.TxtStockInNo.TabStop = True

                    Me.TxtCurrency.ReadOnly = False
                    Me.TxtCurrency.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                    Me.TxtCurrency.TabStop = True
                    Me.btnFindCurrency.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtStockInNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStockInNo.KeyDown
        Dim i As New Inventory_Module.FrmCariTransaksiStockInOut
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase

        Try
            If e.KeyValue = 13 Then
                Dim PRimary As Double = Me.objInOut.GetPrimary(Me.TxtStockInNo.Text)
                If PRimary > -1 Then
                    Me.txtPrimaryStockInPurchase.Text = PRimary
                    Me.IsiDataStockIn()
                    'kalo ada item maka stock in no tidak bisa diubah
                    Me.FormatTxtStockIN()
                    Me.TxtReasonOfReturn.Focus()
                Else
                    Me.btnFindStockInPurchase.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtStockInNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.Leave
        Dim i As New Inventory_Module.FrmCariTransaksiStockInOut
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase

        Try
            If Me.TxtStockInNo.ReadOnly = False And Me.TxtRemark.Enabled = True Then
                If Me.TxtStockInNo.Text <> "" Then
                    Dim PRimary As Double = Me.objInOut.GetPrimary(Me.TxtStockInNo.Text)

                    If PRimary > -1 Then
                        Me.txtPrimaryStockInPurchase.Text = PRimary
                        Me.IsiDataStockIn()
                        'kalo ada item maka stock in no tidak bisa diubah
                        Me.FormatTxtStockIN()
                    Else
                        Me.btnFindStockInPurchase.PerformClick()
                    End If
                Else
                    Me.txtPrimaryStockInPurchase.Text = -1
                    Me.IsiDataStockIn()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub IsiDataStockIn()
        Try
            If Me.txtPrimaryStockInPurchase.Text = "" Or Me.txtPrimaryStockInPurchase.Text = "-1" Then
                Me.txtPrimaryStockInPurchase.Text = "-1"
                Me.TxtStockInNo.Text = ""
                Me.txtGRNNo.Clear()
                Me.txtPrimaryGRN.Text = "-1"
                Me.txtAopNo.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtSupplierDONumber.Clear()
            Else
                Me.TxtStockInNo.Text = Me.objInOut.GetNoTransaksi(Me.txtPrimaryStockInPurchase.Text)
                '  Dim PrimaryDO As Double = Me.ObjStockInOut.GetPrimaryDO(Me.txtPrimaryStockIn.Text)
                '    Me.txtPrimaryDO.Text = PrimaryDO
                Me.objGRN.GetInfo(Me.txtPrimaryStockInPurchase.Text, _
                Me.txtGRNNo.Text, Me.txtPrimaryGRN.Text, Me.txtAopNo.Text, _
                Me.txtPrimarySupplier.Text, Me.txtSupplierCode.Text, Me.TxtSupplierName.Text, _
                Me.txtSupplierDONumber.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJml(ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, ByVal vBatchNo As String, _
    ByVal vExpiry As Date, ByVal vSecQty As Double) As Double
        Dim row As DataGridViewRow
        Dim Qty As Double = 0
        Dim Lanjut As Boolean = False
        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And _
                row.Cells(Me.ColBatchNo.Index).Value = vBatchNo Then
                    'cek ed
                    If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                        If vExpiry = #1/1/1950# Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    Else
                        If row.Cells(Me.ColExpiryDate.Index).Value = vExpiry Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If
                    'cek sec qty
                    If Lanjut = True Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                            If row.Cells(Me.ColSecQty.Index).Value = vSecQty Then
                                Qty += row.Cells(Me.ColQty.Index).Value
                            End If
                        Else
                            If vSecQty = -1 Then
                                Qty += row.Cells(Me.ColQty.Index).Value
                            End If
                        End If
                    End If
                End If
            Next

            Return Qty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cboTipePurchaseReturn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipePurchaseReturn.SelectedIndexChanged
        Try
            If Me.cboTipePurchaseReturn.SelectedIndex = 0 Then
                'revisi aop
                Me.dgvItem.Columns(Me.ColUnitPrice.Index).Visible = False
                Me.dgvItem.Columns(Me.ColDiscount.Index).Visible = False
                Me.dgvItem.Columns(Me.ColNettPrice.Index).Visible = False
                Me.dgvItem.Columns(Me.ColAmount.Index).Visible = False
                Me.dgvItem.Columns(Me.ColExchangeRate.Index).Visible = False
                Me.dgvItem.Columns(Me.ColCurrency.Index).Visible = False
            Else
                'nota return
                Me.dgvItem.Columns(Me.ColUnitPrice.Index).Visible = True
                Me.dgvItem.Columns(Me.ColDiscount.Index).Visible = True
                Me.dgvItem.Columns(Me.ColNettPrice.Index).Visible = True
                Me.dgvItem.Columns(Me.ColAmount.Index).Visible = True
                Me.dgvItem.Columns(Me.ColExchangeRate.Index).Visible = True
                Me.dgvItem.Columns(Me.ColCurrency.Index).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub HitungTotal_Default()
    '    Dim rowdgv As DataGridViewRow
    '    Dim SubTotal, TaxAmount, Total As Double
    '    Dim Rounding As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
    '    Dim AmountInIDR As Decimal = 0
    '    Dim SubTotal1 As Decimal = 0
    '    Dim DefPrecIDR As Integer = 0

    '    Try
    '        DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency
    '        If Me.cboTipePurchaseReturn.SelectedIndex = 1 Then
    '            'hitung subtotal
    '            For Each rowdgv In Me.dgvItem.Rows
    '                SubTotal += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
    '                CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
    '                Rounding, MidpointRounding.AwayFromZero)
    '            Next
    '            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
    '            'hitung pajak
    '            TaxAmount = Me.ObjCD.DigitGrouping(SubTotal * CDec(Me.txtTax.Text) * CDec(Me.txtTaxExchangeRate.Text) / 100, _
    '            0, ClsCekData.EnumTipeRounding.Down)
    '            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, DefPrecIDR)
    '            'dutystamp
    '            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtDutystamp.Text), DefPrecIDR)
    '            If Me.TxtPrimaryCurrency.Text = 1 Then
    '                'jika rupiah, maka total adalah = sub total + tax + duty stamp
    '                Total = SubTotal + TaxAmount + CDec(Me.TxtDutystamp.Text)
    '                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Total, Rounding))
    '            Else
    '                Total = SubTotal
    '                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Total, Rounding))
    '            End If
    '        Else
    '            Me.txtSubTotalAfterDP.Text = 0
    '            Me.TxtDutystamp.Text = 0
    '            Me.TxtTotal.Text = 0
    '            Me.txtTaxExchangeRate.Text = 1
    '            Me.txtTaxAmount.Text = 0
    '            Me.txtTax.Text = 0
    '        End If

    '        'hitung purchase price yang akan digunakan untuk jurnal closing
    '        AmountInIDR = 0
    '        SubTotal1 = 0
    '        For Each rowdgv In Me.dgvItem.Rows
    '            SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
    '            CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
    '            Rounding, MidpointRounding.AwayFromZero)
    '        Next
    '        SubTotal = Math.Round(SubTotal1 * CDec(Me.txtTaxExchangeRate.Text), DefPrecIDR, MidpointRounding.AwayFromZero)

    '        For Each rowdgv In Me.dgvItem.Rows
    '            If rowdgv.Index = Me.dgvItem.Rows.Count - 1 Then
    '                rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = SubTotal - AmountInIDR
    '            Else
    '                If SubTotal1 <> 0 Then
    '                    rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
    '                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
    '                    DefPrecIDR, MidpointRounding.AwayFromZero)
    '                    AmountInIDR += CDec(rowdgv.Cells(Me.ColPurchasInIDR.Index).Value)
    '                End If
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub HitungTotal_Multi()
        Dim rowdgv As DataGridViewRow
        Dim SubTotal As Decimal ', TaxAmount, Total, SubtotalAfterDP, Dutystamp, Tax, TaxExchangeRate
        Dim Rounding As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecIDR As Integer = 0
        Dim DefPrecCur As Integer = 0
        Dim AmountInIDR As Decimal = 0
        Dim SubTotal1 As Decimal = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.HitungTotal_Multi_MiddlePoint()
                Exit Sub
            End If
            DefPrecCur = Me.ObjCurrency.GetDecimal_DefaultCurrency
            Dim DefPrecTax As Integer = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefPrecCur
            End If


            'If Me.cboTipePurchaseReturn.SelectedIndex = 1 Then
            'hitung subtotal
            SubTotal = 0
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                    SubTotal += ObjCD.DigitGrouping(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value), Rounding)
                Else
                    SubTotal += ObjCD.DigitGrouping(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value), Rounding)
                End If
            Next
            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))

            'hitung pajak
            If Me.TxtPrimaryCurrency.Text = "1" Then
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
            Else
                If Me.txtPrimaryCurrTax.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                          CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                Else

                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                           CDec(Me.txtSubTotalAfterDP.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                End If
            End If

            'dutystamp
            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutystamp.Text, DefPrecTax)
            'Else
            '    Me.txtSubTotalAfterDP.Text = 0
            '    Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
            '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
            '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, Rounding)
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Rounding)
            '    Me.txtTaxExchangeRate.Text = 1
            '    Me.txtTax.Text = 0
            'End If

            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.TxtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                    decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.TxtDutystamp.Text) Then
                    decTotal += CDec(Me.TxtDutystamp.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
            Else
                decTotal = 0
                If Me.TxtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.TxtDutystamp.Text) Then
                        decTotal += CDec(Me.TxtDutystamp.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
                Else
                    decTotal = 0
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
                End If
            End If

            'DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency

            'If Me.cboTipePurchaseReturn.SelectedIndex = 1 Then
            '    If IsNumeric(Me.TxtDutystamp.Text) Then
            '        Dutystamp = CDec(Me.TxtDutystamp.Text)
            '    Else
            '        Dutystamp = 0
            '    End If

            '    If IsNumeric(Me.txtTaxAmount.Text) Then
            '        TaxAmount = CDec(Me.txtTaxAmount.Text)
            '    Else
            '        TaxAmount = 0
            '    End If

            '    If IsNumeric(Me.txtTax.Text) Then
            '        Tax = CDec(Me.txtTax.Text)
            '    Else
            '        Tax = 0
            '    End If

            '    If IsNumeric(Me.txtTaxExchangeRate.Text) Then
            '        TaxExchangeRate = CDec(Me.txtTaxExchangeRate.Text)
            '    Else
            '        TaxExchangeRate = 0
            '    End If

            '    'hitung subtotal
            '    For Each rowdgv In Me.dgvItem.Rows
            '        SubTotal += Math.Round(CDbl(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
            '        CDbl(rowdgv.Cells(Me.ColQty.Index).Value) * CDbl(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
            '        Rounding, MidpointRounding.AwayFromZero)
            '    Next
            '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
            '    SubtotalAfterDP = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
            '    'hitung pajak

            '    Dim DefPrecTax As Integer = 0
            '    DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            '    If Me.chkTax.Checked = True Then
            '        If Me.cboPPNDipungu.SelectedIndex = 0 Then
            '            'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
            '            If Me.TxtPrimaryCurrency.Text = "1" Then
            '                TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
            '                SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), Me.ObjCurrency.GetDefaultPrecission(TxtPrimaryCurrency.Text))
            '            Else
            '                If Me.txtPrimaryCurrTax.Text = "1" Then
            '                    TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(TaxExchangeRate * Tax * _
            '                                          SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
            '                Else

            '                    TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
            '                                           SubtotalAfterDP / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
            '                End If
            '            End If
            '        End If
            '    End If
            'Else
            '    Me.TxtDutystamp.Text = 0
            '    Me.txtTaxAmount.Text = 0
            '    Me.txtTax.Text = 0
            'End If

            'Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Dutystamp, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(TaxExchangeRate, 2)
            'Me.txtTax.Text = Me.ObjCD.DigitGrouping(Tax, 2)

            'Dim decTotal As Decimal = 0
            ''hitung total = sub total + taxj 
            'If Me.TxtPrimaryCurrency.Text = "1" Then
            '    'kalo IDR, duty stamp diperhitungkan
            '    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
            '        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
            '    End If
            '    If IsNumeric(Me.txtTaxAmount.Text) Then
            '        decTotal += CDec(Me.txtTaxAmount.Text)
            '    End If
            '    If IsNumeric(Me.TxtDutystamp.Text) Then
            '        decTotal += CDec(Me.TxtDutystamp.Text)
            '    End If
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            'Else
            '    If Me.TxtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
            '        'kalo IDR, duty stamp diperhitungkan
            '        If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
            '            decTotal += CDec(Me.txtSubTotalAfterDP.Text)
            '        End If
            '        If IsNumeric(Me.txtTaxAmount.Text) Then
            '            decTotal += CDec(Me.txtTaxAmount.Text)
            '        End If
            '        If IsNumeric(Me.TxtDutystamp.Text) Then
            '            decTotal += CDec(Me.TxtDutystamp.Text)
            '        End If
            '        'kalo Currencynya sama diperhitungkan duty stamp dan tax
            '        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            '    Else
            '        'kalo IDR, duty stamp diperhitungkan
            '        If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
            '            decTotal += CDec(Me.txtSubTotalAfterDP.Text)
            '        End If
            '        'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
            '        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            '    End If
            'End If



            'hitung purchase price yang akan digunakan untuk jurnal closing
            AmountInIDR = 0
            SubTotal1 = 0
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                Else
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                End If
            Next
            SubTotal = Math.Round(SubTotal1 * CDec(Me.txtTaxExchangeRate.Text), DefPrecIDR, MidpointRounding.AwayFromZero)

            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Index = Me.dgvItem.Rows.Count - 1 Then
                    rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = SubTotal - AmountInIDR
                Else
                    If SubTotal1 <> 0 Then
                        If rowdgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                            rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        Else
                            rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        End If
                        AmountInIDR += CDec(rowdgv.Cells(Me.ColPurchasInIDR.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungTotal_Multi_MiddlePoint()
        Dim rowdgv As DataGridViewRow
        Dim SubTotal As Decimal ', TaxAmount, Total, SubtotalAfterDP, Dutystamp, Tax, TaxExchangeRate
        Dim Rounding As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
        Dim DefPrecIDR As Integer = 0
        Dim DefPrecCur As Integer = 0
        Dim AmountInIDR As Decimal = 0
        Dim SubTotal1 As Decimal = 0

        Try
            DefPrecCur = Me.ObjCurrency.GetDecimal_DefaultCurrency
            Dim DefPrecTax As Integer = 0
            If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            Else
                DefPrecTax = DefPrecCur
            End If


            'If Me.cboTipePurchaseReturn.SelectedIndex = 1 Then
            'hitung subtotal
            SubTotal = 0
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                    SubTotal += ObjCD.DigitGrouping(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value), Rounding)
                Else
                    SubTotal += ObjCD.DigitGrouping(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value), Rounding)
                End If
            Next
            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))

            'hitung pajak
            If Me.TxtPrimaryCurrency.Text = "1" Then
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
            Else
                If Me.txtPrimaryCurrTax.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                    CDec(Me.txtSubTotalAfterDP.Text) / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                Else

                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    CDec(Me.txtSubTotalAfterDP.Text) / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                End If
            End If

            'dutystamp
            Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutystamp.Text, DefPrecTax)
            'Else
            '    Me.txtSubTotalAfterDP.Text = 0
            '    Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
            '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(0, DefPrecTax)
            '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, Rounding)
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Rounding)
            '    Me.txtTaxExchangeRate.Text = 1
            '    Me.txtTax.Text = 0
            'End If

            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.TxtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                    decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.TxtDutystamp.Text) Then
                    decTotal += CDec(Me.TxtDutystamp.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
            Else
                decTotal = 0
                If Me.TxtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.TxtDutystamp.Text) Then
                        decTotal += CDec(Me.TxtDutystamp.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
                Else
                    decTotal = 0
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Rounding))
                End If
            End If

            'DefPrecIDR = Me.ObjCurrency.GetDecimal_DefaultCurrency

            'If Me.cboTipePurchaseReturn.SelectedIndex = 1 Then
            '    If IsNumeric(Me.TxtDutystamp.Text) Then
            '        Dutystamp = CDec(Me.TxtDutystamp.Text)
            '    Else
            '        Dutystamp = 0
            '    End If

            '    If IsNumeric(Me.txtTaxAmount.Text) Then
            '        TaxAmount = CDec(Me.txtTaxAmount.Text)
            '    Else
            '        TaxAmount = 0
            '    End If

            '    If IsNumeric(Me.txtTax.Text) Then
            '        Tax = CDec(Me.txtTax.Text)
            '    Else
            '        Tax = 0
            '    End If

            '    If IsNumeric(Me.txtTaxExchangeRate.Text) Then
            '        TaxExchangeRate = CDec(Me.txtTaxExchangeRate.Text)
            '    Else
            '        TaxExchangeRate = 0
            '    End If

            '    'hitung subtotal
            '    For Each rowdgv In Me.dgvItem.Rows
            '        SubTotal += Math.Round(CDbl(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
            '        CDbl(rowdgv.Cells(Me.ColQty.Index).Value) * CDbl(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
            '        Rounding, MidpointRounding.AwayFromZero)
            '    Next
            '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
            '    SubtotalAfterDP = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, Rounding))
            '    'hitung pajak

            '    Dim DefPrecTax As Integer = 0
            '    DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
            '    If Me.chkTax.Checked = True Then
            '        If Me.cboPPNDipungu.SelectedIndex = 0 Then
            '            'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
            '            If Me.TxtPrimaryCurrency.Text = "1" Then
            '                TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
            '                SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), Me.ObjCurrency.GetDefaultPrecission(TxtPrimaryCurrency.Text))
            '            Else
            '                If Me.txtPrimaryCurrTax.Text = "1" Then
            '                    TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(TaxExchangeRate * Tax * _
            '                                          SubtotalAfterDP / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
            '                Else

            '                    TaxAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Tax * _
            '                                           SubtotalAfterDP / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
            '                End If
            '            End If
            '        End If
            '    End If
            'Else
            '    Me.TxtDutystamp.Text = 0
            '    Me.txtTaxAmount.Text = 0
            '    Me.txtTax.Text = 0
            'End If

            'Me.TxtDutystamp.Text = Me.ObjCD.DigitGrouping(Dutystamp, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(TaxAmount, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(TaxExchangeRate, 2)
            'Me.txtTax.Text = Me.ObjCD.DigitGrouping(Tax, 2)

            'Dim decTotal As Decimal = 0
            ''hitung total = sub total + taxj 
            'If Me.TxtPrimaryCurrency.Text = "1" Then
            '    'kalo IDR, duty stamp diperhitungkan
            '    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
            '        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
            '    End If
            '    If IsNumeric(Me.txtTaxAmount.Text) Then
            '        decTotal += CDec(Me.txtTaxAmount.Text)
            '    End If
            '    If IsNumeric(Me.TxtDutystamp.Text) Then
            '        decTotal += CDec(Me.TxtDutystamp.Text)
            '    End If
            '    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            'Else
            '    If Me.TxtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
            '        'kalo IDR, duty stamp diperhitungkan
            '        If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
            '            decTotal += CDec(Me.txtSubTotalAfterDP.Text)
            '        End If
            '        If IsNumeric(Me.txtTaxAmount.Text) Then
            '            decTotal += CDec(Me.txtTaxAmount.Text)
            '        End If
            '        If IsNumeric(Me.TxtDutystamp.Text) Then
            '            decTotal += CDec(Me.TxtDutystamp.Text)
            '        End If
            '        'kalo Currencynya sama diperhitungkan duty stamp dan tax
            '        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            '    Else
            '        'kalo IDR, duty stamp diperhitungkan
            '        If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
            '            decTotal += CDec(Me.txtSubTotalAfterDP.Text)
            '        End If
            '        'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
            '        Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)))
            '    End If
            'End If



            'hitung purchase price yang akan digunakan untuk jurnal closing
            AmountInIDR = 0
            SubTotal1 = 0
            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                Else
                    SubTotal1 += Math.Round(CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                    CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value), _
                    Rounding, MidpointRounding.AwayFromZero)
                End If
            Next
            SubTotal = Math.Round(SubTotal1 * CDec(Me.txtTaxExchangeRate.Text), DefPrecIDR, MidpointRounding.AwayFromZero)

            For Each rowdgv In Me.dgvItem.Rows
                If rowdgv.Index = Me.dgvItem.Rows.Count - 1 Then
                    rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = SubTotal - AmountInIDR
                Else
                    If SubTotal1 <> 0 Then
                        If rowdgv.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                            rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColTotalQtyMulti.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        Else
                            rowdgv.Cells(Me.ColPurchasInIDR.Index).Value = Math.Round((CDec(rowdgv.Cells(Me.ColNettPrice.Index).Value) * _
                            CDec(rowdgv.Cells(Me.ColTotalQty.Index).Value) * CDec(rowdgv.Cells(Me.ColExchangeRate.Index).Value) * SubTotal) / SubTotal1, _
                            DefPrecIDR, MidpointRounding.AwayFromZero)
                        End If
                        AmountInIDR += CDec(rowdgv.Cells(Me.ColPurchasInIDR.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetExchangeRate(ByVal vTBL As DseTransaksiAOP.tbltrans_aop_item_exchangerateDataTable, _
   ByVal vTipePerhitungan As ClsTransaksiAOP.EnumCalculatesBy, ByVal vPrimaryPO As Double, _
   ByVal vPrimaryGRN As Double, ByVal vPrimaryAOP As Double) As Double
        Dim Filter As String = ""
        Dim row As DseTransaksiAOP.tbltrans_aop_item_exchangerateRow

        Try
            If vTipePerhitungan = ClsTransaksiAOP.EnumCalculatesBy.AOP Then
                If vTBL.Rows.Count > 0 Then
                    row = vTBL.Rows(0)
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            ElseIf vTipePerhitungan = ClsTransaksiAOP.EnumCalculatesBy.GRN Then
                Filter = "[Primary Transaksi Rujukan] = " + vPrimaryGRN.ToString
                vTBL.DefaultView.RowFilter = Filter
                If vTBL.DefaultView.Count > 0 Then
                    row = vTBL.DefaultView(0).Row
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            ElseIf vTipePerhitungan = ClsTransaksiAOP.EnumCalculatesBy.PO Then
                Filter = "[Primary Transaksi Rujukan] = " + vPrimaryPO.ToString
                vTBL.DefaultView.RowFilter = Filter
                If vTBL.DefaultView.Count > 0 Then
                    row = vTBL.DefaultView(0).Row
                    Return row.Exhange_Rate
                Else
                    Return 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtDutystamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.Enter
        Try
            If Me.TxtDutystamp.ReadOnly = False Then
                Me.ObjCD.Number_Enter(Me.TxtDutystamp, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDutystamp.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDutystamp, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDutystamp.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDutystamp, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                'If Me.TxtDutystamp.Focused = True Then
                '    If My.Computer.Keyboard.ShiftKeyDown = True Then
                '        If Me.txtTaxExchangeRate.TabStop = True Then
                '            Me.txtTaxExchangeRate.Focus()
                '        Else
                '            Me.chkTax.Focus()
                '        End If
                '    Else
                '        Me.TxtCodeNo.Focus()
                '    End If
                'End If
                Me.HitungTotal_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDutystamp.ReadOnly = False Then
                    Me.TxtDutystamp.BackColor = Me.BackColorAll
                Else
                    Me.TxtDutystamp.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDutystamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDutystamp.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtDutystamp.ReadOnly = False Then
                    Me.TxtDutystamp.BackColor = Me.TxtPrimaryCurrency.BackColor
                Else
                    Me.TxtDutystamp.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentTerm.ReadOnly = True Then
                    Me.TxtPaymentTerm.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPaymentTerm.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPaymentTerm.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtPaymentTerm, sender, e, 0, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaymentTerm.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPaymentTerm.ReadOnly = True Then
                    Me.TxtPaymentTerm.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtPaymentTerm.BackColor = Me.TxtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Leave
        Try
            If Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then

                ''@rosadinnn
                ''@MUaT
                Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                Me.SetQtyDefault()
                'Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))

            End If

            If Me.TxtQtyMulti.Focused = True And Me.TxtRemark.Enabled = True Then
                If My.Computer.Keyboard.ShiftKeyDown Then
                    If Me.cboLocation.Visible = True Then
                        Me.cboLocation.Focus()
                    ElseIf Me.cboBatchNo.Visible = True Then
                        Me.cboBatchNo.Focus()
                    ElseIf Me.cboCodeNo.Visible = True Then
                        Me.cboCodeNo.Focus()
                    Else
                        Me.TxtCodeNo.Focus()
                    End If
                Else
                    If Me.CboUnitMulti.Items.Count <= 1 Then
                        If Me.TxtUnitPrice.ReadOnly = False Then
                            Me.TxtUnitPrice.Focus()
                        Else
                            Me.TxtRemarkItem.Focus()
                        End If
                    Else
                        Me.CboUnitMulti.Focus()
                    End If
                End If
            End If
            Me.HitungSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmCetakTransaksiAOP

        Try
            If Me.TxtPrimaryPurchaseReturn.Text <> "" Then
                If IsNumeric(Me.TxtPrimaryPurchaseReturn.Text) = True Then
                    If CDbl(Me.TxtPrimaryPurchaseReturn.Text) >= 0 Then
                        'jika bukan quick purchase return maka tidak boleh print.
                        If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = False Then
                            MsgBox("Please print this data from purchase return form.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If

                        objPurchaseReturn.Cetak(ij.crv1, Me.TxtPrimaryPurchaseReturn.Text, True, Me.txtPrimarySupplier.Text)
                        ij.Text = "Purchase Return"
                        ij.ShowDialog()
                        Exit Sub
                    Else
                        Me.btnFind.PerformClick()
                    End If
                Else
                    Me.btnFind.PerformClick()
                End If
            Else
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryPurchaseReturn.Text <> "" And Me.TxtPrimaryPurchaseReturn.Text <> "-1" Then
                'jika quick purchase return maka tidak boleh diedit.
                If Me.objPurchaseReturn.Apakah_QuickPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text) = False Then
                    MsgBox("Please edit this data from purchase return form.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                objPurchaseReturn.Cetak(ij.crv1, Me.TxtPrimaryPurchaseReturn.Text, True, Me.txtPrimarySupplier.Text)
                ij.Text = "Purchase Return"
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyOnHand()
        Dim JmlGudang, JmlObjLama, JmlDGV As Decimal
        Dim Dec As Integer = 0
        Dim BatchNo As String = ""

        Try
            If Me.TxtRemark.Enabled = True And Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then
                Dec = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)

                'get qty on hand
                'MsgBox("Batch Number diasumsikan dari cbo. Harus update jika bisa return tanpa po.", MsgBoxStyle.Information, "Information")
                JmlGudang = Me.ObjProduct.GetJml3(Me.txtPrimaryCodeNo.Text, -1, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                If Me.baru = False Then
                    JmlObjLama = Me.ObjPurchaseReturnItem.GetJml2(Me.TxtPrimaryPurchaseReturn.Text, Me.txtPrimaryCodeNo.Text, _
                    Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                Else
                    JmlObjLama = 0
                End If
                JmlDGV = Me.GetJmlDGV2(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, _
                Me.cboBatchNo.Text)

                Dim Max_ As String = ((JmlGudang + JmlObjLama - JmlDGV) * CDec(Me.TxtRatioReference.Text)) / CDec(Me.TxtRatioDefault.Text)
                Max_ = Me.ObjCD.RoundDown1(Max_, Dec)

                Me.TxtQtyOnHand.Text = Me.ObjCD.DigitGrouping(Max_, Dec)
                If Me.CboUnitMulti.Visible = True Then
                    Me.LblQtyOnHand.Text = "Qty On Hand (" + Me.CboUnitMulti.Text + ")"
                Else
                    Me.LblQtyOnHand.Text = "Qty On Hand (" + Me.LblUnitsDefault.Text + ")"
                End If
            Else
                Me.TxtQtyOnHand.Text = "0"
                Me.LblQtyOnHand.Text = "Qty On Hand"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNo.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.cboBatchNo.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        If Me.cboPO.Visible = False Then
                            Me.TxtCodeNo.Focus()
                        Else
                            Me.cboPO.Focus()
                        End If
                    Else
                        If Me.cboLocation.Visible = True Then
                            Me.cboLocation.Focus()
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNo.SelectedValueChanged
        Try
            If Me.chkLocationControlled.Checked = True Then
                If Me.cboBatchNo.Text <> "" Then
                    'isikan data lokasi
                    Me.IsiCboLokasi(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryPurchaseReturn.Text, Me.cboBatchNo.Text)
                End If
            Else
                'kosongkan data lokasi
                Me.cboLocation.DataSource = Nothing
                Me.cboLocation.Items.Clear()
            End If
            'isikan data expiry date 
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.DtpExpired.Value = Me.ObjProduct.GetExpiryDate(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            Else
                If Me.DtpExpired.Visible = True Then
                    Me.DtpExpired.Value = Today
                Else
                    Me.DtpExpired.Value = #1/1/1950#
                End If
            End If

            If Me.cboBatchNo.SelectedIndex <> -1 And Me.TxtRemark.Enabled = True Then
                Me.SetQtyOnHand()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboLokasi(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboLocation(vPrimaryProduct, Me.cboLocation, _
            Me.TBL2, vBatchNo)
            Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item6Row
            Dim row2 As dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjPurchaseReturnItem.IsikanTBLLocation(Me.TBL5, vPrimaryProduct, vPrimaryTransaksi, vBatchNo)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL5.Rows.Count > 0
                    row1 = Me.TBL5.Rows(0)
                    For Each row2 In Me.TBL2.Rows
                        If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL2.Rows.Add(row1.Lot_No, row1.Location, row1.Qty)
                    End If
                    Me.TBL5.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLLocation(vPrimaryProduct, vBatchNo)
            While Me.TBL5.Rows.Count > 0
                row1 = Me.TBL5.Rows(0)
                For Each row2 In Me.TBL2.Rows
                    If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL2.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL2.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL5.Rows.Remove(row1)
            End While

            Me.TBL2.DefaultView.Sort = "Location ASC"
            Me.cboLocation.DataSource = Me.TBL2.DefaultView
            Me.cboLocation.ValueMember = "Location"
            Me.cboLocation.DisplayMember = "Location"
            Me.cboLocation.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLLocation(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item6Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And CStr(row.Cells(Me.ColBatchNo.Index).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL5.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColBatchNo.Index).Value) And row1.Location = row.Cells(Me.ColLocation.Index).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL5.Rows.Add(CStr(row.Cells(Me.ColBatchNo.Index).Value), row.Cells(Me.ColLocation.Index).Value, CDbl(row.Cells(Me.ColTotalQty.Index).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocation_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocation.Enter

    End Sub

    Private Sub cboLocation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocation.GotFocus

    End Sub

    Private Sub cboLocation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocation.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.cboLocation.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = False Then
                        Me.TxtQtyMulti.Focus()
                    Else
                        If Me.cboBatchNo.Visible = True Then
                            Me.cboBatchNo.Focus()
                        ElseIf Me.cboCodeNo.Visible = True Then
                            Me.cboCodeNo.Focus()
                        Else
                            Me.TxtCodeNo.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocation_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocation.SelectedValueChanged
        Try
            If Me.cboLocation.Text = "" Or objLoc.ApakahSudahAda(Me.cboLocation.Text) = False Then
                Me.txtPrimaryLocation.Text = "-1"
            Else
                Me.txtPrimaryLocation.Text = objLoc.GetPrimary(Me.cboLocation.Text)
            End If

            If Me.TxtRemark.Enabled = True And Me.cboLocation.SelectedIndex <> -1 Then
                Me.SetQtyOnHand()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrency_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.Enter
        Try
            Me.buffer = Me.TxtCurrency.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCurrency_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtCurrency.ReadOnly = False And Me.TxtCurrency.Enabled = True Then
                TxtCurrency.BackColor = BackColorAll
            Else
                If Me.TxtCurrency.ReadOnly = True Then
                    Me.TxtCurrency.BackColor = Me.txtReadOnly.BackColor
                Else
                    TxtCurrency.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                End If
            End If
        End If
    End Sub

    Private Sub TxtCurrency_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtCurrency.ReadOnly = False And Me.TxtCurrency.Enabled = True Then
                TxtCurrency.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub TxtCurrency_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCurrency.TextChanged
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtCurrency.ReadOnly = False And Me.TxtCurrency.Enabled = True Then
                TxtCurrency.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub TxtCurrency_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCurrency.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrency.Text) = True Then
                    Me.TxtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.TxtCurrency.Text)
                    Me.chkTax.Focus()
                Else
                    If Me.TxtCurrency.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrency.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.TxtCurrency.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                        Me.chkTax.Focus()
                    Else
                        Me.TxtCurrency.Clear()
                        Me.TxtPrimaryCurrency.Text = "-1"
                        Me.TxtCurrency.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCurrency_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrency.Leave
        Dim i As New FrmCariMasterCurrency
        Dim DefPrec As Integer = 0
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtCurrency.Text <> "" Then
                    If ObjCurrency.ApakahSudahAda(Me.TxtCurrency.Text) = True Then
                        Me.TxtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.TxtCurrency.Text)

                        If Me.TxtCurrency.Focused = True Then
                            If My.Computer.Keyboard.ShiftKeyDown = True Then
                                Me.TxtReasonOfReturn.Focus()
                            Else
                                Me.TxtRemark.Focus()
                            End If
                        End If

                        Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                        DefPrec = Me.DefaultPrecision
                        For Each rowDGV In Me.dgvItem.Rows
                            Me.HitungItem(rowDGV.Index, DefPrec)
                        Next
                        If Me.buffer <> Me.TxtCurrency.Text Then
                            Me.SetTaxRate(True)
                        End If
                        Me.HitungTotal_Multi()
                    Else
                        If Me.TxtCurrency.Text <> "" Then
                            i.txtCari.Text = Me.TxtCurrency.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryCurrency > 0 Then
                            Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                            Me.TxtCurrency.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)

                            If Me.TxtCurrency.Focused = True Then
                                If My.Computer.Keyboard.ShiftKeyDown = True Then
                                    Me.TxtReasonOfReturn.Focus()
                                Else
                                    Me.TxtRemark.Focus()
                                End If
                            End If

                            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                            DefPrec = Me.DefaultPrecision
                            For Each rowDGV In Me.dgvItem.Rows
                                Me.HitungItem(rowDGV.Index, DefPrec)
                            Next
                            If Me.buffer <> Me.TxtCurrency.Text Then
                                Me.SetTaxRate(True)
                            End If
                            Me.HitungTotal_Multi()
                        Else
                            Me.TxtCurrency.Clear()
                            Me.TxtPrimaryCurrency.Text = "-1"
                            Me.TxtCurrency.Focus()
                            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                        End If
                    End If
                Else
                    Me.TxtPrimaryCurrency.Text = "-1"
                    MsgBox("Please speficy currency.", MsgBoxStyle.Information, "Information")
                    Me.TxtCurrency.Focus()
                    Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetTaxRate(ByVal vGetTaxRateFromDB As Boolean)
        Dim Rate As Decimal = 0

        Try
            If Me.TxtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(1, Me.ObjCurrency.GetDecimal_DefaultCurrency)
                Me.txtTaxExchangeRate.ReadOnly = True
                Me.txtTaxExchangeRate.TabStop = False
                Me.txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
            Else
                Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryCurrency.BackColor
                If vGetTaxRateFromDB = True Then
                    Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.TxtPrimaryCurrency.Text)
                    If Rate = 0 Then
                        Rate = 1
                    End If
                    Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, _
                    Me.ObjTaxRate.DefaultPrec)
                End If
                Me.txtTaxExchangeRate.ReadOnly = False
                Me.txtTaxExchangeRate.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.TxtCurrency.Text) = True Then
                Me.chkTax.Focus()
                Me.TxtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.TxtCurrency.Text)
                Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                Me.SetTaxRate(True)
                Me.SetDefaultPrecision()
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.TxtCurrency.Text <> "" Then
                    i.txtCari.Text = Me.TxtCurrency.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.TxtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.TxtCurrency.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                    Me.SetTaxRate(True)
                    Me.SetDefaultPrecision()
                    Me.chkTax.Focus()
                Else
                    Me.TxtCurrency.Clear()
                    Me.TxtCurrency.Focus()
                    Me.SetTaxRate(True)
                    Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetDefaultPrecision(Optional ByVal vHitungTotal As Boolean = True)
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtPrimaryCurrency.Text = "1" Then
                Me.txtTaxExchangeRate.Text = "1"
                Me.txtTaxExchangeRate.ReadOnly = True
                Me.txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
            Else
                Me.txtTaxExchangeRate.ReadOnly = False
                Me.txtTaxExchangeRate.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)

            If vHitungTotal = True Then
                For Each rowDGV In Me.dgvItem.Rows
                    rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency.Text
                    Me.HitungItem(rowDGV.Index, Me.DefaultPrecision)
                Next
                Me.HitungTotal_Multi()
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

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
                TxtUnitPrice.BackColor = BackColorAll
            Else
                TxtUnitPrice.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
                TxtUnitPrice.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
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
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
                TxtDiscount.BackColor = BackColorAll
            Else
                TxtDiscount.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
                TxtDiscount.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
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
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtdiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtdiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungSubTotal()
        Try
            Dim Decimal_Def As Integer = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text)

            If Me.TxtDiscFromPerct.Text = "" Then
                Me.TxtDiscFromPerct.Text = "1"
            End If
            If Me.TxtDiscFromPerct.Text = "1" Then
                'hitung disc amount
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtUnitPrice.Text) * CDec(Me.TxtDiscount.Text)) / 100, Decimal_Def)
            Else
                'hitung disc %
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtDiscAmount.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2)
            End If

            Me.TxtNettPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount.Text), Decimal_Def)
            'Obs
            'Me.TxtNettPrice.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.TxtUnitPrice.Text) * _
            '(100 - CDec(Me.TxtDiscount.Text)) / 100, Decimal_Def))

            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.TxtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtNettPrice.Text) * _
                CDec(Me.TxtQtyMulti.Text), Decimal_Def)
            Else
                Me.TxtAmount.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtNettPrice.Text) * _
                CDec(Me.TxtQtyDefault.Text), Decimal_Def)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
                If CDec(Me.TxtUnitPrice.Text) = 0 Then
                    Me.TxtDiscount.Text = 0
                    Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
                End If
                If CDec(Me.TxtDiscount.Text) <> CDec(Me.buffer) Then
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.HitungSubTotal()
                End If

                If Me.TxtDiscount.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtDiscAmount.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtUnitPrice, sender, e, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text), Me.ObjCD)
            'unit price harus selalu lebih besar dari disc amount
            'hanya berlaku kalo ngitung disc berdasarkan disc amount, klo berdasarkan perc, langsung hitung ulang disc amountnya
            If Me.TxtDiscFromPerct.Text = "0" Then
                If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtUnitPrice.Text = Me.buffer
                    Exit Sub
                End If
            End If
            'Obs
            'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, _
            'Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))

            If Me.TxtRemark.Enabled = True Then
                If Me.TxtUnitPrice.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        If Me.CboUnitMulti.Visible = True Then
                            Me.CboUnitMulti.Focus()
                        Else
                            Me.TxtQtyMulti.Focus()
                        End If
                    Else
                        Me.TxtDiscount.Focus()
                    End If
                End If

                Me.HitungSubTotal()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTax(ByVal vGetTaxRateFromDB As Boolean)
        Dim SettingValue As String = ""

        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.chkTax.Checked = True Then
                    Me.txtTax.BackColor = Me.txtPrimaryCodeNo.BackColor
                    Me.txtTax.TabStop = True
                    Me.txtTax.ReadOnly = False
                    SettingValue = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(SettingValue) = True Then
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping(CDec(SettingValue), 2)
                    Else
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                    End If
                    Me.txtTax.TabStop = True
                Else
                    Me.txtTax.TabStop = False
                    Me.txtTax.ReadOnly = True
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                    Me.txtTax.BackColor = Me.txtReadOnly.BackColor
                End If

                Me.SetTaxRate(vGetTaxRateFromDB)
                Me.HitungTotal_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                SetTax(True)
                Me.HitungTotal_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
            If Me.TxtRemark.Enabled = True Then
                If Me.txtTax.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        Me.chkTax.Focus()
                    Else
                        If Me.txtTaxExchangeRate.ReadOnly = False Then
                            Me.txtTaxExchangeRate.Focus()
                        Else
                            Me.TxtDutystamp.Focus()
                        End If
                    End If
                End If

                Me.ObjCD.Number_Leave(Me.txtTax, sender, e, 2, Me.ObjCD)
                Me.HitungTotal_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Enter
        Try
            Try
                If Me.TxtRemark.Enabled = True Then
                    If Me.txtTaxExchangeRate.ReadOnly = False Then
                        Me.ObjCD.Number_Enter(Me.txtTaxExchangeRate, sender, e)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If My.Computer.Keyboard.ShiftKeyDown = True Then
                    If Me.txtTax.ReadOnly = False Then
                        Me.txtTax.Focus()
                    Else
                        Me.chkTax.Focus()
                    End If
                Else
                    Me.TxtDutystamp.Focus()
                End If

                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.txtTaxExchangeRate.Text, 2)
                Me.HitungTotal_Multi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtTaxExchangeRate, sender, e)
    End Sub

    Private Sub txtTaxExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtTaxExchangeRate, sender, e, False)
    End Sub

    Private Sub txtTax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtTax, sender, e, 100)
    End Sub

    Private Sub txtTax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtTax, sender, e)
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtTax.ReadOnly = False Then
                If Me.txtTax.ReadOnly = False And Me.txtTax.Enabled = True Then
                    txtTax.BackColor = BackColorAll
                Else
                    txtTax.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                End If
            Else
                txtTax.BackColor = Me.txtReadOnly.BackColor
            End If
        End If
    End Sub

    Private Sub txtTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtTax.ReadOnly = False And Me.txtTax.Enabled = True Then
                txtTax.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub txtTaxExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtTaxExchangeRate.ReadOnly = False Then
                If Me.txtTaxExchangeRate.ReadOnly = False And Me.txtTaxExchangeRate.Enabled = True Then
                    txtTaxExchangeRate.BackColor = BackColorAll
                Else
                    txtTaxExchangeRate.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                End If
            Else
                txtTaxExchangeRate.BackColor = Me.txtReadOnly.BackColor
            End If
        End If
    End Sub

    Private Sub txtTaxExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtTaxExchangeRate.ReadOnly = False And Me.txtTaxExchangeRate.Enabled = True Then
                txtTaxExchangeRate.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub chkTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTax.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.chkTax.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        Me.TxtCurrency.Focus()
                    Else
                        If Me.txtTax.TabStop = True Then
                            Me.txtTax.Focus()
                        Else
                            Me.TxtDutystamp.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        Dim ApakahLokal As Boolean = False
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjSup.GetPrimary(Me.txtSupplierCode.Text)
                If Primary <> -1 Then
                    Me.txtPrimarySupplier.Text = Primary

                    Me.txtSupplierCode.Text = Me.ObjSup.GetSupplierCode(Me.txtPrimarySupplier.Text)
                    Me.TxtSupplierName.Text = Me.ObjSup.GetSupplierName(Me.txtPrimarySupplier.Text)
                    If Me.TxtCurrency.ReadOnly = False Then
                        Me.TxtPrimaryCurrency.Text = Me.ObjSup.GetPrimaryDefaultCurrency(Me.txtPrimarySupplier.Text)
                        Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    End If
                    'Update NPWP 16 Digit
                    Me.TxtNPWPNo.Text = Me.ObjSup.GetNPWP(Me.txtPrimarySupplier.Text)
                    '2025 Core Tax Tahap 2.1
                    If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                        Me.TxtNPWPNo.Text = ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
                    End If
                Else
                    Me.btnFindSupplier.PerformClick()
                End If

                If Me.dgvItem.RowCount = 0 Then
                    If Me.txtSupplierCode.Text <> "" Then
                        Me.TxtStockInNo.ReadOnly = True
                        Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                        Me.TxtStockInNo.TabStop = False
                        Me.btnFindStockInPurchase.Enabled = False
                    Else
                        Me.TxtStockInNo.ReadOnly = False
                        Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                        Me.TxtStockInNo.TabStop = True
                        Me.btnFindStockInPurchase.Enabled = True
                    End If
                End If

                If Me.TxtRemark.Enabled = True Then
                    If Me.txtSupplierCode.Focused = True Then
                        If My.Computer.Keyboard.ShiftKeyDown = True Then
                            If Me.TxtStockInNo.ReadOnly = False Then
                                Me.TxtStockInNo.Focus()
                            Else
                                Me.dtpDate.Focus()
                            End If
                        Else
                            Me.TxtReasonOfReturn.Focus()
                        End If
                    End If
                End If
            End If

            Me.SetTax(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If Me.txtSupplierCode.Text <> "" Then
                Dim Primary As Double = ObjSup.GetPrimary(Me.txtSupplierCode.Text)
                If Primary <> -1 Then
                    Me.txtPrimarySupplier.Text = Primary
                    Me.txtSupplierCode.Text = Me.ObjSup.GetSupplierCode(Me.txtPrimarySupplier.Text)
                    Me.TxtSupplierName.Text = Me.ObjSup.GetSupplierName(Me.txtPrimarySupplier.Text)
                    If Me.TxtCurrency.ReadOnly = False Then
                        Me.TxtPrimaryCurrency.Text = Me.ObjSup.GetPrimaryDefaultCurrency(Me.txtPrimarySupplier.Text)
                        Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                    End If
                    'Update NPWP 16 Digit
                    Me.TxtNPWPNo.Text = Me.ObjSup.GetNPWP(Me.txtPrimarySupplier.Text)
                    '2025 Core Tax Tahap 2.1
                    If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                        Me.TxtNPWPNo.Text = ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
                    End If
                Else
                    Me.btnFindSupplier.PerformClick()
                End If
            Else
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                'Update NPWP 16 Digit
                Me.TxtNPWPNo.Clear()
            End If

            If Me.dgvItem.RowCount = 0 Then
                If Me.txtSupplierCode.Text <> "" Then
                    Me.TxtStockInNo.ReadOnly = True
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtStockInNo.TabStop = False
                    Me.btnFindStockInPurchase.Enabled = False
                Else
                    Me.TxtStockInNo.ReadOnly = False
                    Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                    Me.TxtStockInNo.TabStop = True
                    Me.btnFindStockInPurchase.Enabled = True
                End If
            End If

            If Me.TxtRemark.Enabled = True Then
                If Me.txtSupplierCode.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        If Me.TxtStockInNo.ReadOnly = False Then
                            Me.TxtStockInNo.Focus()
                        Else
                            Me.dtpDate.Focus()
                        End If
                    Else
                        Me.TxtReasonOfReturn.Focus()
                    End If
                End If
            End If

            Me.SetTax(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        '2025 Core Tax Tahap 2.1
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.txtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = i.PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSup.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSup.GetSupplierName(Me.txtPrimarySupplier.Text)
                'Update NPWP 16 Digit
                Me.TxtNPWPNo.Text = Me.ObjSup.GetNPWP(Me.txtPrimarySupplier.Text)
                '2025 Core Tax Tahap 2.1
                If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                    Me.TxtNPWPNo.Text = ObjSetNPWP.format16DigitSpasi(Me.TxtNPWPNo.Text)
                End If
                If Me.TxtCurrency.ReadOnly = False Then
                    Me.TxtPrimaryCurrency.Text = Me.ObjSup.GetPrimaryDefaultCurrency(Me.txtPrimarySupplier.Text)
                    Me.TxtCurrency.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency.Text)
                End If
            Else
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                'Update NPWP 16 Digit
                Me.TxtNPWPNo.Clear()
                Me.txtSupplierCode.Focus()
            End If

            If Me.dgvItem.RowCount = 0 Then
                If Me.txtSupplierCode.Text <> "" Then
                    Me.TxtStockInNo.ReadOnly = True
                    Me.TxtStockInNo.BackColor = Me.txtReadOnly.BackColor
                    Me.TxtStockInNo.TabStop = False
                    Me.btnFindStockInPurchase.Enabled = False
                Else
                    Me.TxtStockInNo.ReadOnly = False
                    Me.TxtStockInNo.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                    Me.TxtStockInNo.TabStop = True
                    Me.btnFindStockInPurchase.Enabled = True
                End If
            End If

            Me.SetTax(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtSupplierCode.ReadOnly = False And Me.txtSupplierCode.Enabled = True Then
                txtSupplierCode.BackColor = BackColorAll
            Else
                txtSupplierCode.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtSupplierCode.ReadOnly = False And Me.txtSupplierCode.Enabled = True Then
                txtSupplierCode.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub txtSupplierDONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierDONumber.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtSupplierDONumber.ReadOnly = False And Me.txtSupplierDONumber.Enabled = True Then
                txtSupplierDONumber.BackColor = BackColorAll
            Else
                If Me.txtSupplierDONumber.ReadOnly = True Then
                    txtSupplierDONumber.BackColor = Me.txtReadOnly.BackColor
                Else
                    txtSupplierDONumber.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
                End If

            End If
        End If
    End Sub

    Private Sub txtSupplierDONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierDONumber.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.txtSupplierDONumber.ReadOnly = False And Me.txtSupplierDONumber.Enabled = True Then
                txtSupplierDONumber.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub


    'Private Sub SetSizeTxt()
    '    Dim LeftAcuan As Double = 0

    '    Try
    '        Me.cboBatchNo.Left = Me.TxtReturnBalance.Left
    '        Me.cboBatchNo.Width = Me.TxtReturnBalance.Width

    '        LeftAcuan = Me.SplitContainer4.SplitterDistance + Me.SplitContainer4.SplitterWidth
    '        Me.LblExpiry.Left = LeftAcuan + Me.LblQtyOnHand.Left
    '        Me.lblExpiryDate.Left = LeftAcuan + Me.TxtQtyOnHand.Left
    '        Me.DtpExpired.Left = LeftAcuan + Me.TxtQtyOnHand.Left
    '        Me.DtpExpired.Width = Me.TxtQtyOnHand.Width

    '        Me.TxtUnitPrice.Left = Me.TxtNettPrice.Left
    '        Me.TxtUnitPrice.Width = Me.TxtNettPrice.Width

    '        LeftAcuan = Me.SplitContainer6.SplitterDistance + Me.SplitContainer6.SplitterWidth
    '        Me.LblDiscount.Left = LeftAcuan + Me.LblAmount.Left
    '        Me.TxtDiscount.Left = LeftAcuan + Me.TxtAmount.Left
    '        Me.TxtDiscount.Width = Me.TxtAmount.Width
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub BtnInput_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnInput.Leave
        Try
            If Me.mFromApakahDataBenar = True Then
                If My.Computer.Keyboard.ShiftKeyDown = False Then
                    If Me.TxtCodeNo.Enabled = True Then
                        Me.TxtCodeNo.Focus()
                    End If
                Else
                    Me.TxtRemarkItem.Focus()
                End If
            Else
                Me.mFromApakahDataBenar = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub PrepareMultiUnit(ByVal vPrimaryProduct As Double)
        Dim dtaMultiUnit As New DseTransaksiPOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim objUnit As New Core_Module.ClsMasterUnit
        Try
            Me.CboUnitMulti.DataSource = Nothing
            If Me.ObjPO.IsMultiUnitProduct(vPrimaryProduct) = False Then
                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtPrimaryUnitsMulti.Text = ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
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
                    Me.LblUnitPrice.Text = "Unit Price"
                End If
                If Me.LblUnitsDefault.Text = "" Then
                    Me.LblUnitsDefault.Text = "Units"
                End If

            Else
                Me.LblUnitsMulti.Visible = False
                Me.CboUnitMulti.Visible = True
                Me.CboUnitMulti.TabStop = True
                dtaMultiUnit.Connection.ConnectionString = Core_Module.GetConstring
                dtaMultiUnit.FillBy_PrimaryProduct(Me.tblMultiUnit, Me.txtPrimaryCodeNo.Text)
                Me.ObjPO.IsiCboUnit(vPrimaryProduct, Me.CboUnitMulti, Me.tblCboUnit, Me.tblMultiUnit, Me.TxtPrimaryUnitsMulti)
                Me.CboUnitMulti.SelectedIndex = -1
                Me.CboUnitMulti.SelectedIndex = 0
            End If

            If Me.txtPrimaryCodeNo.Text <> -1 Then
                If Me.CboUnitMulti.Visible = True Then
                    Me.LblQtyOnHand.Text = "Qty On Hand (" + Me.CboUnitMulti.Text + ")"
                Else
                    Me.LblQtyOnHand.Text = "Qty On Hand (" + Me.LblUnitsDefault.Text + ")"
                End If
            Else
                Me.LblQtyOnHand.Text = "Qty On Hand"

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SetQtyDefault()
        Try
            Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboUnitMulti_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnitMulti.Leave
        Try
            If Me.CboUnitMulti.Text <> Me.cboPriceByMultiUnit.Text Then
                Me.txtPriceByMultiUnit.Text = 0
            Else
                Me.txtPriceByMultiUnit.Text = 1
            End If
            Me.HitungSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboUnitMulti_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnitMulti.SelectedIndexChanged
        'Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Dim row As DseTransaksiPO.CboUnitRow
        Try
            If Me.CboUnitMulti.SelectedIndex = -1 Then
                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtPrimaryUnitsMulti.Text = ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
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
                Me.TxtRatioReference.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioReference(Me.txtPrimaryCodeNo.Text, row.Primary), ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, row.Primary))
                Me.TxtRatioDefault.Text = Me.ObjCD.DigitGrouping(Me.ObjPO.GetRatioDefault(Me.txtPrimaryCodeNo.Text, row.Primary), ObjProduct.GetAngkaDecimal(row.Primary))
                Me.TxtPrimaryUnitsMulti.Text = row.Primary
                Me.TxtQtyMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                Me.TxtPrimaryUnitsDefault.Text = ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
                Me.LblUnitsMulti.Text = Me.CboUnitMulti.Text
                Me.LblUnitsDefault.Text = objUnit.GetUnitCode(Me.TxtPrimaryUnitsDefault.Text)
                Me.LblUnitsMulti.Visible = False
                Me.CboUnitMulti.Visible = True
                Me.CboUnitMulti.TabStop = True
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnitMulti.Text
            End If
            Me.SetQtyDefault()
            Me.SetQtyOnHand()
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
                    Me.txtPriceByMultiUnit.Text = "1"
                    Me.cboPriceByMultiUnit.Left = Me.LblUnitPrice.Left + Me.LblUnitPrice.Width
                End If
            End If

            Me.HitungSubTotal()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
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

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
        Dim row As DataGridViewRow
        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPurchaseReturn_Quick_AddByCustomProduct") = True Then
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                If e.RowIndex <> -1 Then
                    If Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value Or _
                            Me.TxtRemark.Enabled = False Or Me.dgvItem.ReadOnly = True Then
                        Exit Sub
                    End If

                    Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    Dim Expiry As Date = #1/1/1950#
                    If IsDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value) = True Then
                        Expiry = Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value
                    Else
                        Expiry = #1/1/1950#
                    End If

                    frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)
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

                        If Me.ApakahEditSah(Me.dgvItem.Rows(e.RowIndex), Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value, -1, _
                        Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value, _
                        Expiry, -1, Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value, True, False) = False Then
                            MsgBox("Ratio value restored to previous value, because ratio you entered will cause negative stock quantity.", MsgBoxStyle.Information, "Information")

                            Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                            Me.SetRatioAndQtyDefault(row, True)
                        End If

                        If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) <= 0 Then
                            MsgBox("Ratio value restored to previous value, because ratio you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                            Me.SetRatioAndQtyDefault(row, True)
                        End If

                        Me.HitungItem(row.Index, Me.DefaultPrecision)
                    End If
                End If
            ElseIf e.ColumnIndex = Me.colPriceByMultiUnit_Str.Index And Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value Then
                    Dim FrmRef As New Core_Module.FrmTrans_PriceByMultiUnit
                    FrmRef.UnitMulti = Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value
                    FrmRef.UnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    FrmRef.Buffer_PriceByMultiUnit = Me.dgvItem.Item(Me.colPriceByMultiUnit.Index, e.RowIndex).Value

                    If FrmRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.colPriceByMultiUnit.Index, e.RowIndex).Value = FrmRef.cboPriceByMultiUnit.SelectedIndex
                        If Me.dgvItem.Item(Me.colPriceByMultiUnit.Index, e.RowIndex).Value = "1" Then
                            Me.dgvItem.Item(Me.colPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value.ToString
                        Else
                            Me.dgvItem.Item(Me.colPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString
                        End If
                        'hitung ulang row
                        Me.HitungItem(e.RowIndex, Me.DefaultPrecision)
                        'hitung total
                        Me.HitungTotal_Multi()
                    End If

                    FrmRef = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetRatioAndQtyDefault(ByVal vDgvRow As DataGridViewRow, ByVal vSetQtyDefault As Boolean)
        Try
            With vDgvRow
                If .Cells(Me.ColPrimaryUnitMulti.Index).Value <> .Cells(Me.ColPrimaryUnit.Index).Value Then
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

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.dgvItem.Rows.RemoveAt(Me.EraseRowIndex)
            'hitung total yang baru
            Me.HitungTotal_Multi()
            Me.ObjCD.SetNumber(Me.dgvItem)
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
            Me.ObjCD.Number_Leave(Me.TxtDiscAmount, sender, e, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text), Me.ObjCD)
            'disc amount tidak boleh lebih besar dari unit price
            If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                Me.TxtDiscAmount.Text = Me.buffer
                Exit Sub
            End If
            If CDec(Me.TxtDiscAmount.Text) <> CDec(Me.buffer) Then
                Me.TxtDiscFromPerct.Text = "0"
                Me.HitungSubTotal()
            End If
            If Me.TxtRemarkItem.Enabled = True Then
                If Me.TxtDiscAmount.Focused = True Then
                    If My.Computer.Keyboard.ShiftKeyDown = True Then
                        Me.TxtDiscount.Focus()
                    Else
                        Me.TxtRemarkItem.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.GotFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtDiscAmount.ReadOnly = False And Me.TxtDiscAmount.Enabled = True Then
                TxtDiscAmount.BackColor = BackColorAll
            Else
                TxtDiscAmount.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Private Sub TxtDiscAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.LostFocus
        If Me.TxtRemark.Enabled = True Then
            If Me.TxtDiscAmount.ReadOnly = False And Me.TxtDiscAmount.Enabled = True Then
                TxtDiscAmount.BackColor = Me.TxtPrimaryPurchaseReturn.BackColor
            End If
        End If
    End Sub

    Public Sub FromCutOff_PReturn_Quick_Create(ByVal vPrimaryTransaksi As Double, ByVal vPrimarySupp As Double)
        Try
            Me.baru = False
            Me.objPurchaseReturn.ResetObject(Me.dgvItem)
            Me.objPurchaseReturn.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me, Me.TblCustomDummy, Me.TblItemInOutCustom)
            Me.objPurchaseReturn.PrimarySupplier = vPrimarySupp
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_PReturn_Quick_Save(ByVal NoStokckOutPR As String) As Double
        Try
            Me.baru = True
            Me.objPurchaseReturn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me, Me.TblCustomDummy, False, NoStokckOutPR, False, False, True)
            Return Me.objPurchaseReturn.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
#Region "Taufik"
    Dim TblCustomDummy As New CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableDataTable
    Dim ObjCustom2 As New CustomSMESS.ClsTransaksiSalesInvoiceQuick_AddByCustomProduct
    Dim ObjUser As New Core_Module.ClsUser
    Public TblItemInOutCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
    Sub Atur1()
        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPurchaseReturn_Quick") = True Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPurchaseReturn_Quick_AddByCustomProduct") = False Then
                    'Me.gboItem.Enabled = True
                    Me.SplitContainer1.Enabled = True
                    Me.AddByCustomProductToolStripMenuItem.Visible = False
                    Me.dgvItem.ReadOnly = False
                    Me.dgvItem.AllowUserToDeleteRows = True
                Else
                    Me.dgvItem.ReadOnly = True
                    'Me.gboItem.Enabled = False
                    Me.SplitContainer1.Enabled = False
                    Me.AddByCustomProductToolStripMenuItem.Visible = True
                    Me.dgvItem.AllowUserToDeleteRows = False
                End If
            Else
                'Me.gboItem.Enabled = True
                Me.SplitContainer1.Enabled = True
                Me.AddByCustomProductToolStripMenuItem.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
            End If
            Me.AddItemByCustom.Visible = False
            Me.LblAdvance.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAdvance.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(Me.LblAdvance, New Point(0, 15))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddByCustomProductToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddByCustomProductToolStripMenuItem.Click
        Dim ObjCustIdentifier As New CustomSMESS.ClsIdentifier
        Try
            If ObjCustIdentifier.GetIdentifier = "smess_admiral2" Then
                Me.LoadCustomForAdmiral(sender, e)
            Else
                Me.LoadCustomForBinaya(sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub LoadCustomForBinaya(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Frm As New CustomSMESS.FrmTransaksiPurchaseReturn_Quick_AddByCustomProduct
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            Frm.PrimaryPurchaseReturn = Me.TxtPrimaryPurchaseReturn.Text
            Frm.PrimarySupplier = Me.txtPrimarySupplier.Text
            Frm.PrimaryCurrency = Me.TxtPrimaryCurrency.Text
            Frm.PrimaryCurrencyTax = Me.txtPrimaryCurrTax.Text
            Frm.PrimaryPO = Me.txtPrimaryPO.Text
            Frm.PONo = Me.cboPO.Text
            Frm.baru = Me.baru
            If Me.txtPrimaryGRN.Text = "" Then
                Me.txtPrimaryGRN.Text = "-1"
            End If
            Frm.PrimaryGRN = Me.txtPrimaryGRN.Text

            If Me.cboPO.Visible = True Then
                Frm.cboPO = True
            Else
                Frm.cboPO = False
            End If
            If Me.chkTax.Checked = True Then
                Frm.IsTax = True
            End If
            If Me.btnOK.Enabled = True Then
                Frm.AktifForm = True
            Else
                Frm.AktifForm = False
            End If

            Frm.TblDummy = Me.TblCustomDummy

            If Frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblCustomDummy = Frm.TblDummy
                Me.IsikanDgvFromTblDummyBinaya()
                Me.HitungTotal_Multi()
                Me.FormatTxtStockIN()
                Me.FormatDGV()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub
    Public Function IsikanDgvFromTblDummyBinaya() As Boolean
        Dim rowDummy As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim RoundingItem As Integer = 0
        Dim PrimaryPO As Double = -1
        Dim PO_Number As String = ""
        Dim ObjPO As New ClsTransaksiPO
        Dim objProduct As New Core_Module.clsMasterProduct

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'kosongkan dgv
            Me.dgvItem.Rows.Clear()

            PrimaryPO = -1
            PO_Number = ""

            'isikan tabeljj    
            Me.dgvItem.AllowUserToAddRows = False
            For Each rowDummy In Me.TblCustomDummy.Rows

                Dim Index As Integer = 0

                If Me.cboPO.Visible = True Then
                    Index = Me.GetIndex(rowDummy.Primary_Product, rowDummy.Batch_No, _
                    rowDummy.Secondary_Qty, rowDummy.Expiry_Date, rowDummy.Primary_PO, rowDummy.primary_location)
                Else
                    Index = Me.GetIndex(rowDummy.Primary_Product, rowDummy.Batch_No, _
                    rowDummy.Secondary_Qty, rowDummy.Expiry_Date, rowDummy.Primary_PO, rowDummy.primary_location, _
                    rowDummy.Unit_Price, rowDummy.Discount, rowDummy.Primary_Unit_Multi)
                End If

                'cari data yang sama  
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)

                    ''@rosadinnn
                    ''@MUaT
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + rowDummy.Total_Qty_Multi
                    Me.SetQtyDefaultOnDgv(row.Index)

                    RoundingItem = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)
                    If row.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                                CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), RoundingItem))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                                CDec(row.Cells(Me.ColTotalQty.Index).Value), RoundingItem))
                    End If

                    'Me.HitungTotal_Multi()
                Else
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        If TblCustomDummy.Rows.IndexOf(rowDummy) = 0 Then
                            rowDummy = TblCustomDummy.Rows(0)
                            RoundingItem = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)
                        End If

                        Dim expiryDate As String
                        If rowDummy.Expiry_Date <> #1/1/1950# Then
                            expiryDate = Format(rowDummy.Expiry_Date, "dd MMM yyyy")
                        Else
                            expiryDate = "-"
                        End If


                        rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDummy.Product_Code
                        rowDGV.Cells(Me.ColDescription.Index).Value = rowDummy.Description
                        rowDGV.Cells(Me.ColPONO.Index).Value = rowDummy.PO_No
                        rowDGV.Cells(Me.ColBatchNo.Index).Value = rowDummy.Batch_No
                        rowDGV.Cells(Me.ColExpiryDate.Index).Value = expiryDate
                        rowDGV.Cells(Me.ColQtyMaster.Index).Value = 0
                        rowDGV.Cells(Me.ColQtyPerMaster.Index).Value = 0
                        rowDGV.Cells(Me.ColMasterUnits.Index).Value = ""
                        rowDGV.Cells(Me.ColQty.Index).Value = rowDummy.Qty
                        rowDGV.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(rowDummy.Total_Qty, rowDummy._Decimal)
                        rowDGV.Cells(Me.ColUnit.Index).Value = rowDummy.Unit_Code
                        rowDGV.Cells(Me.ColSecQty.Index).Value = -1
                        rowDGV.Cells(Me.ColSecUnit.Index).Value = ""
                        rowDGV.Cells(Me.colRemarks.Index).Value = rowDummy.Remarks
                        rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = rowDummy.Primary_Product
                        rowDGV.Cells(Me.ColPrimaryPO.Index).Value = rowDummy.Primary_PO
                        rowDGV.Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryUnit.Index).Value = rowDummy.Primary_Units
                        rowDGV.Cells(Me.ColPrimarySecUnit.Index).Value = rowDummy.Primary_Secondary_Units

                        rowDGV.Cells(Me.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Unit_Price, RoundingItem))
                        rowDGV.Cells(Me.ColDiscount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Discount, 2))
                        rowDGV.Cells(Me.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Disc_Amount, RoundingItem))
                        rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = rowDummy.Disc_From_Perct
                        rowDGV.Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Nett_Price, RoundingItem))
                        rowDGV.Cells(Me.ColAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Amount, RoundingItem))
                        rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = rowDummy.Primary_Currency
                        rowDGV.Cells(Me.ColCurrency.Index).Value = Me.ObjCurrency.GetCurrencyCode(rowDummy.Primary_Currency)
                        rowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.RemoveDecimalZero(CDec(rowDummy.Exchanged_Rate)))
                        rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = rowDummy.primary_location
                        rowDGV.Cells(Me.ColLocation.Index).Value = rowDummy.Storage_Area_Code
                        rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        rowDGV.Cells(Me.ColPurchasInIDR.Index).Value = rowDummy.Purchase_Price_In_IDR

                        ''@rosadinnn
                        ''@MUaT
                        rowDGV.Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Reference, objProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Unit_Multi))
                        rowDGV.Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Default, objProduct.GetAngkaDecimal(rowDummy.Primary_Product))
                        rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value = rowDummy.Primary_Unit_Multi
                        If rowDummy.Primary_Units <> rowDummy.Primary_Unit_Multi Then
                            rowDGV.Cells(Me.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Reference, _
                                    objProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Unit_Multi)) _
                                    & " " & rowDummy.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(rowDummy.Ratio_Default, _
                                    objProduct.GetAngkaDecimal(rowDummy.Primary_Product)) & " " & rowDummy.Unit_Code
                        Else
                            rowDGV.Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(rowDummy.Total_Qty, objProduct.GetAngkaDecimal(rowDummy.Primary_Product))
                        rowDGV.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(rowDummy.Total_Qty_Multi, objProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Unit_Multi))
                        rowDGV.Cells(Me.ColUnitsCodeMulti.Index).Value = rowDummy.Unit_Code_Multi

                    End With

                    'isikan data harga+hitung total
                    'Me.HitungTotal_Multi()
                End If

                'format dgv
                'Me.FormatDGV()
                'Me.FormatTxtStockIN()
                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While


            Next
            'kosongkan gbo item
            Me.KosongkanItem()
            Me.TxtCodeNo.Focus()
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Items.Clear()
            Me.cboBatchNo.Visible = False
            Me.cboBatchNo.TabStop = False
            Me.DtpExpired.Visible = False
            Me.DtpExpired.TabStop = False
            Me.cboLocation.Visible = False
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Items.Clear()
            Me.cboLocation.TabStop = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtCurrency.ReadOnly = True
                Me.TxtCurrency.BackColor = Me.txtReadOnly.BackColor
                Me.TxtCurrency.TabStop = False
                Me.btnFindCurrency.Enabled = False
            End If
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL = Nothing
            ObjCD = Nothing
            ObjPO = Nothing
        End Try
    End Function

    Public Function IsikanDgvFromTblDummyAdmiral() As Boolean
        Dim rowDummy As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim RoundingItem As Integer = 0
        Dim PrimaryPO As Double = -1
        Dim PO_Number As String = ""
        Dim ObjPO As New ClsTransaksiPO
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim primaryRoll As Double = 22

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'kosongkan dgv
            Me.dgvItem.Rows.Clear()

            PrimaryPO = -1
            PO_Number = ""

            'isikan tabeljj    
            Me.dgvItem.AllowUserToAddRows = False
            For Each rowDummy In Me.TblCustomDummy.Rows

                Dim Index As Integer = 0

                If Me.cboPO.Visible = True Then
                    Index = Me.GetIndex(rowDummy.Primary_Product, rowDummy.Batch_No, _
                    rowDummy.Secondary_Qty, rowDummy.Expiry_Date, rowDummy.Primary_PO, rowDummy.primary_location)
                Else
                    Index = Me.GetIndex(rowDummy.Primary_Product, rowDummy.Batch_No, _
                    rowDummy.Secondary_Qty, rowDummy.Expiry_Date, rowDummy.Primary_PO, rowDummy.primary_location, _
                    rowDummy.Unit_Price, rowDummy.Discount, rowDummy.Primary_Unit_Multi)
                End If

                'cari data yang sama  
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)

                    ''@rosadinnn
                    ''@MUaT

                    row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + rowDummy.Total_Qty_Multi, objProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitMulti.Index).Value))

                    If rowDummy.Is_Kain = True Then
                        row.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQty.Index).Value) + rowDummy.Total_Qty, objProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                        row.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQtyMulti.Index).Value, Me.ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product, primaryRoll)) + _
                        " " + row.Cells(Me.ColUnitsCodeMulti.Index).Value + " = " + Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value, Me.ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product)) + " " + row.Cells(Me.ColUnit.Index).Value
                        row.Cells(Me.ColRatioReference.Index).Value = rowDummy.Total_Qty_Multi
                        row.Cells(Me.ColRatioDefault.Index).Value = rowDummy.Total_Qty
                    Else
                        Me.SetQtyDefaultOnDgv(row.Index)
                    End If


                    RoundingItem = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)
                    If row.Cells(Me.colPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                                CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), RoundingItem))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColNettPrice.Index).Value) * _
                                CDec(row.Cells(Me.ColTotalQty.Index).Value), RoundingItem))
                    End If
                    If rowDummy.Is_Kain = True Then
                        row.Cells(Me.ColRatioDesc.Index).Value = row.Cells(ColTotalQtyMulti.Index).Value _
                                                   & " " & rowDummy.Unit_Code_Multi & " = " & row.Cells(ColTotalQty.Index).Value & " " & rowDummy.Unit_Code
                    End If

                    If String.IsNullOrEmpty(row.Cells(Me.colRemarks.Index).Value) = False Then
                        row.Cells(Me.colRemarks.Index).Value = row.Cells(Me.colRemarks.Index).Value & ", " & rowDummy.Remarks
                    Else
                        row.Cells(Me.colRemarks.Index).Value = rowDummy.Remarks
                    End If
                    'Me.HitungTotal_Multi()
                Else
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        If TblCustomDummy.Rows.IndexOf(rowDummy) = 0 Then
                            rowDummy = TblCustomDummy.Rows(0)
                            RoundingItem = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)
                        End If

                        Dim expiryDate As String
                        If rowDummy.Expiry_Date <> #1/1/1950# Then
                            expiryDate = Format(rowDummy.Expiry_Date, "dd MMM yyyy")
                        Else
                            expiryDate = "-"
                        End If


                        rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDummy.Product_Code
                        rowDGV.Cells(Me.ColDescription.Index).Value = rowDummy.Description
                        rowDGV.Cells(Me.ColPONO.Index).Value = rowDummy.PO_No
                        rowDGV.Cells(Me.ColBatchNo.Index).Value = rowDummy.Batch_No
                        rowDGV.Cells(Me.ColExpiryDate.Index).Value = expiryDate
                        rowDGV.Cells(Me.ColQtyMaster.Index).Value = 0
                        rowDGV.Cells(Me.ColQtyPerMaster.Index).Value = 0
                        rowDGV.Cells(Me.ColMasterUnits.Index).Value = ""
                        rowDGV.Cells(Me.ColQty.Index).Value = rowDummy.Qty
                        rowDGV.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(rowDummy.Total_Qty, rowDummy._Decimal)
                        rowDGV.Cells(Me.ColUnit.Index).Value = rowDummy.Unit_Code
                        rowDGV.Cells(Me.ColSecQty.Index).Value = -1
                        rowDGV.Cells(Me.ColSecUnit.Index).Value = ""
                        rowDGV.Cells(Me.colRemarks.Index).Value = rowDummy.Remarks
                        rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = rowDummy.Primary_Product
                        rowDGV.Cells(Me.ColPrimaryPO.Index).Value = rowDummy.Primary_PO
                        rowDGV.Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        rowDGV.Cells(Me.ColPrimaryUnit.Index).Value = rowDummy.Primary_Units
                        rowDGV.Cells(Me.ColPrimarySecUnit.Index).Value = rowDummy.Primary_Secondary_Units

                        rowDGV.Cells(Me.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Unit_Price, RoundingItem))
                        rowDGV.Cells(Me.ColDiscount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Discount, 2))
                        rowDGV.Cells(Me.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Disc_Amount, RoundingItem))
                        rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = rowDummy.Disc_From_Perct
                        rowDGV.Cells(Me.ColNettPrice.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Nett_Price, RoundingItem))
                        rowDGV.Cells(Me.ColAmount.Index).Value = ObjCD.DigitGrouping(ObjCD.SetDecimal(rowDummy.Amount, RoundingItem))
                        rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = rowDummy.Primary_Currency
                        rowDGV.Cells(Me.ColCurrency.Index).Value = Me.ObjCurrency.GetCurrencyCode(rowDummy.Primary_Currency)
                        rowDGV.Cells(Me.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(ObjCD.RemoveDecimalZero(CDec(rowDummy.Exchanged_Rate)))
                        rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = rowDummy.primary_location
                        rowDGV.Cells(Me.ColLocation.Index).Value = rowDummy.Storage_Area_Code
                        rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        rowDGV.Cells(Me.ColPurchasInIDR.Index).Value = rowDummy.Purchase_Price_In_IDR

                        ''@rosadinnn
                        ''@MUaT
                        rowDGV.Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Reference, objProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Unit_Multi))
                        rowDGV.Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Default, objProduct.GetAngkaDecimal(rowDummy.Primary_Product))
                        rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value = rowDummy.Primary_Unit_Multi
                        rowDGV.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(rowDummy.Total_Qty, objProduct.GetAngkaDecimal(rowDummy.Primary_Product))
                        rowDGV.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(rowDummy.Total_Qty_Multi, objProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Unit_Multi))
                        If rowDummy.Is_Kain = False Then
                            If rowDummy.Primary_Units <> rowDummy.Primary_Unit_Multi Then
                                rowDGV.Cells(Me.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(rowDummy.Ratio_Reference, _
                                        objProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Unit_Multi)) _
                                        & " " & rowDummy.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(rowDummy.Ratio_Default, _
                                        objProduct.GetAngkaDecimal(rowDummy.Primary_Product)) & " " & rowDummy.Unit_Code
                            Else
                                rowDGV.Cells(Me.ColRatioDesc.Index).Value = "-"
                            End If
                        Else
                            rowDGV.Cells(Me.ColRatioDesc.Index).Value = rowDGV.Cells(ColTotalQtyMulti.Index).Value _
                                                       & " " & rowDummy.Unit_Code_Multi & " = " & rowDGV.Cells(ColTotalQty.Index).Value & " " & rowDummy.Unit_Code
                        End If


                        rowDGV.Cells(Me.ColUnitsCodeMulti.Index).Value = rowDummy.Unit_Code_Multi
                        If rowDummy.Is_Kain = True Then
                            rowDGV.Cells(Me.ColRatioReference.Index).Value = rowDummy.Total_Qty_Multi
                            rowDGV.Cells(Me.ColRatioDefault.Index).Value = rowDummy.Total_Qty
                        End If
                        rowDGV.Cells(Me.colPriceByMultiUnit.Index).Value = rowDummy.Price_by_Multi_Unit
                        If rowDummy.Price_by_Multi_Unit = "1" Then
                            rowDGV.Cells(Me.colPriceByMultiUnit_Str.Index).Value = "/" + rowDummy.Unit_Code_Multi
                        Else
                            rowDGV.Cells(Me.colPriceByMultiUnit_Str.Index).Value = "/" + rowDummy.Unit_Code
                        End If

                    End With

                    'isikan data harga+hitung total
                    'Me.HitungTotal_Multi()
                End If

                'format dgv
                'Me.FormatDGV()
                'Me.FormatTxtStockIN()
                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While


            Next
            'kosongkan gbo item
            'Me.KosongkanItem()
            Me.TxtCodeNo.Focus()
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Items.Clear()
            Me.cboBatchNo.Visible = False
            Me.cboBatchNo.TabStop = False
            Me.DtpExpired.Visible = False
            Me.DtpExpired.TabStop = False
            Me.cboLocation.Visible = False
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Items.Clear()
            Me.cboLocation.TabStop = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtCurrency.ReadOnly = True
                Me.TxtCurrency.BackColor = Me.txtReadOnly.BackColor
                Me.TxtCurrency.TabStop = False
                Me.btnFindCurrency.Enabled = False
            End If
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL = Nothing
            ObjCD = Nothing
            ObjPO = Nothing
        End Try
    End Function

    Public Sub LoadCustomForAdmiral(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Frm As New CustomSMESS.FrmTransaksiPurchaseReturn_Quick_AddByCustomProduct
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            Frm.UserCode = Me.UserCode
            Frm.BackColorAll = Me.BackColorAll
            Frm.baru = Me.baru
            Frm.DefaultPrecision = Me.DefaultPrecision
            Frm.PrimaryPurchaseReturn = Me.txtPrimaryStockInPurchase.Text
            If Me.txtPrimaryGRN.Text = "" Then
                Me.txtPrimaryGRN.Text = -1
            End If
            Frm.PrimaryGRN = Me.txtPrimaryGRN.Text
            Frm.TanggalTransaksi = Me.dtpDate.Value.Date
            If Me.TxtRemark.Enabled = True Then
                Frm.CanEditForm = True
            Else
                Frm.CanEditForm = False
            End If

            Frm.PrimaryPurchaseReturn = Me.TxtPrimaryPurchaseReturn.Text
            Frm.PrimarySupplier = Me.txtPrimarySupplier.Text
            Frm.PrimaryCurrency = Me.TxtPrimaryCurrency.Text
            Frm.PrimaryStockIn = Me.txtPrimaryStockInPurchase.Text

            'Frm.PrimaryCurrencyTax = Me.txtPrimaryCurrTax.Text
            Frm.PrimaryPO = Me.txtPrimaryPO.Text
            Frm.PONo = Me.cboPO.Text
            Frm.baru = Me.baru
            If Me.txtPrimaryGRN.Text = "" Then
                Me.txtPrimaryGRN.Text = "-1"
            End If
            Frm.PrimaryGRN = Me.txtPrimaryGRN.Text

            If Me.cboPO.Visible = True Then
                Frm.cboPO = True
            Else
                Frm.cboPO = False
            End If
            If Me.chkTax.Checked = True Then
                Frm.IsTax = True
            End If
            If Me.btnOK.Enabled = True Then
                Frm.AktifForm = True
            Else
                Frm.AktifForm = False
            End If

            Frm.TblDummy = Me.TblCustomDummy.Copy

            If Frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                'Me.TblItemInOutCustom.Clear()
                'Me.TblCustomDummy.Clear()
                Me.TblCustomDummy = Frm.TblDummy.Copy
                Me.TblItemInOutCustom = Frm.TblItemCustomInOut.Copy
                Me.IsikanDgvFromTblDummyAdmiral()
                Me.HitungTotal_Multi()
                Me.FormatTxtStockIN()
                Me.FormatDGV()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub
#End Region

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiPurchaseReturn_Quick_AddByCustomProduct
        Try
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
                            Me.HitungSubTotal()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class