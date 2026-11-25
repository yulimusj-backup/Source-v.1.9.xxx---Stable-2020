Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiQuotationOtorisasi2
    Public UserCode As String = ""
    Public vPrimaryUserCode As Double = -1
    Public AuthorizationLevel As Integer = 0
    Public CancelAuthorizationLevel As Integer = 1
    Public StrictSalesAuthorization As Boolean = True
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0

    Dim ObjQuo As New ClsTransaksiQuotation
    Dim ObjCust As New clsMasterCustomer

    Dim TBL As New DseTransaksiQuotation.tbltrans_qoutation_authDataTable
    Dim ObjCurrency As New clsMasterCurrency
    Dim ObjCd As New ClsCekData
    Dim DefaultPrecision As Integer = 0
    Dim ObjUser As New ClsUser
    Dim CurrenctIndex As Integer = -1
    Dim vTBLItemDelivery As New DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable
    Dim vFrmQuo As New FrmTransaksiQuotation
    Dim EraseRowIndex As Integer = 0
    Dim buffer As String = ""
    Dim ObjCustomer As New clsMasterCustomer

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCLose.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasi_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            If Me.TabControl2.SelectedIndex = 0 Then
                Me.LoadOtorisasi()
            ElseIf Me.TabControl2.SelectedIndex = 1 Then
                Me.LoadCancelOtorisasi()
            Else
                Me.LoadAlreadyAuthorize()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasi_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'tanyakan apakah ada yang akan dicancel otorisasi/ otorisasi
        Dim CheckAuth As Boolean = False
        Try
            If Me.TabControl2.SelectedIndex = 0 Then
                'otorisasi
                For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
                    If rowDgv.Cells(Me.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString Then
                        CheckAuth = True
                        Exit For
                    End If
                Next
                If CheckAuth = True Then
                    Dim i As MsgBoxResult = MsgBox("Authorize before closing?", MsgBoxStyle.YesNoCancel, _
                    "Confirm")
                    If i = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjQuo.authorizeQuotation(Me.DgvOtorisasi, Me.vPrimaryUserCode, _
                            Me.AuthorizationLevel) = True Then
                                MsgBox("Authorization complete")
                            End If
                        End If
                    ElseIf i = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            Else
                'cancel otorisasi
                For Each rowDgv As DataGridViewRow In Me.DgvCancelOtorisasi.Rows
                    If rowDgv.Cells(Me.ColDgvUnAuth_Cancel.Index).Value = Boolean.TrueString Then
                        CheckAuth = True
                        Exit For
                    End If
                Next
                If CheckAuth = True Then
                    Dim i As MsgBoxResult = MsgBox("Cancel authorization before closing?", _
                    MsgBoxStyle.YesNoCancel, "Confirm")
                    If i = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjQuo.CancelAuthorizeQuotation(Me.DgvCancelOtorisasi, Me.vPrimaryUserCode, _
                            Me.CancelAuthorizationLevel) = True Then
                                MsgBox("Cancellation complete")
                            End If
                        End If
                    ElseIf i = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            End If
            ObjQuo.SaveUserItem(UserCode, Me.DgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasi_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'load data pertama
            ObjQuo.LoadUserItem(UserCode, Me.DgvUser)
            Me.vPrimaryUserCode = ObjUser.GetPrimary(Me.UserCode)
            Me.LoadOtorisasi()
            Me.DgvAlreadyAuthorize.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCancelOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.MinimizeBox = True
            Me.dgvItem.ReadOnly = True
            Me.CboStatusQuotation.SelectedIndex = 0
            Me.DgvUser.AllowUserToAddRows = True
            Me.DtpFrom.Enabled = False
            Me.DtpTo.Enabled = False
            Me.CboStatusQuotation.Enabled = False

            Me.TabPage3.Text = "Cancel Auht."
            Me.TabPage4.Text = "List"
            Me.ColSupplierCodeNo.Visible = False
            Me.ColSuppDesc.Visible = False

            Me.dtpDate.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.dtpDate.Width
            Me.TxtStockInNo.Width = Me.dtpDate.Left - 5 - Me.TxtStockInNo.Left
            Me.TxtCustomerName.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.TxtCustomerName.Left
            Me.LlbAddress.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.LlbAddress.Width
            Me.cboCustomerAddress.Width = Me.LlbAddress.Left - 5 - Me.cboCustomerAddress.Left
            Me.llbAuthorizationDetail.Left = Me.LlbAddress.Left
            Me.txtContactPerson.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.txtContactPerson.Left
            Me.txtPhoneNo.Width = Me.txtContactPerson.Width
            Me.txtAutohirizationStatus.Width = Me.LlbAddress.Left - 5 - Me.txtAutohirizationStatus.Left
            Me.cboPaymentDueDateReference.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.cboPaymentDueDateReference.Width
            Me.Label20.Left = Me.cboPaymentDueDateReference.Left - 5 - Me.Label20.Width
            Me.txtPaymentDueDate.Width = Me.Label20.Left - 5 - Me.txtPaymentDueDate.Left
            Me.txtNoNPWP.Width = Me.txtContactPerson.Width
            Me.TxtInquiryNo.Width = Me.txtContactPerson.Width

            'Me.txtPaymentTerm.Left = Me.SplitContainer1.Panel2.Width - 5 - Me.txtPaymentTerm.Width
            'Me.Label8.Left = Me.txtPaymentTerm.Left - 5 - Me.Label8.Width
            'Me.txtCurrencyCode.Width = Me.Label8.Left - 5 - Me.txtCurrencyCode.Left
            Me.txtSubTotal.Width = Me.SplitContainer1.Panel2.Width - 5 - Me.txtSubTotal.Left
            Me.TxtTotal.Width = Me.txtSubTotal.Width
            Me.TxtRemark.Width = Me.txtSubTotal.Width
            Me.txtTermsNCondition.Width = Me.txtSubTotal.Width
            Me.txtTaxAmount.Width = Me.SplitContainer1.Panel2.Width - 5 - Me.txtTaxAmount.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            Try
                If Me.TabControl2.SelectedIndex = 1 Then
                    Me.ChkQuoDate.Checked = False
                    Me.CboStatusQuotation.Enabled = False
                    Me.LoadCancelOtorisasi()
                    Me.tsbAuthorize.Text = "Cancel Authorization"
                    Me.tsbAuthorize.Image = My.Resources.lock_open
                    Me.tsbAuthorize.Visible = True
                    Me.DtpFrom.Value = Now
                    Me.DtpTo.Value = Now
                    Me.TsbPrint.Visible = False
                    'Me.tsbRemove.Visible = False
                    'Me.tsbShowRemoved.Visible = False
                ElseIf Me.TabControl2.SelectedIndex = 0 Then
                    Me.ChkQuoDate.Checked = False
                    Me.CboStatusQuotation.Enabled = False
                    Me.LoadOtorisasi()
                    Me.tsbAuthorize.Text = "Authorize"
                    Me.tsbAuthorize.Image = My.Resources.lock_closed
                    Me.tsbAuthorize.Visible = True
                    Me.DtpFrom.Value = Now
                    Me.DtpTo.Value = Now
                    Me.TsbPrint.Visible = False
                    'Me.tsbRemove.Visible = True
                    'Me.tsbShowRemoved.Visible = True
                ElseIf Me.TabControl2.SelectedIndex = 2 Then
                    Me.ChkQuoDate.Checked = True
                    Me.DtpFrom.Value = DateAdd("m", -3, Me.DtpTo.Value)
                    Me.DtpTo.Value = Now
                    Me.CboStatusQuotation.Enabled = True
                    Me.LoadAlreadyAuthorize()
                    Me.tsbAuthorize.Visible = False
                    Me.TsbPrint.Visible = True
                    'Me.tsbRemove.Visible = False
                    'Me.tsbShowRemoved.Visible = False
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadOtorisasi()
        Dim Primary_ As Double = -1
        Try
            If Me.UserCode <> "" And Me.AuthorizationLevel >= 0 And Me.AuthorizationLevel < 3 Then
                Me.ObjQuo.IsiDgvBoxOtorisasi(Me.AuthorizationLevel, _
                Me.DgvOtorisasi, Me.TBL, Me.DgvUser, Me.StrictSalesAuthorization, Me.UserCode, _
                Me.TxtNoQuotation.Text, Me.DtpFrom.Value, Me.DtpTo.Value, Me.ChkQuoDate.Checked)
                If Me.DgvOtorisasi.Rows.Count > 0 Then
                    DgvOtorisasi.Rows(0).Selected = True
                    Primary_ = Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjQuo.CariDetail(Primary_, Me.dgvItem, Me.vTBLItemDelivery)
                    Me.SalinKeField()
                Else
                    Me.KosongkanField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCancelOtorisasi()
        Dim primary_ As Double
        Try
            'Me.BtnFilter.PerformClick()
            'DgvCancelOtorisasi.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.UserCode <> "" And Me.CancelAuthorizationLevel >= 1 And Me.AuthorizationLevel < 4 Then
                Me.ObjQuo.IsiDgvBoxBatalOtorisasi(Me.CancelAuthorizationLevel, _
                Me.DgvCancelOtorisasi, Me.TBL, Me.DgvUser, Me.StrictSalesAuthorization, Me.UserCode, _
                Me.TxtNoQuotation.Text, Me.DtpFrom.Value, Me.DtpTo.Value, Me.ChkQuoDate.Checked)
                If Me.DgvCancelOtorisasi.Rows.Count > 0 Then
                    DgvCancelOtorisasi.Rows(0).Selected = True
                    primary_ = Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjQuo.CariDetail(primary_, Me.dgvItem, Me.vTBLItemDelivery)
                    Me.SalinKeField()
                Else
                    Me.KosongkanField()
                End If
            End If
            'DgvCancelOtorisasi.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadAlreadyAuthorize()
        Dim primary_ As Double
        Try
            'Me.BtnFilter.PerformClick()
            If Me.UserCode <> "" And Me.AuthorizationLevel >= 0 And Me.AuthorizationLevel < 3 Then
                Me.ObjQuo.IsiDgvBoxAlreadyAuthorize(Me.AuthorizationLevel, _
                Me.DgvAlreadyAuthorize, Me.TBL, Me.DgvUser, Me.StrictSalesAuthorization, Me.UserCode, _
                Me.TxtNoQuotation.Text, Me.DtpFrom.Value, Me.DtpTo.Value, Me.CboStatusQuotation.SelectedIndex, Me.ChkQuoDate.Checked)
                If Me.DgvAlreadyAuthorize.Rows.Count > 0 Then
                    DgvAlreadyAuthorize.Rows(0).Selected = True
                    primary_ = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjQuo.CariDetail(primary_, Me.dgvItem, Me.vTBLItemDelivery)
                    Me.SalinKeField()
                Else
                    Me.KosongkanField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsbAuthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbAuthorize.Click
        Dim CheckAuth As Boolean = False
        Try
            If Me.TabControl2.SelectedIndex = 0 Then
                'otorisasi
                Me.DgvOtorisasi.EndEdit()
                For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
                    If rowDgv.Cells(Me.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString Then
                        CheckAuth = True
                        Exit For
                    End If
                Next
                If CheckAuth = True Then
                    If MsgBox("Authorize?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjQuo.authorizeQuotation(Me.DgvOtorisasi, Me.vPrimaryUserCode, _
                            Me.AuthorizationLevel) = True Then
                                MsgBox("Authorization complete")
                                Me.LoadOtorisasi()
                            End If
                        End If
                    End If
                Else
                    MsgBox("Please select item to authorize", _
                    MsgBoxStyle.Information, "Information")
                End If
            Else
                Me.DgvCancelOtorisasi.EndEdit()
                'batal otorisasi
                For Each rowDgv As DataGridViewRow In Me.DgvCancelOtorisasi.Rows
                    If rowDgv.Cells(Me.ColDgvUnAuth_Cancel.Index).Value = Boolean.TrueString Then
                        CheckAuth = True
                        Exit For
                    End If
                Next
                If CheckAuth = True Then
                    If MsgBox("Cancel authorization?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjQuo.CancelAuthorizeQuotation(Me.DgvCancelOtorisasi, Me.vPrimaryUserCode, _
                            Me.CancelAuthorizationLevel) = True Then
                                MsgBox("Cancellation complete")
                                Me.LoadCancelOtorisasi()
                            End If
                        End If
                    End If
                Else
                    MsgBox("Please select item to cancel authorize", _
                    MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim ij As New FrmTransaksiQuotation_SetDelivery
        Dim DI As Integer = 0
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today
        Dim i As New FrmTransaksiQuotation

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                If Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value <> "-" And _
                Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value = Boolean.TrueString Then
                    ij.vPrimaryCodeNo = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value
                    ij.VIdentifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                    ij.FrmRef = i
                    i.vTBLItemDelivery = Me.vTBLItemDelivery
                    ij.vReadOnly = True
                    ij.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtSubTotal.Text = "0"
            Me.txtTax.Text = "0"
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.txtPaymentDueDate.Text = "0"
            Me.txtContactPerson.Text = ""
            Me.txtNoNPWP.Text = ""
            Me.txtPhoneNo.Text = ""
            Me.txtAutohirizationStatus.Text = ""
            Me.txtPaymentTerm.Text = ""
            Me.TxtInquiryNo.Text = ""
            Me.cboCustomerAddress.Text = ""
            Me.cboPaymentDueDateReference.Text = ""
            Me.txtTermsNCondition.Text = ""

            Me.TxtStockInNo.Text = ""
            Me.txtCustomerCode.Text = ""
            Me.TxtCustomerName.Text = ""
            Me.txtPrimaryCustomer.Text = -1
            Me.dtpDate.Value = Now
            Me.txtCurrencyCode.Text = ""
            Me.TxtRemark.Text = ""
            Me.TxtCustomerCodeShipTo.Text = ""
            Me.TxtCustomerDescShipTo.Text = ""
            Me.CboAddressCustomerShipTo.Items.Clear()
            Me.CboAddressShipTo.Items.Clear()

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()

            Me.chkTax.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjQuo
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtPrimaryCustomer.Text = .PrimaryCustomer
                Me.txtNoNPWP.Text = .CustNPWP
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtPhoneNo.Text = .CustPhone
                Me.txtPaymentDueDate.Text = .PaymentDue
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtContactPerson.Text = .CustInquiryNo
                Me.txtPaymentTerm.Text = .PaymentTerm
                Me.cboPaymentDueDateReference.SelectedIndex = .RefPaymentDate
                Me.TxtStockInNo.Text = .NoTransaksi
                ObjCust.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.txtCustomerCode.Text = Me.ObjCust.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCust.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.txtCustomerAddIdentifier.Text = .CustIdentifierAdd
                Me.cboCustomerAddress.Text = Me.ObjCust.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision()
                Me.txtSubTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.TaxAmount, Me.DefaultPrecision))
                Me.txtTax.Text = .TaxPercentage
                Me.TxtInquiryNo.Text = .CustInquiryNo
                Me.txtTermsNCondition.Text = .TermCondition
                Me.TxtTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.Total, Me.DefaultPrecision))
                Me.chkTax.Checked = .UseTax
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                If Me.ObjQuo.AuthorizationLevel = 0 Then
                    Me.txtAutohirizationStatus.Text = "Unauthorized"
                ElseIf Me.ObjQuo.AuthorizationLevel = 1 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Sales Staff"
                ElseIf Me.ObjQuo.AuthorizationLevel = 2 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Sales Manager"
                ElseIf Me.ObjQuo.AuthorizationLevel = 3 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Director"
                End If
                If .NeedDP = 1 Then
                    Me.chkUseDP.Checked = True
                Else
                    Me.chkUseDP.Checked = False
                End If
                Me.TxtDPFromPercent.Text = .DPFromPercent
                Me.TxtDPPercent.Text = Me.ObjCd.DigitGrouping(.DPPercent, 2)
                Me.TxtDPAmount.Text = Me.ObjCd.DigitGrouping(.DPAmount, Me.DefaultPrecision)
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDefaultPrecision()
        Try
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
            Me.FormatDGV()
            Me.Hitung_Total()
            Me.txtSubTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            Me.txtTaxAmount.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.txtTaxAmount.Text, DefaultPrecision))
            Me.TxtTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
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
            Me.txtSubTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(SubTotal, DefaultPrecision))
            'hitung pajak
            If Me.chkTax.Checked = True Then
                Me.txtTaxAmount.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.txtTax.Text * SubTotal / 100, _
                DefaultPrecision))
            Else
                Me.txtTax.Text = "0"
                Me.txtTaxAmount.Text = "0"
            End If
            'hitung total = sub total + tax
            Me.TxtTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(CDbl(Me.txtSubTotal.Text) + _
            CDbl(Me.txtTaxAmount.Text), DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk delivery date expectation yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColDeliveryDate.Index).Value) = False Then
                    row.Cells(Me.ColDeliveryDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColDeliveryDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    Else
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColDeliveryDate.Index).Value = Format(CDate(row.Cells(Me.ColDeliveryDate.Index).Value), "dd MMM yyyy")
                    End If
                End If
                'digit grouping: qty, unit price, nett price, amount
                If IsNumeric(row.Cells(Me.ColQty.Index).Value) = True Then
                    row.Cells(Me.ColQty.Index).Value = Me.ObjCd.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
                End If
                If IsNumeric(row.Cells(Me.ColUnitPrice.Index).Value) = True Then
                    row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value, Me.DefaultPrecision))
                End If
                If IsNumeric(row.Cells(Me.ColNettPrice.Index).Value) = True Then
                    row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(row.Cells(Me.ColNettPrice.Index).Value, Me.DefaultPrecision))
                End If
                If IsNumeric(row.Cells(Me.ColAmount.Index).Value) = True Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(row.Cells(Me.ColAmount.Index).Value, Me.DefaultPrecision))
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAuthorizationDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthorizationDetail.LinkClicked
        If Me.TxtPrimaryStockIn.Text = "" Or Me.TxtPrimaryStockIn.Text = "-1" Then
            Exit Sub
        End If

        Dim jmlbaris As Integer = 0
        Dim i As New FrmTransaksiQuotationAuthorizationDetail
        Dim row As DataGridViewRow
        Try
            i.dgvDetail.Rows.Clear()
            'tambahkan 3 atau 2 baris authorization
            If Me.ObjQuo.AuthorizationLevelNeeded = 0 Then
                'untuk yang baru, anggap 2 level
                jmlbaris = 2
            Else
                jmlbaris = Me.ObjQuo.AuthorizationLevelNeeded
            End If
            For j As Integer = 1 To jmlbaris
                i.dgvDetail.Rows.Add()
                row = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)
                If row.Index = 0 Then
                    row.Cells(Me.ColCodeNo.Index).Value = "Sales Staff Authorization"
                    If Me.ObjQuo.SalesStaffAuthorizationPrimary <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjQuo.SalesStaffAuthorizationPrimary)
                        row.Cells(3).Value = Format(Me.ObjQuo.SalesStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 1 Then
                    row.Cells(Me.ColCodeNo.Index).Value = "Sales Manager Authorization"
                    If Me.ObjQuo.SalesManagerAuthorizationPrimary <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjQuo.SalesManagerAuthorizationPrimary)
                        row.Cells(3).Value = Format(Me.ObjQuo.SalesManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 2 Then
                    row.Cells(Me.ColCodeNo.Index).Value = "General Manager Authorization"
                    If Me.ObjQuo.GeneralManagerAuthorizationPrimary <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjQuo.GeneralManagerAuthorizationPrimary)
                        row.Cells(3).Value = Format(Me.ObjQuo.GeneralManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                End If
            Next
            'tunjukkan form
            Me.ObjCd.SetNumber(i.dgvDetail)
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindPO.Click
    '    Me.Close()
    'End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub DgvOtorisasi_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvOtorisasi.SelectionChanged
        Dim Primary_ As Double
        Try
            Me.KosongkanField()
            If Me.DgvOtorisasi.SelectedRows.Count > 0 Then
                'If Me.CurrenctIndex <> Me.DgvOtorisasi.SelectedRows Then
                'Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                'Dim rowview As DataRowView
                Primary_ = Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells.Item(0).RowIndex).Value
                'i = rowview.Row
                Me.ObjQuo.CariDetail(Primary_, Me.dgvItem, Me.vTBLItemDelivery)
                Me.SalinKeField()
                'End If
            Else
                Me.KosongkanField()
            End If
            'Me.CurrenctIndex = Me.chkBelumOtorisasi.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCancelOtorisasi_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCancelOtorisasi.SelectionChanged
        Dim Primary_ As Double
        Try
            Me.KosongkanField()
            If Me.DgvCancelOtorisasi.SelectedRows.Count > 0 Then
                'If Me.CurrenctIndex <> Me.DgvOtorisasi.SelectedRows Then
                'Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                'Dim rowview As DataRowView
                Primary_ = Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells.Item(0).RowIndex).Value
                'i = rowview.Row
                Me.ObjQuo.CariDetail(Primary_, Me.dgvItem, Me.vTBLItemDelivery)
                Me.SalinKeField()
                'End If
            Else
                Me.KosongkanField()
            End If
            'Me.CurrenctIndex = Me.chkBelumOtorisasi.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvAlreadyAuthorize_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvAlreadyAuthorize.CellContentClick
        'Dim ij As New Inventory_Module.FrmCetakTransaksiPO

        'Try
        '    Me.DgvAlreadyAuthorize.EndEdit()
        '    If e.RowIndex = -1 Then
        '        Exit Sub
        '    End If
        '    If e.ColumnIndex = Me.ColAA_Print.Index Then
        '        If Me.TxtPrimaryStockIn.Text <> "" Then
        '            If IsNumeric(Me.TxtPrimaryStockIn.Text) = True Then
        '                'hanya Quotation yang sudah diotorisasi yang boleh diprint
        '                If Me.ObjQuo.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjQuo.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
        '                    MsgBox("This quotation require more authorization, please check authorization needed.", _
        '                    MsgBoxStyle.Information, "Information")
        '                    Exit Sub
        '                Else
        '                    If ObjQuo.CetakQuotation(ij.crv1, Me.TxtPrimaryStockIn.Text) = True Then
        '                        ij.ShowDialog()
        '                    End If
        '                    Exit Sub
        '                End If
        '            End If
        '        End If

        '        If Me.TxtPrimaryStockIn.Text <> "" And Me.TxtPrimaryStockIn.Text <> "-1" Then
        '            'hanya PO yang sudah diotorisasi yang boleh diprint
        '            If Me.ObjQuo.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjQuo.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
        '                MsgBox("This quotation require more authorization, please check authorization needed.", _
        '                MsgBoxStyle.Information, "Information")
        '                Exit Sub
        '            Else
        '                If ObjQuo.CetakQuotation(ij.crv1, Me.TxtPrimaryStockIn.Text) = True Then
        '                    ij.ShowDialog()
        '                End If
        '            End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    ij.Dispose()
        'End Try
    End Sub

    Private Sub DgvAlreadyAuthorize_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvAlreadyAuthorize.SelectionChanged
        Dim Primary_ As Double
        Try
            Me.KosongkanField()
            If Me.DgvAlreadyAuthorize.SelectedRows.Count > 0 Then
                'If Me.CurrenctIndex <> Me.DgvOtorisasi.SelectedRows Then
                'Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                'Dim rowview As DataRowView
                Primary_ = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                'i = rowview.Row
                Me.ObjQuo.CariDetail(Primary_, Me.dgvItem, Me.vTBLItemDelivery)
                Me.SalinKeField()
                'End If
            Else
                Me.KosongkanField()
            End If
            'Me.CurrenctIndex = Me.chkBelumOtorisasi.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbCheckAll_Authorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_Authorize.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
                rowDgv.Cells(Me.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LlbUncheckAll_Athorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_Athorize.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
                rowDgv.Cells(Me.ColDgvAuth_Authorize.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LlbCheckAll_Cancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_Cancel.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDgv As DataGridViewRow In Me.DgvCancelOtorisasi.Rows
                rowDgv.Cells(Me.ColDgvUnAuth_Cancel.Index).Value = Boolean.TrueString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LlbUncheckAll_Cancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_Cancel.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each rowDgv As DataGridViewRow In Me.DgvCancelOtorisasi.Rows
                rowDgv.Cells(Me.ColDgvUnAuth_Cancel.Index).Value = Boolean.FalseString
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub TxtNoQuotation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNoQuotation.KeyDown
        Dim ObjQuo As New ClsTransaksiQuotation
        Try
            If e.KeyValue = 13 Then
                If Me.TxtNoQuotation.Text <> "" Then
                    If Me.ObjQuo.ApakahSudahAda(Me.TxtNoQuotation.Text) = True Then
                        Me.TxtPrimaryQuo.Text = Me.ObjQuo.GetPrimary(Me.TxtNoQuotation.Text)
                        Me.TxtNoQuotation.Text = Me.ObjQuo.GetNoTransaksiQuo(Me.TxtPrimaryQuo.Text)
                        Me.CboStatusQuotation.Focus()
                    Else
                        Me.BtnFindQuo.PerformClick()
                    End If
                Else
                    Me.BtnFindQuo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjQuo = Nothing
        End Try
    End Sub

    Private Sub TxtNoQuotation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoQuotation.Leave
        Dim hj As New FrmCariTransaksiQuotation
        Dim ObjQuo As New ClsTransaksiQuotation
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtNoQuotation.Text <> "" Then
                If Me.ObjQuo.ApakahSudahAda(Me.TxtNoQuotation.Text) = True Then
                    Me.TxtPrimaryQuo.Text = Me.ObjQuo.GetPrimary(Me.TxtNoQuotation.Text)
                    Me.TxtNoQuotation.Text = Me.ObjQuo.GetNoTransaksiQuo(Me.TxtPrimaryQuo.Text)
                Else
                    If Me.TxtNoQuotation.Text <> "" Then
                        hj.txtCari.Text = Me.TxtNoQuotation.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryTransaksi > 0 Then
                        Me.TxtPrimaryQuo.Text = hj.PrimaryTransaksi
                        Me.TxtNoQuotation.Text = Me.ObjQuo.GetNoTransaksiQuo(Me.TxtPrimaryQuo.Text)
                        Me.CboStatusQuotation.Focus()
                    Else
                        Me.TxtPrimaryQuo.Text = "-1"
                        Me.TxtNoQuotation.Clear()
                        Me.TxtNoQuotation.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryQuo.Text = "-1"
                Me.TxtNoQuotation.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindQuo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindQuo.Click
        Dim hj As New FrmCariTransaksiQuotation
        Dim ObjQuo As New ClsTransaksiQuotation
        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtNoQuotation.Text <> "" Then
                hj.txtCari.Text = Me.TxtNoQuotation.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryTransaksi > 0 Then
                Me.TxtPrimaryQuo.Text = hj.PrimaryTransaksi
                Me.TxtNoQuotation.Text = Me.ObjQuo.GetNoTransaksiQuo(Me.TxtPrimaryQuo.Text)
                Me.CboStatusQuotation.Focus()
            Else
                Me.TxtPrimaryQuo.Text = "-1"
                Me.TxtNoQuotation.Clear()
                Me.TxtNoQuotation.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vKodeUser Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim KodeUser As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            KodeUser = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(KodeUser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.BgwUser.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            If Me.TxtNoQuotation.Enabled = True Then
                rowDGV = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColCreatedBy.Index).Value = ""
                    .Cells(Me.ColCreatedBy_Name.Index).Value = ""
                    .Cells(Me.ColCreatedBy_Primary.Index).Value = "-1"
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUser.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFilter.Click
        Try
            If Me.TabControl2.SelectedIndex = 0 Then
                Me.LoadOtorisasi()
            ElseIf Me.TabControl2.SelectedIndex = 1 Then
                Me.LoadCancelOtorisasi()
            Else
                Me.LoadAlreadyAuthorize()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkQuoDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkQuoDate.CheckedChanged
        Try
            If Me.ChkQuoDate.Checked = True Then
                Me.DtpFrom.Enabled = True
                Me.DtpTo.Enabled = True
            Else
                Me.DtpFrom.Enabled = False
                Me.DtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNoQuotation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoQuotation.GotFocus
        Me.TxtNoQuotation.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtNoQuotation_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoQuotation.LostFocus
        Me.TxtNoQuotation.BackColor = Me.TxtPrimaryQuo.BackColor
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTermsNCondition.TextChanged

    End Sub

    Private Sub DgvOtorisasi_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvOtorisasi.CellContentClick

    End Sub

    Private Sub LlbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbAddress.LinkClicked
        Try
            If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                Exit Sub
            End If
            Dim ij As New Core_Module.FrmDetailAddress
            ij.vAddIdentifier = Me.txtCustomerAddIdentifier.Text
            ij.PrimaryCust = Me.txtPrimaryCustomer.Text
            ij.btnExit.BringToFront()
            ij.AcceptButton = ij.btnExit
            ij.btnOK.TabStop = False
            ij.cboCustAddCode.Enabled = False
            ij.ShowDialog()
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TsbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TsbPrint.Click
        Dim rowDgv As New DataGridViewRow
        Dim ListPrimary As New List(Of Double)
        Dim ListPrimaryX As New List(Of Double)
        'Dim ij As New Inventory_Module.FrmCetakTransaksiPO

        Try
            Me.DgvAlreadyAuthorize.EndEdit()
            If Me.TabControl2.SelectedIndex = 2 Then
                If Me.DgvAlreadyAuthorize.Rows.Count > 0 Then

                    For Each rowDgv In Me.DgvAlreadyAuthorize.Rows
                        Dim PrimaryQuo As Double = -1
                        If rowDgv.Cells(Me.ColAA_Print.Index).Value = Boolean.TrueString Then
                            If IsNumeric(rowDgv.Cells(Me.ColAA_Primary.Index).Value) Then
                                PrimaryQuo = rowDgv.Cells(Me.ColAA_Primary.Index).Value
                                If Me.ObjQuo.GetAuthorizationLevel(PrimaryQuo) >= Me.ObjQuo.GetAuthorizationLevelNeeded(PrimaryQuo) Then
                                    ListPrimary.Add(PrimaryQuo)
                                Else
                                    ListPrimaryX.Add(PrimaryQuo)
                                    'rowDgv.Cells(Me.ColAA_ChekedPrint.Index).Value = Boolean.FalseString
                                End If
                            End If
                        End If
                    Next

                    Dim strmsg As String = ""
                    If ListPrimaryX.Count > 0 Then
                        For intA As Integer = 0 To ListPrimaryX.Count - 1
                            If strmsg = "" Then
                                strmsg &= Me.ObjQuo.GetNoTransaksiQuo(ListPrimaryX.Item(intA))
                            Else
                                strmsg &= ", " & Me.ObjQuo.GetNoTransaksiQuo(ListPrimaryX.Item(intA))
                            End If
                        Next
                        If strmsg <> "" Then
                            MsgBox("There's list of PO No were require more authorization, please check authorization needed : " & strmsg, MsgBoxStyle.Information, "Information")
                        End If
                    End If

                    If ListPrimary.Count > 0 Then
                        'For intA As Integer = 0 To ListPrimary.Count - 1
                        '    Dim PrimaryPO As Double = -1
                        '    PrimaryPO = ListPrimary.Item(intA)
                        'Next
                        ObjQuo.CetakQuotation_Batch(ListPrimary, 1, 0, True)
                        'If ObjQuo.CetakQuotation_Batch(ij.crv1, ListPrimary) = True Then
                        '    ij.User = Me.UserCode
                        '    ij.ListPrimaryQuo = ListPrimary
                        '    ij.ShowDialog()
                        '    'If ij.Printed = True Then
                        '    '    Dim Primary As Double = -1
                        '    '    Primary = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                        '    '    If Me.ObjQuo.CariDetail(Primary, Me.dgvItem, Me.vTBLItemDelivery) = True Then
                        '    '        Me.SalinKeField()
                        '    '    End If
                        '    'End If
                        'End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

   
    Private Sub LlbViewAddressShipTo_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressShipTo.LinkClicked
        Try
            If Me.CboAddressShipTo.Text <> "" Then
             
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.TxtIdentifierAddressShipTo.Text
                ij.PrimaryCust = Me.txtPrimaryCustomer.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                If Me.TxtRemark.Enabled = False Then
                    ij.cboCustAddCode.Enabled = False
                End If
                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    'Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    'Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbViewAddressCustomerShipTo_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressCustomerShipTo.LinkClicked
        Try
            If Me.CboAddressCustomerShipTo.Text <> "" Then
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.TxtIdentifierAddressCustomerShipTo.Text
                ij.PrimaryCust = Me.TxtPrimaryCustomerShipTo.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                If Me.TxtRemark.Enabled = False Then
                    ij.cboCustAddCode.Enabled = False
                End If
                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    'Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    'Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class