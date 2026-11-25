Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiPOOtorisasi2
    Public UserCode As String = ""
    Public vPrimaryUserCode As Double = -1
    Public AuthorizationLevel As Integer = 0
    Public CancelAuthorizationLevel As Integer = 1
    Public StrictSalesAuthorization As Boolean = True
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0

    Dim ObjPO As New ClsTransaksiPO
    Dim ObjSupplier As New clsMasterSupplier

    Dim TBL As New DseTransaksiPO.tbltranspo_authorize_view2_2DataTable
    Dim ObjCurrency As New clsMasterCurrency
    Dim ObjCd As New ClsCekData
    Dim DefaultPrecision As Integer = 0
    Dim ObjUser As New ClsUser
    Dim CurrenctIndex As Integer = -1
    Dim vTBLItemDelivery As New DseTransaksiPO.tbltrans_po_itemdeliveryDataTable
    Dim vFrmPO As New FrmTransaksiPO
    Dim EraseRowIndex As Integer = 0
    Dim buffer As String = ""

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCLose.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPOOtorisasi_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            If Me.TabControl2.SelectedIndex = 0 Then
                Me.LoadOtorisasi()
            ElseIf Me.TabControl2.SelectedIndex = 1 Then
                Me.LoadCancelOtorisasi()
            Else
                Me.LoadAlreadyAuthorize()
            End If

            Me.TxtStockInNo.Width = 140
            Me.dtpDate.Left = Me.TxtStockInNo.Left + Me.TxtStockInNo.Width + 5
            Me.dtpDate.Width = Me.TxtSupplierName.Left + Me.TxtSupplierName.Width - Me.dtpDate.Left
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
                            If Me.ObjPO.Authorize2(Me.DgvOtorisasi, Me.vPrimaryUserCode, _
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
                            If Me.ObjPO.CancelAuthorize2(Me.DgvCancelOtorisasi, Me.vPrimaryUserCode, _
                            Me.CancelAuthorizationLevel) = True Then
                                MsgBox("Cancellation complete")
                            End If
                        End If
                    ElseIf i = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            End If
            ObjPO.SaveUserItem(UserCode, Me.DgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasi_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'load data pertama
            ObjPO.LoadUserItem(UserCode, Me.DgvUser)
            Me.vPrimaryUserCode = ObjUser.GetPrimary(Me.UserCode)
            Me.LoadOtorisasi()
            Me.DgvAlreadyAuthorize.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCancelOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.MinimizeBox = True
            Me.dgvItem.ReadOnly = True
            Me.CboStatusPO.SelectedIndex = 0
            Me.cboStatusPrint.SelectedIndex = 0
            Me.DgvUser.AllowUserToAddRows = True
            Me.DtpFrom.Enabled = False
            Me.DtpTo.Enabled = False
            Me.CboStatusPO.Enabled = False

            'Me.TxtNoPO.Width = Me.BtnFindPO.Left - 5
            'Me.BtnFindPO.BringToFront()
            Me.TabPage3.Text = "Cancel Auth."
            Me.TabPage4.Text = "List"
            Me.TxtNoPO.Anchor = Me.DtpFrom.Anchor
            Me.CboStatusPO.Anchor = Me.DtpFrom.Anchor
            Me.cboStatusPrint.Anchor = Me.DtpFrom.Anchor

            Me.dtpDate.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.dtpDate.Width
            Me.TxtStockInNo.Width = Me.dtpDate.Left - 5 - Me.TxtStockInNo.Left
            Me.TxtSupplierName.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.TxtSupplierName.Left
            Me.LinkLabel2.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.LinkLabel2.Width
            Me.cboSupplierAddress.Width = Me.LinkLabel2.Left - 5 - Me.cboSupplierAddress.Left
            Me.LinkLabel1.Left = Me.LinkLabel2.Left
            Me.llbAuthorizationDetail.Left = Me.LinkLabel2.Left
            Me.txtContactPerson.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.txtContactPerson.Left
            Me.txtPhoneNo.Width = Me.txtContactPerson.Width
            Me.txtAutohirizationStatus.Width = Me.LinkLabel2.Left - 5 - Me.txtAutohirizationStatus.Left
            Me.TxtShipVia.Width = Me.txtAutohirizationStatus.Width
            Me.cboPaymentDueDateReference.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.cboPaymentDueDateReference.Left

            Me.txtPaymentTerm.Width = Me.SplitContainer1.Panel2.Width - 5 - Me.txtPaymentTerm.Left
            Me.txtSubTotal.Width = Me.SplitContainer1.Panel2.Width - 5 - Me.txtSubTotal.Left
            Me.TxtTotal.Width = Me.txtSubTotal.Width
            Me.TxtRemark.Width = Me.txtSubTotal.Width
            Me.txtPrintedBy.Width = Me.txtSubTotal.Width
            Me.txtTaxAmount.Width = Me.SplitContainer1.Panel2.Width - 5 - Me.txtTaxAmount.Left

            Me.ShowIcon = False
            Me.TxtDPAmount.Left = Me.txtTaxAmount.Left
            Me.TxtDPAmount.Width = Me.txtTaxAmount.Width
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            Try
                If Me.TabControl2.SelectedIndex = 1 Then
                    Me.CboStatusPO.Enabled = False
                    Me.LoadCancelOtorisasi()
                    Me.tsbAuthorize.Text = "Cancel Authorization"
                    Me.tsbAuthorize.Image = My.Resources.lock_open1
                    Me.tsbAuthorize.Visible = True
                    Me.tsbRemove.Visible = False
                    Me.tsbShowRemoved.Visible = False
                ElseIf Me.TabControl2.SelectedIndex = 0 Then
                    Me.CboStatusPO.Enabled = False
                    Me.LoadOtorisasi()
                    Me.tsbAuthorize.Text = "Authorize"
                    Me.tsbAuthorize.Image = My.Resources.lock_closed1
                    Me.tsbAuthorize.Visible = True
                    Me.tsbRemove.Visible = True
                    Me.tsbShowRemoved.Visible = True
                ElseIf Me.TabControl2.SelectedIndex = 2 Then
                    Me.DtpFrom.Value = DateAdd("m", -3, Me.DtpTo.Value)
                    Me.ChkPODate.Checked = True
                    Me.CboStatusPO.Enabled = True
                    Me.LoadAlreadyAuthorize()
                    'Me.tsbAuthorize.Text = "Already Authorize"
                    'Me.tsbAuthorize.Image = My.Resources.lock_closed1
                    Me.tsbAuthorize.Visible = False
                    Me.tsbRemove.Visible = False
                    Me.tsbShowRemoved.Visible = False
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
                Me.ObjPO.IsiDgvBoxOtorisasi(Me.AuthorizationLevel, _
                Me.DgvOtorisasi, Me.TBL, Me.DgvUser, Me.StrictSalesAuthorization, Me.UserCode, _
                Me.TxtNoPO.Text, Me.DtpFrom.Value, Me.DtpTo.Value, Me.ChkPODate.Checked)
                If Me.DgvOtorisasi.Rows.Count > 0 Then
                    DgvOtorisasi.Rows(0).Selected = True
                    primary_ = Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjPO.CariDetail(primary_, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
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
                Me.ObjPO.IsiDgvBoxBatalOtorisasi(Me.CancelAuthorizationLevel, _
                Me.DgvCancelOtorisasi, Me.TBL, Me.DgvUser, Me.StrictSalesAuthorization, Me.UserCode, _
                Me.TxtNoPO.Text, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Me.ChkPODate.Checked)
                If Me.DgvCancelOtorisasi.Rows.Count > 0 Then
                    DgvCancelOtorisasi.Rows(0).Selected = True
                    primary_ = Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjPO.CariDetail(primary_, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
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
                Me.ObjPO.IsiDgvBoxAlreadyAuthorize(Me.AuthorizationLevel, _
                Me.DgvAlreadyAuthorize, Me.TBL, Me.DgvUser, Me.StrictSalesAuthorization, Me.UserCode, _
                Me.TxtNoPO.Text, Me.DtpFrom.Value, Me.DtpTo.Value, Me.CboStatusPO.SelectedIndex, Me.ChkPODate.Checked, Me.cboStatusPrint.SelectedIndex)
                If Me.DgvAlreadyAuthorize.Rows.Count > 0 Then
                    DgvAlreadyAuthorize.Rows(0).Selected = True
                    primary_ = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjPO.CariDetail(primary_, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
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
                            If Me.ObjPO.Authorize2(Me.DgvOtorisasi, Me.vPrimaryUserCode, _
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
                            If Me.ObjPO.CancelAuthorize2(Me.DgvCancelOtorisasi, Me.vPrimaryUserCode, _
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

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim ij As New FrmTransaksiPO_SetDelivery
        Dim DI As Integer = 0
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today
        Dim i As New FrmTransaksiPO

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
            Me.txtAutohirizationStatus.Text = ""
            Me.TxtShipVia.Text = ""
            Me.txtPaymentTerm.Text = ""

            Me.TxtStockInNo.Text = ""
            Me.txtSupplierCode.Text = ""
            Me.TxtSupplierName.Text = ""
            Me.txtPrimarySupplier.Text = -1
            Me.txtSupplierAddIdentifier.Text = -1
            Me.cboSupplierAddress.SelectedIndex = -1
            Me.dtpDate.Value = Now
            Me.txtPhoneNo.Text = ""
            Me.txtCurrencyCode.Text = ""
            Me.TxtRemark.Text = ""

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
            With Me.ObjPO
                Me.TxtPrimaryStockIn.Text = .primary
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtPrimarySupplier.Text = .primarySupplier
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtPaymentDueDate.Text = .PaymentDue
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtContactPerson.Text = .SuppContactPerson
                Me.txtPaymentTerm.Text = .PaymentTerm
                Me.cboPaymentDueDateReference.SelectedIndex = .ReferencePaymentDue
                Me.txtPhoneNo.Text = .SuppPhoneNo
                Me.TxtStockInNo.Text = .NoTransaksi
                ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.txtSupplierAddIdentifier.Text = .IdentifierSuppAddress
                Me.cboSupplierAddress.Text = Me.ObjSupplier.GetKodeAddress(Me.txtSupplierAddIdentifier.Text)
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision()
                Me.txtSubTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.TaxAmount, Me.DefaultPrecision))
                Me.txtTax.Text = .Tax
                Me.TxtShipVia.Text = .ShipVia
                Me.TxtTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.Total, Me.DefaultPrecision))
                Me.chkTax.Checked = .TaxExclude
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                If Me.ObjPO.Authorizationlevel = 0 Then
                    Me.txtAutohirizationStatus.Text = "Unauthorized"
                ElseIf Me.ObjPO.Authorizationlevel = 1 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Purchase Staff"
                ElseIf Me.ObjPO.Authorizationlevel = 2 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Purchase Manager"
                ElseIf Me.ObjPO.Authorizationlevel = 3 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Director"
                End If

                If .PrintedBy <> "-" Then
                    Me.txtPrintedBy.Text = .PrintedBy & " - " + Format(.PrintedDate, "dd-MMM-yyyy") + "-" + .PrintedDate.ToShortTimeString
                Else
                    Me.txtPrintedBy.Text = "Unprinted"
                End If

                If .NeedDP = 1 Then
                    Me.chkUseDP.Checked = True
                Else
                    Me.chkUseDP.Checked = False
                End If
                Me.TxtDPFromPercent.Text = .DPFromPercent
                Me.TxtDPPercent.Text = Me.ObjCd.DigitGrouping(.DPPercent, 2)
                Me.TxtDPAmount.Text = Me.ObjCd.DigitGrouping(.DPAmount, Me.DefaultPrecision)

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
                If row.Cells(Me.colUserSecondaryUnits.Index).Value = Boolean.FalseString Then
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
        Dim AdaPriceByDefault As Boolean = False
        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk delivery date expectation yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColDeliveryDate.Index).Value) = False Then
                    If row.Cells(Me.ColDeliveryDate.Index).Value <> "Multiple" Then
                        row.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    End If
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
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    AdaPriceByDefault = True
                End If
            Next

            Me.ColPriceByMultiUnit_Str.Visible = AdaPriceByDefault
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAuthorizationDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthorizationDetail.LinkClicked
        If Me.TxtPrimaryStockIn.Text = "" Or Me.TxtPrimaryStockIn.Text = "-1" Then
            Exit Sub
        End If

        Dim i As New FrmTransaksiPO_AuthorizationDetail
        Dim jmlbaris As Integer = 0
        Dim row As DataGridViewRow
        Dim ObjStockIn As New Inventory_Module.clsTransaksiStockInOut
        Dim ObjPackingList As New ClsTransaksiPackingList
        Dim ObjPOC As New ClsTransaksiPOC
        Dim ObjAOPDP As New ClsTransaksiAOP

        Try
            i.dgvDetail.Rows.Clear()
            'tambahkan 3 baris authorization
            jmlbaris = Me.ObjPO.AuthorizationNeeded
            'masukkan baris dan detail data otorisasi
            For j As Integer = 1 To jmlbaris
                i.dgvDetail.Rows.Add()
                row = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)
                If row.Index = 0 Then
                    row.Cells(1).Value = "Purchase Staff Authorization"
                    If Me.ObjPO.PrimaryPurchaseStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPO.PrimaryPurchaseStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPO.PurchaseStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 1 Then
                    row.Cells(1).Value = "Purchase Manager Authorization"
                    If Me.ObjPO.PrimaryManagerStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPO.PrimaryManagerStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPO.PurchaseManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 2 Then
                    row.Cells(1).Value = "Director Authorization"
                    If Me.ObjPO.PrimaryDirectorStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPO.PrimaryDirectorStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPO.PurchaseDirectorAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                End If
            Next
            'tunjukkan form
            Me.ObjCd.SetNumber(i.dgvDetail)

            'set chk
            '1. stock in purchase
            If ObjStockIn.GetJmlByPrimaryPO(Me.TxtPrimaryStockIn.Text) > 0 Then
                i.chkStockIn.Checked = True
            Else
                i.chkStockIn.Checked = False
            End If
            'packing list
            If ObjPackingList.GetJmlItem_ByPrimaryPO(Me.TxtPrimaryStockIn.Text) > 0 Then
                i.chkPackingList.Checked = True
            Else
                i.chkPackingList.Checked = False
            End If
            'poc
            If ObjPOC.GetJmlByPO(Me.TxtPrimaryStockIn.Text) > 0 Then
                i.chkPOC.Checked = True
            Else
                i.chkPOC.Checked = False
            End If
            'purchase dp
            If ObjAOPDP.GetJml_AOPDP_ByPrimaryPO(Me.TxtPrimaryStockIn.Text) > 0 Then
                i.chkDP.Checked = True
            Else
                i.chkDP.Checked = False
            End If

            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
            ObjStockIn = Nothing
            ObjPackingList = Nothing
            ObjPOC = Nothing
            ObjAOPDP = Nothing
        End Try
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindPO.Click
    '    Me.Close()
    'End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        If Me.TxtPrimaryStockIn.Text = "" Or Me.TxtPrimaryStockIn.Text = "-1" Then
            Exit Sub
        End If

        Dim ij As New FrmTransaksiPO_DetailSIgn

        Try
            ij.TxtPreparedBy.ReadOnly = True
            ij.TxtPreparedBy_Jabatan.ReadOnly = True
            ij.TxtApprovedBy.ReadOnly = True
            ij.TxtApprovedBy_Jabatan.ReadOnly = True
            ij.TxtAcknowledgeBy.ReadOnly = True
            ij.TxtAcknowledgeBy_Jabatan.ReadOnly = True
            ij.btnOK.Visible = False

            ij.TxtPreparedBy.Text = Me.ObjPO.PreparedBy
            ij.TxtPreparedBy_Jabatan.Text = Me.ObjPO.PreparedBy_Jabatan
            ij.TxtApprovedBy.Text = Me.ObjPO.ApprovedBy
            ij.TxtApprovedBy_Jabatan.Text = Me.ObjPO.ApprovedBy_Jabatan
            ij.TxtAcknowledgeBy.Text = Me.ObjPO.AcknowledgeBy
            ij.TxtAcknowledgeBy_Jabatan.Text = Me.ObjPO.AcknowledgeBy_Jabatan

            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbRemove.Click
        'Dim row As DseTransaksiPO.tbltranspo_authorize_view2_2Row
        'Dim rowview As DataRowView
        Dim vFrm As New FrmTransaksiPOOtorisasi2
        Dim DTAPO As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim CheckAuth As Boolean = False
        Dim Primary_ As Double = -1
        Try
            For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
                If rowDgv.Cells(Me.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString Then
                    CheckAuth = True
                    Exit For
                End If
            Next
            If CheckAuth = True Then
                If MsgBox("Remove PO from authorization form?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
                        If rowDgv.Cells(vFrm.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString Then
                            Primary_ = rowDgv.Cells(vFrm.ColDgvAuth_Primary.Index).Value
                            DTAPO.UpdateQuery_SetRemoved(1, Primary_)
                        End If
                    Next

                    Me.LoadOtorisasi()
                End If
            Else
                MsgBox("Please select item to remove from authorization form.", _
                MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPO.Dispose()
        End Try
    End Sub

    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbShowRemoved.Click
        Dim i As New FrmTransaksiPOOtorisasi_Removed
        Try
            i.ShowDialog()
            Me.LoadOtorisasi()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvOtorisasi_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvOtorisasi.SelectionChanged
        Dim Primary_ As Double
        Try
            If Me.DgvOtorisasi.SelectedRows.Count > 0 Then
                'If Me.CurrenctIndex <> Me.DgvOtorisasi.SelectedRows Then
                'Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                'Dim rowview As DataRowView
                Primary_ = Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells.Item(0).RowIndex).Value
                'i = rowview.Row
                Me.ObjPO.CariDetail(Primary_, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
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
            If Me.DgvCancelOtorisasi.SelectedRows.Count > 0 Then
                'If Me.CurrenctIndex <> Me.DgvOtorisasi.SelectedRows Then
                'Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                'Dim rowview As DataRowView
                Primary_ = Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells.Item(0).RowIndex).Value
                'i = rowview.Row
                Me.ObjPO.CariDetail(Primary_, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
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
        Dim ij As New Inventory_Module.FrmCetakTransaksiPO

        Try
            Me.DgvAlreadyAuthorize.EndEdit()
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColAA_Print.Index Then
                If Me.TxtPrimaryStockIn.Text <> "" Then
                    If IsNumeric(Me.TxtPrimaryStockIn.Text) = True Then
                        'hanya PO yang sudah diotorisasi yang boleh diprint
                        If Me.ObjPO.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
                            MsgBox("This purchase order require more authorization, please check authorization needed.", _
                            MsgBoxStyle.Information, "Information")
                            Exit Sub
                        Else
                            If ObjPO.CetakPO(ij.crv1, Me.TxtPrimaryStockIn.Text) = True Then
                                ij.ShowDialog()
                            End If
                            Exit Sub
                        End If
                    End If
                End If

                If Me.TxtPrimaryStockIn.Text <> "" And Me.TxtPrimaryStockIn.Text <> "-1" Then
                    'hanya PO yang sudah diotorisasi yang boleh diprint
                    If Me.ObjPO.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
                        MsgBox("This purchase order require more authorization, please check authorization needed.", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        If ObjPO.CetakPO(ij.crv1, Me.TxtPrimaryStockIn.Text) = True Then
                            ij.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub DgvAlreadyAuthorize_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvAlreadyAuthorize.SelectionChanged
        Dim Primary_ As Double
        Try
            If Me.DgvAlreadyAuthorize.SelectedRows.Count > 0 Then
                'If Me.CurrenctIndex <> Me.DgvOtorisasi.SelectedRows Then
                'Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                'Dim rowview As DataRowView
                Primary_ = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                'i = rowview.Row
                Me.ObjPO.CariDetail(Primary_, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
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

    Private Sub TxtNoPO_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNoPO.KeyDown
        Dim ObjPO As New ClsTransaksiPO
        Try
            If e.KeyValue = 13 Then
                If Me.TxtNoPO.Text <> "" Then
                    If Me.ObjPO.ApakahSudahAda(Me.TxtNoPO.Text) = True Then
                        Me.TxtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.TxtNoPO.Text)
                        Me.TxtNoPO.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                        Me.CboStatusPO.Focus()
                    Else
                        Me.BtnFindPO.PerformClick()
                    End If
                Else
                    Me.BtnFindPO.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjPO = Nothing
        End Try
    End Sub

    Private Sub TxtNoPO_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoPO.Leave
        Dim hj As New FrmCariTransaksiPO
        Dim objPO As New ClsTransaksiPO
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtNoPO.Text <> "" Then
                If Me.ObjPO.ApakahSudahAda(Me.TxtNoPO.Text) = True Then
                    Me.TxtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.TxtNoPO.Text)
                    Me.TxtNoPO.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                Else
                    If Me.TxtNoPO.Text <> "" Then
                        hj.txtCari.Text = Me.TxtNoPO.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryPO > 0 Then
                        Me.TxtPrimaryPO.Text = hj.PrimaryPO
                        Me.TxtNoPO.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                        Me.CboStatusPO.Focus()
                    Else
                        Me.TxtPrimaryPO.Text = "-1"
                        Me.TxtNoPO.Clear()
                        Me.TxtNoPO.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtNoPO.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindPO.Click
        Dim hj As New FrmCariTransaksiPO
        Dim objPO As New ClsTransaksiPO
        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtNoPO.Text <> "" Then
                hj.txtCari.Text = Me.TxtNoPO.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryPO > 0 Then
                Me.TxtPrimaryPO.Text = hj.PrimaryPO
                Me.TxtNoPO.Text = Me.ObjPO.GetNoTransaksiPO(Me.TxtPrimaryPO.Text)
                Me.CboStatusPO.Focus()
            Else
                Me.TxtPrimaryPO.Text = "-1"
                Me.TxtNoPO.Clear()
                Me.TxtNoPO.Focus()
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
            If Me.TxtNoPO.Enabled = True Then
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

    Private Sub ChkPODate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkPODate.CheckedChanged
        Try
            If Me.ChkPODate.Checked = True Then
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

    Private Sub TxtNoPO_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoPO.GotFocus
        Me.TxtNoPO.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtNoPO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoPO.LostFocus
        Me.TxtNoPO.BackColor = Me.TxtPrimaryPO.BackColor
    End Sub

    Private Sub DtpFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpFrom.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.BtnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpTo.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.BtnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim ij As New Core_Module.FrmDetailAddressSupplier

        Try
            If Me.txtPrimarySupplier.Text = "" Or Me.txtPrimarySupplier.Text = "-1" Then
                If Me.TxtRemark.Enabled = True Then
                    MsgBox("Please specify supplier first.", MsgBoxStyle.Information, "Information")
                    Me.txtSupplierCode.Focus()
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If

            ij.vAddIdentifier = Me.txtSupplierAddIdentifier.Text
            ij.PrimarySup = Me.txtPrimarySupplier.Text
            If Me.TxtRemark.Enabled = True Then
                ij.btnOK.BringToFront()
            Else
                ij.cboCustAddCode.Enabled = False
            End If
            ij.AcceptButton = ij.btnOK
            ij.btnExit.TabStop = False
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimarySupplier.Text)
                Me.cboSupplierAddress.Text = Me.ObjSupplier.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbCheckAll_Print_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll_Print.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DgvAlreadyAuthorize.AllowUserToResizeColumns = False
            Me.DgvAlreadyAuthorize.AllowUserToResizeRows = False

            For Each rowDgv As DataGridViewRow In Me.DgvAlreadyAuthorize.Rows
                rowDgv.Cells(Me.ColAA_ChekedPrint.Index).Value = Boolean.TrueString
            Next
            Me.DgvAlreadyAuthorize.AllowUserToResizeColumns = True
            Me.DgvAlreadyAuthorize.AllowUserToResizeRows = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LlbUncheckAll_Print_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll_Print.LinkClicked
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DgvAlreadyAuthorize.AllowUserToResizeColumns = False
            Me.DgvAlreadyAuthorize.AllowUserToResizeRows = False
            For Each rowDgv As DataGridViewRow In Me.DgvAlreadyAuthorize.Rows
                rowDgv.Cells(Me.ColAA_ChekedPrint.Index).Value = Boolean.FalseString
            Next
            Me.DgvAlreadyAuthorize.AllowUserToResizeColumns = True
            Me.DgvAlreadyAuthorize.AllowUserToResizeRows = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub tsbPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbPrint.Click
        Dim rowDgv As New DataGridViewRow
        Dim ListPrimary As New List(Of Double)
        Dim ListPrimaryX As New List(Of Double)
        Dim ij As New Inventory_Module.FrmCetakTransaksiPO


        Try
            If Me.TabControl2.SelectedIndex = 2 Then
                If Me.DgvAlreadyAuthorize.Rows.Count > 0 Then

                    For Each rowDgv In Me.DgvAlreadyAuthorize.Rows
                        Dim PrimaryPO As Double = -1
                        If rowDgv.Cells(Me.ColAA_ChekedPrint.Index).Value = Boolean.TrueString Then
                            If IsNumeric(rowDgv.Cells(Me.ColAA_Primary.Index).Value) Then
                                PrimaryPO = rowDgv.Cells(Me.ColAA_Primary.Index).Value
                                If Me.ObjPO.GetAuthorizationLevel(PrimaryPO) <> Me.ObjPO.GetAuthorizationLevelNeeded(PrimaryPO) Then
                                    ListPrimaryX.Add(PrimaryPO)
                                    'rowDgv.Cells(Me.ColAA_ChekedPrint.Index).Value = Boolean.FalseString
                                Else
                                    ListPrimary.Add(PrimaryPO)
                                End If
                            End If
                        End If
                    Next

                    Dim strmsg As String = ""
                    If ListPrimaryX.Count > 0 Then
                        For intA As Integer = 0 To ListPrimaryX.Count - 1
                            If strmsg = "" Then
                                strmsg &= Me.ObjPO.GetNoTransaksiPO(ListPrimaryX.Item(intA))
                            Else
                                strmsg &= ", " & Me.ObjPO.GetNoTransaksiPO(ListPrimaryX.Item(intA))
                            End If
                        Next
                        If strmsg <> "" Then
                            MsgBox("There's list of PO No were require more authorization, please check authorization needed : " & strmsg, MsgBoxStyle.Information, "Information")
                        End If
                    End If

                    If ListPrimary.Count > 0 Then
                        For intA As Integer = 0 To ListPrimary.Count - 1
                            Dim PrimaryPO As Double = -1
                            PrimaryPO = ListPrimary.Item(intA)
                        Next
                        If ObjPO.CetakPOBanyak(ij.crv1, ListPrimary) = True Then
                            ij.User = Me.UserCode
                            ij.ListPrimaryPo = ListPrimary
                            ij.ShowDialog()
                            If ij.Printed = True Then
                                Dim Primary As Double = -1
                                Primary = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                                If Me.ObjPO.CariDetail(Primary, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO) = True Then
                                    Me.SalinKeField()
                                End If
                            End If
                        End If
                    End If


                    'If e.ColumnIndex = Me.ColAA_Print.Index Then
                    '    If Me.TxtPrimaryStockIn.Text <> "" Then
                    '        If IsNumeric(Me.TxtPrimaryStockIn.Text) = True Then
                    '            'hanya PO yang sudah diotorisasi yang boleh diprint
                    '            If Me.ObjPO.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
                    '                MsgBox("This purchase order require more authorization, please check authorization needed.", _
                    '                MsgBoxStyle.Information, "Information")
                    '                Exit Sub
                    '            Else
                    '                If ObjPO.CetakPO(ij.crv1, Me.TxtPrimaryStockIn.Text) = True Then
                    '                    ij.ShowDialog()
                    '                End If
                    '                Exit Sub
                    '            End If
                    '        End If
                    '    End If

                    '    If Me.TxtPrimaryStockIn.Text <> "" And Me.TxtPrimaryStockIn.Text <> "-1" Then
                    '        'hanya PO yang sudah diotorisasi yang boleh diprint
                    '        If Me.ObjPO.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
                    '            MsgBox("This purchase order require more authorization, please check authorization needed.", _
                    '            MsgBoxStyle.Information, "Information")
                    '            Exit Sub
                    '        Else
                    '            If ObjPO.CetakPO(ij.crv1, Me.TxtPrimaryStockIn.Text) = True Then
                    '                ij.ShowDialog()
                    '            End If
                    '        End If
                    '    End If
                    'End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class