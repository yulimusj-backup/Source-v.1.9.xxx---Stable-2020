Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiPOOtorisasi
    Public UserCode As String = ""
    Public vPrimaryUserCode As Double = -1
    Public AuthorizationLevel As Integer = 0
    Public CancelAuthorizationLevel As Integer = 1
    Public StrictSalesAuthorization As Boolean = True
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0

    Dim ObjPO As New ClsTransaksiPO
    Dim ObjSupplier As New clsMasterSupplier

    Dim TBL As New DseTransaksiPO.tbltranspo_authorize_view2DataTable
    Dim ObjCurrency As New clsMasterCurrency
    Dim ObjCd As New ClsCekData
    Dim DefaultPrecision As Integer = 0
    Dim ObjUser As New ClsUser
    Dim CurrenctIndex As Integer = -1
    Dim vTBLItemDelivery As New DseTransaksiPO.tbltrans_po_itemdeliveryDataTable
    Dim vFrmPO As New FrmTransaksiPO

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPOOtorisasi_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.LoadOtorisasi()
            Else
                Me.LoadCancelOtorisasi()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasi_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'tanyakan apakah ada yang akan dicancel otorisasi/ otorisasi
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                'otorisasi
                If Me.chkBelumOtorisasi.CheckedItems.Count > 0 Then
                    Dim i As MsgBoxResult = MsgBox("Authorize before closing?", MsgBoxStyle.YesNoCancel, _
                    "Confirm")
                    If i = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjPO.Authorize(Me.chkBelumOtorisasi, Me.vPrimaryUserCode, _
                            Me.AuthorizationLevel) = True Then
                                MsgBox("Authorization complete")
                            End If
                        End If
                    ElseIf i = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            Else
                If Me.chkSudahOtorisasi.CheckedItems.Count > 0 Then
                    Dim i As MsgBoxResult = MsgBox("Cancel authorization before closing?", _
                    MsgBoxStyle.YesNoCancel, "Confirm")
                    If i = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjPO.CancelAuthorize(Me.chkSudahOtorisasi, Me.vPrimaryUserCode, _
                            Me.CancelAuthorizationLevel) = True Then
                                MsgBox("Cancellation complete")
                            End If
                        End If
                    ElseIf i = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasi_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'load data pertama
            Me.vPrimaryUserCode = ObjUser.GetPrimary(Me.UserCode)
            Me.LoadOtorisasi()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.MinimizeBox = True
            Me.dgvItem.ReadOnly = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            Try
                If Me.TabControl1.SelectedIndex = 1 Then
                    Me.LoadCancelOtorisasi()
                    Me.tsbAuthorize.Text = "Cancel Authorization"
                    Me.tsbAuthorize.Image = My.Resources.lock_open1
                    Me.tsbRemove.Visible = False
                    Me.tsbShowRemoved.Visible = False
                ElseIf Me.TabControl1.SelectedIndex = 0 Then
                    Me.LoadOtorisasi()
                    Me.tsbAuthorize.Text = "Authorize"
                    Me.tsbAuthorize.Image = My.Resources.lock_closed1
                    Me.tsbRemove.Visible = True
                    Me.tsbShowRemoved.Visible = True
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadOtorisasi()
        Try
            If Me.UserCode <> "" And Me.AuthorizationLevel >= 0 And Me.AuthorizationLevel < 3 Then
                Me.ObjPO.IsiChecklistBoxOtorisasi(Me.AuthorizationLevel, _
                Me.chkBelumOtorisasi, Me.TBL, Me.StrictSalesAuthorization, Me.UserCode)
                If Me.chkBelumOtorisasi.Items.Count > 0 Then
                    Me.chkBelumOtorisasi.SetSelected(0, True)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCancelOtorisasi()
        Try
            If Me.UserCode <> "" And Me.CancelAuthorizationLevel >= 1 And Me.AuthorizationLevel < 4 Then
                Me.ObjPO.IsiChecklistBoxBatalOtorisasi(Me.CancelAuthorizationLevel, _
                Me.chkSudahOtorisasi, Me.TBL, Me.StrictSalesAuthorization, Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsbAuthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbAuthorize.Click
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                'otorisasi
                If Me.chkBelumOtorisasi.CheckedItems.Count > 0 Then
                    If MsgBox("Authorize?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjPO.Authorize(Me.chkBelumOtorisasi, Me.vPrimaryUserCode, _
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
                'batal otorisasi
                If Me.chkSudahOtorisasi.CheckedItems.Count > 0 Then
                    If MsgBox("Cancel authorization?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjPO.CancelAuthorize(Me.chkSudahOtorisasi, Me.vPrimaryUserCode, _
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


    Private Sub chkBelumOtorisasi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBelumOtorisasi.SelectedIndexChanged
        Try
            If Me.chkBelumOtorisasi.SelectedIndex <> -1 Then
                If Me.CurrenctIndex <> Me.chkBelumOtorisasi.SelectedIndex Then
                    Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                    Dim rowview As DataRowView
                    rowview = Me.chkBelumOtorisasi.Items(Me.chkBelumOtorisasi.SelectedIndex)
                    i = rowview.Row
                    Me.ObjPO.CariDetail(i.Primary, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
                    Me.SalinKeField()
                End If
            Else
                Me.KosongkanField()
            End If
            Me.CurrenctIndex = Me.chkBelumOtorisasi.SelectedIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSudahOtorisasi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSudahOtorisasi.SelectedIndexChanged
        Try
            If Me.chkSudahOtorisasi.SelectedIndex <> -1 Then
                If Me.CurrenctIndex <> Me.chkSudahOtorisasi.SelectedIndex Then
                    Dim i As DseTransaksiPO.tbltranspo_authorize_view2Row
                    Dim rowview As DataRowView
                    rowview = Me.chkSudahOtorisasi.Items(Me.chkSudahOtorisasi.SelectedIndex)
                    i = rowview.Row

                    Me.ObjPO.CariDetail(i.Primary, Me.dgvItem, Me.vTBLItemDelivery, Me.vFrmPO)
                    Me.SalinKeField()
                End If
            Else
                Me.KosongkanField()
            End If
            Me.CurrenctIndex = Me.chkSudahOtorisasi.SelectedIndex
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
        If Me.TxtPrimaryStockIn.Text = "" Or Me.TxtPrimaryStockIn.Text = "-1" _
        And Me.TxtRemark.Enabled = False Then
            Exit Sub
        End If

        Dim i As New FrmTransaksiSupplierAuthorizationDetail
        Dim jmlbaris As Integer = 0
        Dim row As DataGridViewRow
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
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub llbUncheckAllAuthorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAllAuthorize.LinkClicked
        Try
            For i As Integer = 0 To Me.chkBelumOtorisasi.Items.Count - 1
                Me.chkBelumOtorisasi.SetItemChecked(i, False)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAllAuthorize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAllAuthorize.LinkClicked
        Try
            For i As Integer = 0 To Me.chkBelumOtorisasi.Items.Count - 1
                Me.chkBelumOtorisasi.SetItemChecked(i, True)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAllCancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAllCancel.LinkClicked
        Try
            For i As Integer = 0 To Me.chkSudahOtorisasi.Items.Count - 1
                Me.chkSudahOtorisasi.SetItemChecked(i, True)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAllCancel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAllCancel.LinkClicked
        Try
            For i As Integer = 0 To Me.chkSudahOtorisasi.Items.Count - 1
                Me.chkSudahOtorisasi.SetItemChecked(i, False)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCheckAllAuthorize_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAllAuthorize.LinkClicked
        Try
            For i As Integer = 0 To Me.chkBelumOtorisasi.Items.Count - 1
                Me.chkBelumOtorisasi.SetItemChecked(i, True)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAllAuthorize_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAllAuthorize.LinkClicked
        Try
            For i As Integer = 0 To Me.chkBelumOtorisasi.Items.Count - 1
                Me.chkBelumOtorisasi.SetItemChecked(i, False)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

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
        Dim row As DseTransaksiPO.tbltranspo_authorize_view2Row
        Dim rowview As DataRowView
        Dim DTAPO As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            If Me.chkBelumOtorisasi.CheckedItems.Count > 0 Then
                If MsgBox("Remove PO from authorization form?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    For Each rowview In Me.chkBelumOtorisasi.CheckedItems
                        row = rowview.Row
                        DTAPO.UpdateQuery_SetRemoved(1, row.Primary)
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
End Class