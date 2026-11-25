Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiIPOOtorisasi
    Public UserCode As String = "USER"
    Public vPrimaryUserCode As Double = -1
    Public AuthorizationLevel As Integer = 0
    Public CancelAuthorizationLevel As Integer = 1
    Public StrictSalesAuthorization As Boolean = True
    Public BackColorAll As Color
    Public FormNo As Integer = 0

    Dim ObjIPO As New ClsTransaksiIPO
    Dim ObjBrand As New ClsMasterProductGroup
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSupplier As New clsMasterSupplier

    Dim TBL As New dseTransaksiIPO.tbltrans_ipo2DataTable
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjCd As New Core_Module.ClsCekData
    Dim DefaultPrecision As Integer = 0
    Dim ObjUser As New Core_Module.ClsUser
    Dim CurrenctIndex As Integer = -1

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Me.Close()
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
                            If Me.ObjIPO.Authorize(Me.chkBelumOtorisasi, Me.vPrimaryUserCode, _
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
                            If Me.ObjIPO.CancelAuthorize(Me.chkSudahOtorisasi, Me.vPrimaryUserCode, _
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
            Me.WindowState = FormWindowState.Maximized
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
                ElseIf Me.TabControl1.SelectedIndex = 0 Then
                    Me.LoadOtorisasi()
                    Me.tsbAuthorize.Text = "Authorize"
                    Me.tsbAuthorize.Image = My.Resources.lock_closed1
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
                Me.ObjIPO.IsiChecklistBoxOtorisasi(Me.AuthorizationLevel, _
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
                Me.ObjIPO.IsiChecklistBoxBatalOtorisasi(Me.CancelAuthorizationLevel, _
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
                            If Me.ObjIPO.Authorize(Me.chkBelumOtorisasi, Me.vPrimaryUserCode, _
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
                            If Me.ObjIPO.CancelAuthorize(Me.chkSudahOtorisasi, Me.vPrimaryUserCode, _
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

    Private Sub chkBelumOtorisasi_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBelumOtorisasi.SelectedIndexChanged
        Try
            If Me.chkBelumOtorisasi.SelectedIndex <> -1 Then
                If Me.CurrenctIndex <> Me.chkBelumOtorisasi.SelectedIndex Then
                    Dim i As dseTransaksiIPO.tbltrans_ipo2Row
                    Dim rowview As DataRowView
                    rowview = Me.chkBelumOtorisasi.Items(Me.chkBelumOtorisasi.SelectedIndex)
                    i = rowview.Row
                    Me.ObjIPO.CariDetail(i.Primary, Me.dgvItem)
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
                    Dim i As dseTransaksiIPO.tbltrans_ipo2Row
                    Dim rowview As DataRowView
                    rowview = Me.chkSudahOtorisasi.Items(Me.chkSudahOtorisasi.SelectedIndex)
                    i = rowview.Row
                    Me.ObjIPO.CariDetail(i.Primary, Me.dgvItem)
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
            Me.TxtPrimaryIPO.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtSubTotal.Text = "0"
            Me.txtTax.Text = "0"
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.txtDeliveryTime.Text = "0"
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
            With Me.ObjIPO
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtDeliveryTime.Text = .Deliverytime
                Me.dtpETD.Value = Me.dtpDate.Value + TimeSpan.FromDays(.Deliverytime)
                Me.txtContactPerson.Text = .SuppContact
                Me.txtPrimaryBrand.Text = .PrimaryBrand
                Me.TxtBrand.Text = Me.ObjBrand.GetGroupCode(Me.txtPrimaryBrand.Text)
                Me.txtPrimarySO.Text = .PrimarySO
                Me.ObjSO.GetCustCode_CustName_CustPONO_SalesOrderNumber(Me.txtPrimarySO.Text, _
                Me.TxtCustName.Text, Me.txtCustCode.Text, Me.txtCustPoNo.Text, Me.txtSalesOrderNo.Text, _
                Me.txtCustomerContactName.Text)
                Me.txtPhoneNo.Text = .SuppPhoneNumber
                Me.TxtIPONo.Text = .NoTransaksi
                Me.txtPrimarySupplier.Text = .PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.TxtPrimaryIPO.Text = .Primary
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision()
                Me.txtSubTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.TaxAmount, Me.DefaultPrecision))
                Me.txtTax.Text = .Tax
                Me.TxtTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(.Total, Me.DefaultPrecision))
                If .UseTax = ClsTransaksiIPO.EnumUseTax.NoTax Then
                    Me.chkTax.Checked = False
                Else
                    Me.chkTax.Checked = True
                End If
                Me.TxtRemark.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                If Me.ObjIPO.AuthorizationLevel = 0 Then
                    Me.txtAuthorizationStatus.Text = "Unauthorized"
                ElseIf Me.ObjIPO.AuthorizationLevel = 1 Then
                    Me.txtAuthorizationStatus.Text = "Authorized By Sales Staff"
                ElseIf Me.ObjIPO.AuthorizationLevel = 2 Then
                    Me.txtAuthorizationStatus.Text = "Authorized By Sales Manager"
                ElseIf Me.ObjIPO.AuthorizationLevel = 3 Then
                    Me.txtAuthorizationStatus.Text = "Authorized By Purchase Staff"
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
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxAmount.Text, DefaultPrecision))
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
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
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
            'hitung pajak
            If Me.chkTax.Checked = True Then
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTax.Text * SubTotal / 100, _
                DefaultPrecision))
            Else
                Me.txtTax.Text = "0"
                Me.txtTaxAmount.Text = "0"
            End If
            'hitung total = sub total + tax
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) + _
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
                'digit grouping: qty, unit price, nett price, amount
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
                row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
                row.Cells(6).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(6).Value, Me.DefaultPrecision))
                row.Cells(7).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(7).Value, Me.DefaultPrecision))
                row.Cells(8).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(8).Value, Me.DefaultPrecision))
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAuthorizationDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthorizationDetail.LinkClicked
        If Me.TxtPrimaryIPO.Text = "" Or Me.TxtPrimaryIPO.Text = "-1" Then
            Exit Sub
        End If

        Dim i As New FrmTransaksiPO_AuthorizationDetail
        Dim jmlbaris As Integer = 0
        Dim row As DataGridViewRow
        Try
            i.dgvDetail.Rows.Clear()
            i.Text = "Internal Purchase Order Authorization"

            'tambahkan 3 baris
            jmlbaris = 3
            'masukkan baris dan detail data otorisasi
            For j As Integer = 1 To jmlbaris
                i.dgvDetail.Rows.Add()
                row = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)
                If row.Index = 0 Then
                    row.Cells(1).Value = "Sales Staff Authorization"
                    If Me.ObjIPO.PrimaryStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjIPO.PrimaryStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjIPO.SalesStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 1 Then
                    row.Cells(1).Value = "Sales Manager Authorization"
                    If Me.ObjIPO.PrimarySalesManagerAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjIPO.PrimarySalesManagerAuthorization)
                        row.Cells(3).Value = Format(Me.ObjIPO.SalesManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 2 Then
                    row.Cells(1).Value = "Purchase Staff Authorization"
                    If Me.ObjIPO.PrimaryPurchaseStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjIPO.PrimaryPurchaseStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjIPO.PurchaseStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
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
End Class