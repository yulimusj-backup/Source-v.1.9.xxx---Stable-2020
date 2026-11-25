Option Explicit On
Option Strict Off

Public Class FrmTransaksiPOCOtorisasi
    Public UserCode As String = ""
    Public vPrimaryUserCode As Double = -1
    Public AuthorizationLevel As Integer = 0
    Public CancelAuthorizationLevel As Integer = 1
    Public StrictSalesAuthorization As Boolean = True
    Public BackColorAll As Color
    Public FormNo As Integer = 0

    Dim ObjPOC As New ClsTransaksiPOC
    Dim TBL As New DseTransaksiPOC.tbltrans_poc3DataTable
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjCd As New Core_Module.ClsCekData
    Dim DefaultPrecision As Integer = 0
    Dim ObjUser As New Core_Module.ClsUser
    Dim CurrenctIndex As Integer = -1
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjSupplier As New Core_Module.clsMasterSupplier

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TsbExit.Click
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
                            If Me.ObjPOC.Authorize(Me.chkBelumOtorisasi, Me.vPrimaryUserCode, _
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
                            If Me.ObjPOC.CancelAuthorize(Me.chkSudahOtorisasi, Me.vPrimaryUserCode, _
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            Try
                If Me.TabControl1.SelectedIndex = 1 Then
                    Me.LoadCancelOtorisasi()
                    Me.TsbAuthorize.Text = "Cancel Authorization"
                    Me.TsbAuthorize.Image = My.Resources.lock_open1
                ElseIf Me.TabControl1.SelectedIndex = 0 Then
                    Me.LoadOtorisasi()
                    Me.TsbAuthorize.Text = "Authorize"
                    Me.TsbAuthorize.Image = My.Resources.lock_closed1
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
                Me.ObjPOC.IsiChecklistBoxOtorisasi(Me.AuthorizationLevel, _
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
                Me.ObjPOC.IsiChecklistBoxBatalOtorisasi(Me.CancelAuthorizationLevel, _
                Me.chkSudahOtorisasi, Me.TBL, Me.StrictSalesAuthorization, Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsbAuthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TsbAuthorize.Click
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                'otorisasi
                If Me.chkBelumOtorisasi.CheckedItems.Count > 0 Then
                    If MsgBox("Authorize?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjPOC.Authorize(Me.chkBelumOtorisasi, Me.vPrimaryUserCode, _
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
                            If Me.ObjPOC.CancelAuthorize(Me.chkSudahOtorisasi, Me.vPrimaryUserCode, _
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
                    Dim i As DseTransaksiPOC.tbltrans_poc3Row
                    Dim rowview As DataRowView
                    rowview = Me.chkBelumOtorisasi.Items(Me.chkBelumOtorisasi.SelectedIndex)
                    i = rowview.Row
                    Me.ObjPOC.CariDetail(i.Primary, Me.dgvItem)
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
                    Dim i As DseTransaksiPOC.tbltrans_poc3Row
                    Dim rowview As DataRowView
                    rowview = Me.chkSudahOtorisasi.Items(Me.chkSudahOtorisasi.SelectedIndex)
                    i = rowview.Row
                    Me.ObjPOC.CariDetail(i.Primary, Me.dgvItem)
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
            'Me.KosongkanItem()
            Core_Module.ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryPOC.Text = "-1"
            Me.txtPrimaryPO.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim SupName, CPName, PhoneNo As String
        SupName = ""
        CPName = ""
        PhoneNo = ""

        Try
            With Me.ObjPOC
                Me.TxtPrimaryPOC.Text = .primary
                Me.txtPrimaryPO.Text = .primaryPO
                Me.TxtPOCNo.Text = .NoTransaksi
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
                Me.txtCancellationReason.Text = .CancellationReason
                Me.TxtRemark.Text = .Remarks

                Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = SupName
                Me.txtContactPerson.Text = CPName
                Me.txtPhoneNo.Text = PhoneNo
                'Me.txtPrimaryCodeNo.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                Me.txtPONO.Text = ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()

            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                row.Cells(5).Value = Me.ObjCd.DigitGrouping(row.Cells(5).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDefaultPrecision()
        Try
            'Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
            Me.FormatDGV()
            'Me.Hitung_Total()
            'Me.txtSubTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            'Me.txtTaxAmount.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.txtTaxAmount.Text, DefaultPrecision))
            'Me.TxtTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                'digit grouping
                row.Cells(Me.ColTotalQty.Index).Value = Me.ObjCd.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAuthorizationDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthorization.LinkClicked
        Dim i As New FrmTransaksiPOCAuthorizationDetail
        Dim jmlbaris As Integer = 0
        Dim row As DataGridViewRow
        Try
            i.dgvDetail.Rows.Clear()
            'tambahkan 3 atau 2 baris authorization
            If Me.ObjPOC.AuthorizationNeeded = 0 Then
                'untuk yang baru, anggap 2 level
                jmlbaris = 2
            Else
                jmlbaris = Me.ObjPOC.AuthorizationNeeded
            End If
            'masukkan baris dan detail data otorisasi
            For j As Integer = 1 To jmlbaris
                i.dgvDetail.Rows.Add()
                row = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)
                If row.Index = 0 Then
                    row.Cells(1).Value = "Purchase Staff Authorization"
                    If Me.ObjPOC.PrimaryPurchaseStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPOC.PrimaryPurchaseStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPOC.PurchaseStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 1 Then
                    row.Cells(1).Value = "Purchase Manager Authorization"
                    If Me.ObjPOC.PrimaryManagerStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPOC.PrimaryManagerStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPOC.PurchaseManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(2).Value = "-"
                        row.Cells(3).Value = "-"
                    End If
                ElseIf row.Index = 2 Then
                    row.Cells(1).Value = "Purchase Director Authorization"
                    If Me.ObjPOC.PrimaryDirectorStaffAuthorization <> -1 Then
                        row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPOC.PrimaryDirectorStaffAuthorization)
                        row.Cells(3).Value = Format(Me.ObjPOC.PurchaseDirectorAuthorizationDate, "dd MMM yyyy - HH:mm")
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class