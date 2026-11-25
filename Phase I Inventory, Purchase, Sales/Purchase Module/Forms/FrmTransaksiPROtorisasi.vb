Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiPROtorisasi
    Public UserCode As String = ""
    Public vPrimaryUserCode As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjPR As New ClsTransaksiPR

    Dim TBL As New DseTransaksiPR.TblTransPR_AuthorizeDataTable
    Dim ObjCd As New ClsCekData
    Dim ObjUser As New ClsUser
    Dim ObjRefPar As New ClsMasterRefParty
    Dim CurrenctIndex As Integer = -1
    Dim vTBLItemDelivery As New DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable
    Dim vFrmPR As New FrmTransaksiPR
    Dim EraseRowIndex As Integer = 0
    Dim buffer As String = ""
    Dim vReadyForPurchase As ClsTransaksiPR.EnumReadyForPurchase

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCLose.Click
        Me.Close()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiPOOtorisasi_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            Me.TabPage3.Text = "Cancel Auth."
            Me.TabPage4.Text = "List"

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
                            If Me.ObjPR.Authorize2(Me.DgvOtorisasi, Now, Me.vPrimaryUserCode, ClsTransaksiPR.EnumReadyForPurchase.Yes) = True Then
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
                            If Me.ObjPR.CancelAuthorize2(Me.DgvCancelOtorisasi, Now, -1, ClsTransaksiPR.EnumReadyForPurchase.No) = True Then
                                MsgBox("Cancellation complete")
                            End If
                        End If
                    ElseIf i = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                End If
            End If
            ObjPR.SaveUserItem(UserCode, Me.DgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasi_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'load data pertama
            ObjPR.LoadUserItem(UserCode, Me.DgvUser)
            Me.vPrimaryUserCode = ObjUser.GetPrimary(Me.UserCode)
            Me.LoadOtorisasi()
            Me.DgvAlreadyAuthorize.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCancelOtorisasi.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.MinimizeBox = True
            Me.dgvItem.ReadOnly = True
            Me.CboStatusPR.SelectedIndex = 0
            Me.DgvUser.AllowUserToAddRows = True
            Me.DtpFrom.Enabled = False
            Me.DtpTo.Enabled = False
            Me.CboStatusPR.Enabled = False
            'Me.ToolStrip1.Visible = False
            Me.ToolStrip2.Visible = False

            Me.TxtRemark.Width = Me.SplitContainer1.Panel2.Width - Me.TxtRemark.Left
            Me.txtAutohirizationStatus.Width = Me.SplitContainer1.Panel1.Width - Me.txtAutohirizationStatus.Left
            Me.txtReqByCode.Width = 85
            Me.TxtReqByDesc.Left = Me.txtReqByCode.Left + Me.txtReqByCode.Width + 5
            Me.TxtReqByDesc.Width = Me.SplitContainer1.Panel1.Width - Me.TxtReqByDesc.Left
            Me.dtpDate.Left = Me.SplitContainer1.Panel1.Width - Me.dtpDate.Width
            Me.TxtStockInNo.Left = Me.txtReqByCode.Left
            Me.TxtStockInNo.Width = Me.dtpDate.Left - 5 - Me.TxtStockInNo.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            Try
                If Me.TabControl2.SelectedIndex = 1 Then
                    Me.CboStatusPR.Enabled = False
                    Me.LoadCancelOtorisasi()
                    Me.tsbAuthorize.Text = "Cancel Authorization"
                    Me.tsbAuthorize.Image = My.Resources.lock_open1
                    Me.tsbAuthorize.Visible = True
                    Me.tsbRemove.Visible = False
                    Me.tsbShowRemoved.Visible = False
                ElseIf Me.TabControl2.SelectedIndex = 0 Then
                    Me.CboStatusPR.Enabled = False
                    Me.LoadOtorisasi()
                    Me.tsbAuthorize.Text = "Authorize"
                    Me.tsbAuthorize.Image = My.Resources.lock_closed1
                    Me.tsbAuthorize.Visible = True
                    Me.tsbRemove.Visible = True
                    Me.tsbShowRemoved.Visible = True
                ElseIf Me.TabControl2.SelectedIndex = 2 Then
                    Me.DtpFrom.Value = DateAdd("m", -3, Me.DtpTo.Value)
                    Me.ChkPRDate.Checked = True
                    Me.CboStatusPR.Enabled = True
                    Me.LoadAlreadyAuthorize()
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
            If Me.UserCode <> "" Then
                Me.ObjPR.IsiDgvBoxOtorisasi(Me.DgvOtorisasi, Me.TBL, Me.DgvUser, _
                Me.TxtNoPR.Text, Me.DtpFrom.Value, Me.DtpTo.Value, Me.ChkPRDate.Checked)
                If Me.DgvOtorisasi.Rows.Count > 0 Then
                    DgvOtorisasi.Rows(0).Selected = True
                    Primary_ = Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjPR.CariDetail(Primary_, Me.dgvItem, Me.ObjCd, Me.vFrmPR, Me.vTBLItemDelivery)
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
            If Me.UserCode <> "" Then
                Me.ObjPR.IsiDgvBoxBatalOtorisasi(Me.DgvCancelOtorisasi, Me.TBL, Me.DgvUser, Me.UserCode, _
                Me.TxtNoPR.Text, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, Me.ChkPRDate.Checked)
                If Me.DgvCancelOtorisasi.Rows.Count > 0 Then
                    DgvCancelOtorisasi.Rows(0).Selected = True
                    primary_ = Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjPR.CariDetail(primary_, Me.dgvItem, Me.ObjCd, Me.vFrmPR, Me.vTBLItemDelivery)
                    Me.SalinKeField()
                Else
                    Me.KosongkanField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadAlreadyAuthorize()
        Dim primary_ As Double
        Try
            If Me.UserCode <> "" Then
                Me.ObjPR.IsiDgvBoxAlreadyAuthorize(Me.DgvAlreadyAuthorize, Me.TBL, Me.DgvUser, _
                Me.TxtNoPR.Text, Me.DtpFrom.Value, Me.DtpTo.Value, Me.CboStatusPR.SelectedIndex, Me.ChkPRDate.Checked)
                If Me.DgvAlreadyAuthorize.Rows.Count > 0 Then
                    DgvAlreadyAuthorize.Rows(0).Selected = True
                    primary_ = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                    Me.ObjPR.CariDetail(primary_, Me.dgvItem, Me.ObjCd, Me.vFrmPR, Me.vTBLItemDelivery)
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
                            If Me.ObjPR.Authorize2(Me.DgvOtorisasi, Now, Me.vPrimaryUserCode, ClsTransaksiPR.EnumReadyForPurchase.Yes) = True Then
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
                            If Me.ObjPR.CancelAuthorize2(Me.DgvCancelOtorisasi, Now, -1, ClsTransaksiPR.EnumReadyForPurchase.No) = True Then
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
        Dim ij As New FrmTransaksiPR_SetDelivery
        Dim DI As Integer = 0
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today
        Dim i As New FrmTransaksiPR

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                If Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value <> "-" And _
                Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value = Boolean.TrueString Then
                    ij.vPrimaryCodeNo = Me.dgvItem.Item(Me.ColPrimaryProduct.Index, e.RowIndex).Value
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
            Me.TxtStockInNo.Text = ""
            Me.dtpDate.Value = Now
            Me.txtAutohirizationStatus.Text = ""
            Me.txtReqByCode.Text = ""
            Me.TxtReqByDesc.Text = ""
            Me.txtPrimaryReqBy.Text = -1
            Me.TxtRemark.Text = ""

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
        Try
            With Me.ObjPR
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .Tanggal
                Me.txtPrimaryReqBy.Text = .PrimaryRefParty
                Me.txtReqByCode.Text = Me.ObjRefPar.GetCode(Me.txtPrimaryReqBy.Text)
                Me.TxtReqByDesc.Text = Me.ObjRefPar.GetDesc(Me.txtPrimaryReqBy.Text)
                Me.TxtRemark.Text = .Remarks
                If Me.ObjPR.GetReadyForPurchase(.Primary) = ClsTransaksiPR.EnumReadyForPurchase.No Then
                    Me.txtAutohirizationStatus.Text = "Unauthorized"
                ElseIf Me.ObjPR.GetReadyForPurchase(.Primary) = ClsTransaksiPR.EnumReadyForPurchase.Yes Then
                    Me.txtAutohirizationStatus.Text = "Authorized By " & Me.ObjUser.GetKodeUser(.PrimaryUserOtorisasi) & " on " & _
                    Format(.TimeOtorisasi, "dd MMM yyyy - HH.mm")
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

    'Private Sub Hitung_Total()
    '    Dim row As DataGridViewRow
    '    Dim SubTotal As Double = 0

    '    Try
    '        'hitung sub total
    '        For Each row In Me.dgvItem.Rows
    '            'yang diperhitungkan hanya yang tidak pake secondary unit
    '            If row.Cells(Me.colUserSecondaryUnits.Index).Value = Boolean.FalseString Then
    '                SubTotal += CDbl(row.Cells(Me.ColAmount.Index).Value)
    '            End If
    '        Next
    '        Me.txtSubTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(SubTotal, DefaultPrecision))
    '        'hitung pajak
    '        If Me.chkTax.Checked = True Then
    '            Me.txtTaxAmount.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(Me.txtTax.Text * SubTotal / 100, _
    '            DefaultPrecision))
    '        Else
    '            Me.txtTax.Text = "0"
    '            Me.txtTaxAmount.Text = "0"
    '        End If
    '        'hitung total = sub total + tax
    '        Me.TxtTotal.Text = Me.ObjCd.DigitGrouping(Me.ObjCd.SetDecimal(CDbl(Me.txtSubTotal.Text) + _
    '        CDbl(Me.txtTaxAmount.Text), DefaultPrecision))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk delivery date expectation yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColDeliveryDate.Index).Value) = False Then
                    If row.Cells(Me.ColDeliveryDate.Index).Value = "Multiple" Then
                        row.Cells(Me.ColDeliveryDate.Index).Value = "Multiple"
                    ElseIf row.Cells(Me.ColDeliveryDate.Index).Value = "#1/1/1950#" Then
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
                'digit grouping: qty
                If IsNumeric(row.Cells(Me.ColQty.Index).Value) = True Then
                    row.Cells(Me.ColQty.Index).Value = Me.ObjCd.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub llbAuthorizationDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthorizationDetail.LinkClicked
    '    If Me.TxtPrimaryStockIn.Text = "" Or Me.TxtPrimaryStockIn.Text = "-1" Then
    '        Exit Sub
    '    End If

    '    Dim i As New FrmTransaksiPO_AuthorizationDetail
    '    Dim jmlbaris As Integer = 0
    '    Dim row As DataGridViewRow
    '    Dim ObjStockIn As New Inventory_Module.clsTransaksiStockInOut
    '    Dim ObjPackingList As New ClsTransaksiPackingList
    '    Dim ObjPRC As New ClsTransaksiPOC
    '    Dim ObjAOPDP As New ClsTransaksiAOP

    '    Try
    '        i.dgvDetail.Rows.Clear()
    '        'tambahkan 3 baris authorization
    '        jmlbaris = Me.ObjPR.AuthorizationNeeded
    '        'masukkan baris dan detail data otorisasi
    '        For j As Integer = 1 To jmlbaris
    '            i.dgvDetail.Rows.Add()
    '            row = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)
    '            If row.Index = 0 Then
    '                row.Cells(1).Value = "Purchase Staff Authorization"
    '                If Me.ObjPR.PrimaryPurchaseStaffAuthorization <> -1 Then
    '                    row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPR.PrimaryPurchaseStaffAuthorization)
    '                    row.Cells(3).Value = Format(Me.ObjPR.PurchaseStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
    '                Else
    '                    row.Cells(2).Value = "-"
    '                    row.Cells(3).Value = "-"
    '                End If
    '            ElseIf row.Index = 1 Then
    '                row.Cells(1).Value = "Purchase Manager Authorization"
    '                If Me.ObjPR.PrimaryManagerStaffAuthorization <> -1 Then
    '                    row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPR.PrimaryManagerStaffAuthorization)
    '                    row.Cells(3).Value = Format(Me.ObjPR.PurchaseManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
    '                Else
    '                    row.Cells(2).Value = "-"
    '                    row.Cells(3).Value = "-"
    '                End If
    '            ElseIf row.Index = 2 Then
    '                row.Cells(1).Value = "Director Authorization"
    '                If Me.ObjPR.PrimaryDirectorStaffAuthorization <> -1 Then
    '                    row.Cells(2).Value = Me.ObjUser.GetNamaByPrimary(Me.ObjPR.PrimaryDirectorStaffAuthorization)
    '                    row.Cells(3).Value = Format(Me.ObjPR.PurchaseDirectorAuthorizationDate, "dd MMM yyyy - HH:mm")
    '                Else
    '                    row.Cells(2).Value = "-"
    '                    row.Cells(3).Value = "-"
    '                End If
    '            End If
    '        Next
    '        'tunjukkan form
    '        Me.ObjCd.SetNumber(i.dgvDetail)

    '        'set chk
    '        '1. stock in purchase
    '        If ObjStockIn.GetJmlByPrimaryPR(Me.TxtPrimaryStockIn.Text) > 0 Then
    '            i.chkStockIn.Checked = True
    '        Else
    '            i.chkStockIn.Checked = False
    '        End If
    '        'packing list
    '        If ObjPackingList.GetJmlItem_ByPrimaryPR(Me.TxtPrimaryStockIn.Text) > 0 Then
    '            i.chkPackingList.Checked = True
    '        Else
    '            i.chkPackingList.Checked = False
    '        End If
    '        'poc
    '        If ObjPRC.GetJmlByPO(Me.TxtPrimaryStockIn.Text) > 0 Then
    '            i.chkPOC.Checked = True
    '        Else
    '            i.chkPOC.Checked = False
    '        End If
    '        'purchase dp
    '        If ObjAOPDP.GetJml_AOPDP_ByPrimaryPR(Me.TxtPrimaryStockIn.Text) > 0 Then
    '            i.chkDP.Checked = True
    '        Else
    '            i.chkDP.Checked = False
    '        End If

    '        i.ShowDialog()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        i = Nothing
    '        ObjStockIn = Nothing
    '        ObjPackingList = Nothing
    '        ObjPRC = Nothing
    '        ObjAOPDP = Nothing
    '    End Try
    'End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindPO.Click
    '    Me.Close()
    'End Sub

    'Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
    '    If Me.TxtPrimaryStockIn.Text = "" Or Me.TxtPrimaryStockIn.Text = "-1" Then
    '        Exit Sub
    '    End If

    '    Dim ij As New FrmTransaksiPO_DetailSIgn

    '    Try
    '        ij.TxtPreparedBy.ReadOnly = True
    '        ij.TxtPreparedBy_Jabatan.ReadOnly = True
    '        ij.TxtApprovedBy.ReadOnly = True
    '        ij.TxtApprovedBy_Jabatan.ReadOnly = True
    '        ij.TxtAcknowledgeBy.ReadOnly = True
    '        ij.TxtAcknowledgeBy_Jabatan.ReadOnly = True
    '        ij.btnOK.Visible = False

    '        ij.TxtPreparedBy.Text = Me.ObjPR.PreparedBy
    '        ij.TxtPreparedBy_Jabatan.Text = Me.ObjPR.PreparedBy_Jabatan
    '        ij.TxtApprovedBy.Text = Me.ObjPR.ApprovedBy
    '        ij.TxtApprovedBy_Jabatan.Text = Me.ObjPR.ApprovedBy_Jabatan
    '        ij.TxtAcknowledgeBy.Text = Me.ObjPR.AcknowledgeBy
    '        ij.TxtAcknowledgeBy_Jabatan.Text = Me.ObjPR.AcknowledgeBy_Jabatan

    '        ij.ShowDialog()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ij = Nothing
    '    End Try
    'End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    'Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbRemove.Click
    '    'Dim row As DseTransaksiPR.tbltranspo_authorize_view2_2Row
    '    'Dim rowview As DataRowView
    '    Dim vFrm As New FrmTransaksiPOOtorisasi2
    '    Dim DTAPR As New DseTransaksiPRTableAdapters.tbltrans_poTableAdapter
    '    Dim CheckAuth As Boolean = False
    '    Dim Primary_ As Double = -1
    '    Try
    '        For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
    '            If rowDgv.Cells(Me.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString Then
    '                CheckAuth = True
    '                Exit For
    '            End If
    '        Next
    '        If CheckAuth = True Then
    '            If MsgBox("Remove PO from authorization form?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
    '                For Each rowDgv As DataGridViewRow In Me.DgvOtorisasi.Rows
    '                    If rowDgv.Cells(vFrm.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString Then
    '                        Primary_ = rowDgv.Cells(vFrm.ColDgvAuth_Primary.Index).Value
    '                        DTAPR.UpdateQuery_SetRemoved(1, Primary_)
    '                    End If
    '                Next

    '                Me.LoadOtorisasi()
    '            End If
    '        Else
    '            MsgBox("Please select item to remove from authorization form.", _
    '            MsgBoxStyle.Information, "Information")
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTAPR.Dispose()
    '    End Try
    'End Sub

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
                Primary_ = Me.DgvOtorisasi.Item(Me.ColDgvAuth_Primary.Index, Me.DgvOtorisasi.SelectedCells.Item(0).RowIndex).Value
                Me.ObjPR.CariDetail(Primary_, Me.dgvItem, Me.ObjCd, Me.vFrmPR, Me.vTBLItemDelivery)
                Me.SalinKeField()
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCancelOtorisasi_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCancelOtorisasi.SelectionChanged
        Dim Primary_ As Double
        Try
            If Me.DgvCancelOtorisasi.SelectedRows.Count > 0 Then
                Primary_ = Me.DgvCancelOtorisasi.Item(Me.ColDgvUnAuth_Primary.Index, Me.DgvCancelOtorisasi.SelectedCells.Item(0).RowIndex).Value
                Me.ObjPR.CariDetail(Primary_, Me.dgvItem, Me.ObjCd, Me.vFrmPR, Me.vTBLItemDelivery)
                Me.SalinKeField()
            Else
                Me.KosongkanField()
            End If
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
                        ''hanya PO yang sudah diotorisasi yang boleh diprint
                        'If Me.ObjPR.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjPR.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
                        '    MsgBox("This purchase order require more authorization, please check authorization needed.", _
                        '    MsgBoxStyle.Information, "Information")
                        '    Exit Sub
                        'Else
                        If ObjPR.CetakPR(Me.TxtPrimaryStockIn.Text, ij.crv1) = True Then
                            ij.ShowDialog()
                        End If
                        Exit Sub
                        'End If
                    End If
                End If

                If Me.TxtPrimaryStockIn.Text <> "" And Me.TxtPrimaryStockIn.Text <> "-1" Then
                    'hanya PO yang sudah diotorisasi yang boleh diprint
                    'If Me.ObjPR.GetAuthorizationLevel(Me.TxtPrimaryStockIn.Text) <> Me.ObjPR.GetAuthorizationLevelNeeded(Me.TxtPrimaryStockIn.Text) Then
                    '    MsgBox("This purchase order require more authorization, please check authorization needed.", _
                    '    MsgBoxStyle.Information, "Information")
                    '    Exit Sub
                    'Else
                    If ObjPR.CetakPR(Me.TxtPrimaryStockIn.Text, ij.crv1) = True Then
                        ij.ShowDialog()
                    End If
                    'End If
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
                Primary_ = Me.DgvAlreadyAuthorize.Item(Me.ColDgvAuth_Primary.Index, Me.DgvAlreadyAuthorize.SelectedCells.Item(0).RowIndex).Value
                Me.ObjPR.CariDetail(Primary_, Me.dgvItem, Me.ObjCd, Me.vFrmPR, Me.vTBLItemDelivery)
                Me.SalinKeField()
            Else
                Me.KosongkanField()
            End If
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

    Private Sub TxtNoPR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtNoPR.KeyDown
        Dim ObjPR As New ClsTransaksiPR
        Try
            If e.KeyValue = 13 Then
                If Me.TxtNoPR.Text <> "" Then
                    If Me.ObjPR.ApakahSudahAda(Me.TxtNoPR.Text) = True Then
                        Me.TxtPrimaryPR.Text = Me.ObjPR.GetPrimary(Me.TxtNoPR.Text)
                        Me.TxtNoPR.Text = Me.ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                        Me.ChkPRDate.Focus()
                    Else
                        Me.BtnFindPR.PerformClick()
                    End If
                Else
                    Me.BtnFindPR.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjPR = Nothing
        End Try
    End Sub

    Private Sub TxtNoPR_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoPR.Leave
        Dim hj As New FrmTransaksiPR_Cari
        Dim ObjPR As New ClsTransaksiPR
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtNoPR.Text <> "" Then
                If Me.ObjPR.ApakahSudahAda(Me.TxtNoPR.Text) = True Then
                    Me.TxtPrimaryPR.Text = Me.ObjPR.GetPrimary(Me.TxtNoPR.Text)
                    Me.TxtNoPR.Text = Me.ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                Else
                    If Me.TxtNoPR.Text <> "" Then
                        hj.txtCari.Text = Me.TxtNoPR.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryPR > 0 Then
                        Me.TxtPrimaryPR.Text = hj.PrimaryPR
                        Me.TxtNoPR.Text = Me.ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                        Me.CboStatusPR.Focus()
                    Else
                        Me.TxtPrimaryPR.Text = "-1"
                        Me.TxtNoPR.Clear()
                        Me.TxtNoPR.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.TxtNoPR.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindPR.Click
        Dim hj As New FrmTransaksiPR_Cari
        Dim ObjPR As New ClsTransaksiPR
        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtNoPR.Text <> "" Then
                hj.txtCari.Text = Me.TxtNoPR.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryPR > 0 Then
                Me.TxtPrimaryPR.Text = hj.PrimaryPR
                Me.TxtNoPR.Text = Me.ObjPR.GetNoTransaksi(Me.TxtPrimaryPR.Text)
                Me.ChkPRDate.Focus()
            Else
                Me.TxtPrimaryPR.Text = "-1"
                Me.TxtNoPR.Clear()
                Me.TxtNoPR.Focus()
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
            If Me.TxtNoPR.Enabled = True Then
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

    Private Sub ChkPODate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkPRDate.CheckedChanged
        Try
            If Me.ChkPRDate.Checked = True Then
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

    Private Sub TxtNoPR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoPR.GotFocus
        Me.TxtNoPR.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtNoPR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoPR.LostFocus
        Me.TxtNoPR.BackColor = Me.TxtPrimaryPR.BackColor
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

    'Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
    '    Dim ij As New Core_Module.FrmDetailAddressSupplier

    '    Try
    '        If Me.txtPrimaryReqBy.Text = "" Or Me.txtPrimaryReqBy.Text = "-1" Then
    '            If Me.TxtRemark.Enabled = True Then
    '                MsgBox("Please specify supplier first.", MsgBoxStyle.Information, "Information")
    '                Me.txtReqByCode.Focus()
    '                Exit Sub
    '            Else
    '                Exit Sub
    '            End If
    '        End If

    '        ij.vAddIdentifier = Me.txtSupplierAddIdentifier.Text
    '        ij.PrimarySup = Me.txtPrimaryReqBy.Text
    '        If Me.TxtRemark.Enabled = True Then
    '            ij.btnOK.BringToFront()
    '        Else
    '            ij.cboCustAddCode.Enabled = False
    '        End If
    '        ij.AcceptButton = ij.btnOK
    '        ij.btnExit.TabStop = False
    '        ij.ShowDialog()
    '        If ij.DialogResult = Windows.Forms.DialogResult.OK Then
    '            Me.ObjSupplier.IsiCboAddress(Me.cboSupplierAddress, Me.txtPrimaryReqBy.Text)
    '            Me.cboSupplierAddress.Text = Me.ObjSupplier.GetKodeAddress(ij.vAddIdentifierResult)
    '        End If
    '        ij.Dispose()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub
End Class