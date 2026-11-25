Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiPOC
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public StrictEditting As Boolean = True

    Public FormNo As Integer = 0
    Public vFromQuick As Boolean = False
    Public vStopCancellation As Boolean = False
    Public vDirectlySave As Boolean = False
    Public vShowDialog As Boolean = True

    Dim baru As Boolean
    Dim ObjPOC As New ClsTransaksiPOC
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjPOCItem As New ClsTransaksiPOCItem
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjPOItem As New ClsTransaksiPOItem
    Dim ObjUser As New ClsUser

    Dim IsInventoryInstalled As Boolean = False

    'Dim TBL As New DseTransaksiPO.orderbalance_report_view0DataTable
    'Dim TBL2 As New DseTransaksiPOC.tbltrans_poc_itemview2DataTable
    'Dim TBL3 As New DseTransaksiPO.orderbalance_report_view0DataTable
    'Dim TBL4 As New DseTransaksiPOC.tbltrans_poc_itemview2DataTable

    ''@rosadinnn
    ''@MUaT
    Dim TBL As New DseTransaksiPO.orderbalance_report_view0MultiDataTable
    Dim TBL2 As New DseTransaksiPOC.tbltrans_poc_itemview2_MultiDataTable
    Dim TBL3 As New DseTransaksiPO.orderbalance_report_view0MultiDataTable
    Dim TBL4 As New DseTransaksiPOC.tbltrans_poc_itemview2_MultiDataTable

    Dim PersenPajakAcuan, PrimaryCurrencySO As Decimal
    Dim SamaPO As Boolean = False
    Dim SamaCodeNo As Boolean = False
    Dim buffer As String = ""
    Dim bufferProduct As String = ""
    Dim BufferPO As String = ""

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub txtPONO_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONO.GotFocus
        txtPONO.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub txtPONO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONO.LostFocus
        txtPONO.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub txtCancellationReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCancellationReason.GotFocus
        txtCancellationReason.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub txtCancellationReason_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCancellationReason.LostFocus
        txtCancellationReason.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        TxtRemark.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        TxtRemark.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.GotFocus
        TxtQtyperBatchMulti.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        TxtQtyperBatchMulti.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        TxtRemarkItem.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        TxtRemarkItem.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub


    Private Sub AktifkanDetail()
        Try
          

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False

            If Me.dgvItem.RowCount > 0 Then
                Me.txtSupplierCode.Enabled = False
                'Me.btnFindSupplier.Enabled = False
            End If

            If Me.txtPrimaryPO.Text <> "-1" And Me.txtPrimaryPO.Text <> "" Then
                Me.SetGboItem()
            End If

            If Me.baru = True Then
                Me.txtPONO.Enabled = True
                Me.btnFindPO.Enabled = True
            Else
                Me.txtPONO.Enabled = False
                Me.btnFindPO.Enabled = False
            End If

            ''@rosadinnn
            ''@MUaT
            Me.FormatDGV()

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPOC.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryPOC.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryPOC.BackColor, False)

            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryPOC.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryPOC.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryPOC.BackColor, False)

            If Me.txtPrimaryPO.Text <> -1 Then
                Me.SetGboItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)

            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)

            Me.KosongkanItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)

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

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)

            Me.txtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"

            ''@rosadinnn
            ''@MUaT
            Me.txtPrimaryUnitsMulti.Text = "-1"

            'Me.txtPrimaryPO.Text = "-1"
            Me.TxtQtyperBatchMulti.Text = "0"
            Me.lblUnitsMulti.Text = "Units"
            Me.cboCodeNo.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '9. Bila setelah diinput jumlah sales order balance menjadi minus
            Dim JmlObjLama, JmlPOBBaru, JmlDGV, JmlPOB As Decimal
            JmlDGV = Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
            JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text)
            JmlObjLama = Me.ObjPOCItem.GetJmlPOCAktif_Multi(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryPOC.Text)

            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchMulti.Text)
            If JmlPOBBaru < 0 Then
                MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB + JmlObjLama - JmlDGV), ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) + " " + _
                        Me.lblUnitsMulti.Text + ".", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '10. Dalam 1 Packing List bisa ada beberapa PO dengan syarat: primary currency dan pajaknya sama
            'If Me.dgvItem.RowCount > 0 Then
            'Dim mTaxNew, mprimaryCurrency as Decimal
            'Me.ObjPO.GetPrimaryCurrency_Tax(Me.txtPrimaryPO.Text, mTaxNew, mprimaryCurrency)

            '   a. % pajak yang digunakan sama
            'If mTaxNew <> Me.PersenPajakAcuan Then
            'MsgBox("This purchase order item cannot be in one packing list with the existing items. " + _
            '"For several purchase order item written in one packing list their currency and tax must be identical", _
            'MsgBoxStyle.Information, "Information")
            'Return False
            'End If
            '   b. primary currency sama
            'If mprimaryCurrency <> Me.PrimaryCurrencySO Then
            'MsgBox("This purchase order item cannot be in one packing list with the existing items. " + _
            '"For several purchase order item written in one packing list their currency and tax must be identical", _
            ' MsgBoxStyle.Information, "Information")
            'Return False
            'End If
            'End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryPO As Decimal) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo _
                 Then ''row.Cells(Me.ColPONo.Index).Value = vPrimaryPO
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml(ByVal vPrimaryCodeNo As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryPO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryPO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDbl(row.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("POC number")
            KoleksiTextBox.Add(Me.TxtPOCNo)
            KoleksiNamaData.Add("supplier code")
            KoleksiTextBox.Add(Me.txtSupplierCode)
            KoleksiNamaData.Add("cancellation reason")
            KoleksiTextBox.Add(Me.txtCancellationReason)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.ObjPOC.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjPOC.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtPOCNo.Text = Me.ObjPOC.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjPOC.TanggalTransaksi
                        Me.TxtPOCNo.Text = Me.ObjPOC.NoTransaksi
                        Return False
                    End If
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
            With Me.ObjPOC
                .primary = Me.TxtPrimaryPOC.Text
                .primaryPO = Me.txtPrimaryPO.Text
                .NoTransaksi = Me.TxtPOCNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                If Me.baru = True Then
                    .Authorizationlevel = 0
                    .AuthorizationNeeded = 0
                    .PurchaseDirectorAuthorizationDate = #1/1/1950#
                    .PrimaryDirectorStaffAuthorization = -1
                    .PurchaseManagerAuthorizationDate = #1/1/1950#
                    .PrimaryManagerStaffAuthorization = -1
                    .PurchaseStaffAuthorizationDate = #1/1/1950#
                    .PrimaryPurchaseStaffAuthorization = -1
                End If
                .CancellationReason = Me.txtCancellationReason.Text
                .Remarks = Me.TxtRemark.Text

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjPOC
                Me.TxtPrimaryPOC.Text = .primary
                Me.txtPrimaryPO.Text = .primaryPO
                Me.TxtPOCNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi

                If Me.ObjPOC.Authorizationlevel = 0 Then
                    Me.txtAutohirizationStatus.Text = "Unauthorized"
                ElseIf Me.ObjPOC.Authorizationlevel = 1 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Purchase Staff"
                ElseIf Me.ObjPOC.Authorizationlevel = 2 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Purchase Manager"
                ElseIf Me.ObjPOC.Authorizationlevel = 3 Then
                    Me.txtAutohirizationStatus.Text = "Authorized By Director"
                End If

                Me.txtCancellationReason.Text = .CancellationReason
                Me.TxtRemark.Text = .Remarks

                Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                Dim SupName, CPName, PhoneNo As String
                SupName = ""
                CPName = ""
                PhoneNo = ""

                ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = SupName
                Me.txtContactPerson.Text = CPName
                Me.txtPhoneNo.Text = PhoneNo
                Me.txtPrimaryCodeNo.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                Me.txtPONO.Text = ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)

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

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjLS As New clsMasterListedSoftware
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.vFromQuick = False Then
                Me.baru = True
                Me.NonAktifkanDetail()
            End If

            If Me.vFromQuick = True Then
                Me.prgProgress.Visible = True
                Me.btnStopProcess.Visible = True
                Me.btnStopProcess.TabStop = True
                Me.lblProgress.Visible = True
                Me.btnFind.Visible = False
                Me.btnFind.TabStop = False
                Me.btnNew.Visible = False
                Me.btnNew.TabStop = False
                Me.btnEdit.Visible = False
                Me.btnEdit.TabStop = False
                Me.btnDelete.Visible = False
                Me.btnDelete.TabStop = False
                Me.btnExit.Visible = False
                Me.btnExit.TabStop = False
                Me.MinimizeBox = False
                Me.txtPONO.Enabled = False
                Me.btnFindPO.Enabled = False
                If vShowDialog = True Then
                    If IsNumeric(Me.txtPrimaryPO.Text) = True Then
                        If Me.txtPrimaryPO.Text > 0 Then
                            For ij As Integer = Me.dgvItem.Rows.Count - 1 To 0 Step -1
                                rowDGV = Me.dgvItem.Rows(ij)
                                If rowDGV.Cells(Me.ColQtyMulti.Index).Value = 0 Then
                                    Me.dgvItem.Rows.RemoveAt(ij)
                                End If
                            Next
                            Me.IsiCboProduct(Me.txtPrimaryPO.Text)
                            Me.ObjCD.SetNumber(Me.dgvItem)
                        End If
                    End If
                End If
                If vDirectlySave = True Then
                    Me.SimpanLangsung()
                End If
            End If

            Me.cboCodeNo.Width = Me.SplitContainer2.Panel1.Width - 5 - Me.cboCodeNo.Left
            Me.txtDesc.Width = Me.cboCodeNo.Width
            Me.BtnInput.Left = Me.SplitContainer2.Panel2.Width - 5 - Me.BtnInput.Width
            Me.lblUnitsDefault.Left = Me.BtnInput.Left
            Me.lblUnitsMulti.Left = Me.BtnInput.Left
            Me.TxtRemarkItem.Width = Me.BtnInput.Left - 5 - Me.TxtRemarkItem.Left
            Me.TxtQtyperBatchMulti.Width = Me.TxtRemarkItem.Width
            Me.TxtQtyperBatchDefault.Width = Me.TxtRemarkItem.Width
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjLS = Nothing
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
                        If Me.ObjPOC.SimpanBaru(Me.dgvItem, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()

                            If vFromQuick = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPOC.SimpanPerubahan(Me.dgvItem, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
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

    Private Sub SimpanLangsung()
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.ObjPOC.SimpanBaru(Me.dgvItem, Me.UserCode) Then
                        Me.KosongkanItem()
                        Me.SalinKeField()
                        Me.NonAktifkanDetail()
                        Me.btnFind.Focus()

                        If vFromQuick = True Then
                            Me.Close()
                        End If
                    End If
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Me.NonAktifkanDetail()
                    Me.vStopCancellation = True
                    Me.Close()
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
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
                Me.KosongkanItem()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjPOC.CariDetail(Me.TxtPrimaryPOC.Text, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                If Me.vFromQuick = True Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.SetBaru()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetBaru()
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.KosongkanItem()
            Me.dtpDate.Focus()
            Me.ObjPOC.ResetObject(Me.dgvItem)
            Me.TxtPOCNo.Text = Me.ObjPOC.GetNoBaru(Me.dtpDate.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            If Me.vFromQuick = True Then
                Me.btnStopProcess.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim i As New FrmCariTransaksiPOC

        Try
            i.BackColorAll = Me.BackColorAll
            i.ShowDialog()

            If Not i.PrimaryTransaksi = -1 Then
                If Me.ObjPOC.CariDetail(i.PrimaryTransaksi, Me.dgvItem) = True Then
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
            i.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPOCNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPOCNo.Text <> "" Then
                If Me.ObjPOC.ApakahBolehEdit(Me.TxtPrimaryPOC.Text, Me.StrictEditting, Me.UserCode, True) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
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
            If Me.TxtPOCNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPOCNo.Text <> "" Then
                If Me.ObjPOC.ApakahBolehHapus(Me.TxtPrimaryPOC.Text, _
                Me.StrictEditting, Me.UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPOC.Hapus(Me.TxtPrimaryPOC.Text) = True Then
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
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjPOC.SimpanBaru(Me.dgvItem, Me.UserCode) = True Then
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
                            If Me.ObjPOC.SimpanPerubahan(Me.dgvItem, Me.UserCode) = True Then
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

    Private Sub IsikanTBLSO()
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiPOC.tbltrans_poc_itemview2Row

            For Each row In Me.dgvItem.Rows
                Dim ketemu As Boolean = False
                For Each row1 In Me.TBL2.Rows
                    If row1.Primary_Supplier = Me.txtPrimarySupplier.Text And _
                    row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                        'row1.Primary_PO = row.Cells(Me.ColPrimaryPO.Index).Value Then
                        'kalo ketemu ubah data order balance
                        row1.Total_Qty = row1.Total_Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                        ketemu = True
                        Exit For
                    End If
                Next
                'kalo ga ketemu tambahkan baris baru
                If ketemu = False Then
                    Me.TBL2.Rows.Add(Me.TxtPrimaryPOC.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    Me.txtPrimaryPO.Text, row.Cells(Me.ColTotalQty.Index).Value, Me.txtPrimarySupplier.Text, _
                    row.Cells(Me.ColCodeNo.Index).Value, Me.txtPONO.Text)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboProduct(ByVal vPrimaryPO As Decimal)
        Try
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan dari basis data-persedian terakhir
            Me.cboCodeNo.DataSource = Nothing
            Me.ObjPOItem.IsikanTBLProduct_Multi(Me.txtPrimaryPO.Text, Me.TBL3, Me.cboCodeNo, True)

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiPOC.tbltrans_poc_itemview2_MultiRow
            Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjPOCItem.IsiTBLProductByPO_Multi(Me.TxtPrimaryPOC.Text, Me.TBL4, Me.txtPrimaryPO.Text)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    For Each row2 In Me.TBL3.Rows
                        If row1.Primary_Supplier = row2.Primary_Supplier _
                        And row1.Primary_Product = row2.Primary_Product And _
                        row2.Primary_PO = row1.Primary_PO Then
                            row2.OrderBalance = row2.OrderBalance + row1.Total_Qty

                            ''@rosadinnn
                            ''@MUaT
                            row2.OrderBalance_Multi = row2.OrderBalance_Multi + row1.Total_Qty_Multi

                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru

                        ''@rosadinnn
                        ''@MUaT
                        row2 = Me.TBL3.Neworderbalance_report_view0MultiRow
                        With row2
                            .Primary_PO = row1.Primary_PO
                            .Primary_Supplier = row1.Primary_Supplier
                            .Primary_Product = row1.Primary_Product
                            .Product_Code = row1.Product_Code
                            .PO_Number = row1.PO_Number
                            .OrderBalance = row1.Total_Qty
                            .Primary_Quick_Acuan = -1
                            .OrderBalance_Multi = row1.Total_Qty_Multi
                            .Primary_Unit_Multi = row1.Primary_Unit_Multi
                            .Unit_Code_Multi = row1.Unit_Code_Multi
                            .Ratio_Default = row1.Ratio_Default
                            .Ratio_Reference = row1.Ratio_Reference
                            .Primary_Unit_Default = row1.Primary_Units
                            .Unit_Code_Default = row1.Unit_Code
                            .Tolerance_Down = 0
                            .Tolerance_Up = 0
                        End With
                        Me.TBL3.Rows.Add(row2)
                        'Me.TBL3.Rows.Add(row1.Primary_PO, row1.Primary_Supplier, row1.Primary_Product, _
                        'row1.Product_Code, row1.PO_Number, row1.Total_Qty)

                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsiTBLProduct_Multi(Me.txtPrimaryPO.Text)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                For Each row2 In Me.TBL3.Rows
                    If row1.Primary_Supplier = row2.Primary_Supplier _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_PO = row1.Primary_PO Then

                        ''@rosadinnn
                        ''@MUaT
                        row2.OrderBalance = row2.OrderBalance - row1.Total_Qty
                        row2.OrderBalance_Multi = row2.OrderBalance_Multi - row1.Total_Qty_Multi
                        'If row2.OrderBalance = 0 Then
                        '    Me.TBL3.Rows.Remove(row2)
                        'End If
                        If row2.OrderBalance_Multi = 0 Then
                            Me.TBL3.Rows.Remove(row2)
                        End If

                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL3.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4.Rows.Remove(row1)
            End While

            Me.TBL3.DefaultView.Sort = "[Product Code] ASC"
            Me.cboCodeNo.DataSource = Me.TBL3.DefaultView
            Me.cboCodeNo.DisplayMember = "Product Code"
            Me.cboCodeNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTBLProduct_Default(ByVal vPrimaryPO As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiPOC.tbltrans_poc_itemview2Row
            For Each row In Me.dgvItem.Rows
                If Me.txtPrimaryPO.Text = vPrimaryPO Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Primary_Supplier = Me.txtPrimarySupplier.Text And _
                        row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                            'row1.Primary_PO = row.Cells(Me.ColPrimaryPO.Index).Value Then
                            'kalo ketemu ubah data order balance
                            row1.Total_Qty = row1.Total_Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL4.Rows.Add(Me.TxtPrimaryPOC.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        Me.txtPrimaryPO.Text, row.Cells(Me.ColTotalQty.Index).Value, _
                        Me.txtPrimarySupplier.Text, row.Cells(Me.ColCodeNo.Index).Value, Me.txtPONO.Text)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTBLProduct_Multi(ByVal vPrimaryPO As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiPOC.tbltrans_poc_itemview2_MultiRow
            For Each row In Me.dgvItem.Rows
                If Me.txtPrimaryPO.Text = vPrimaryPO Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Primary_Supplier = Me.txtPrimarySupplier.Text And _
                        row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                            'row1.Primary_PO = row.Cells(Me.ColPrimaryPO.Index).Value Then
                            'kalo ketemu ubah data order balance
                            row1.Total_Qty = row1.Total_Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)

                            ''@rosadinnn
                            ''@MUaT
                            row1.Total_Qty_Multi = row1.Total_Qty_Multi + CDbl(row.Cells(Me.ColQtyMulti.Index).Value)

                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then

                        ''@rosadinnn
                        ''@MUaT
                        row1 = Me.TBL4.Newtbltrans_poc_itemview2_MultiRow
                        With row1
                            .PrimaryPOC = Me.TxtPrimaryPOC.Text
                            .Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                            .Primary_PO = Me.txtPrimaryPO.Text
                            .Total_Qty = row.Cells(Me.ColTotalQty.Index).Value
                            .Primary_Supplier = Me.txtPrimarySupplier.Text
                            .PO_Number = Me.txtPONO.Text
                            .Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                            .Total_Qty_Multi = row.Cells(Me.ColQtyMulti.Index).Value
                            .Primary_Unit_Multi = row.Cells(Me.ColPrimaryUnitMulti.Index).Value
                            .Unit_Code_Multi = row.Cells(Me.ColUnitCodeMulti.Index).Value
                            .Primary_Units = row.Cells(Me.ColPrimaryUnits.Index).Value
                            .Unit_Code = row.Cells(Me.ColUnit.Index).Value
                            .Ratio_Reference = row.Cells(Me.ColRatioReference.Index).Value
                            .Ratio_Default = row.Cells(Me.ColRatioDefault.Index).Value
                        End With
                        Me.TBL4.Rows.Add(row1)
                        'Me.TBL4.Rows.Add(Me.TxtPrimaryPOC.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        'Me.txtPrimaryPO.Text, row.Cells(Me.ColTotalQty.Index).Value, _
                        'Me.txtPrimarySupplier.Text, row.Cells(Me.ColCodeNo.Index).Value, Me.txtPONO.Text)

                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItem()
        Try
            'Me.cboPurchaseOrder.DataSource = Nothing
            Me.cboCodeNo.DataSource = Nothing

            'sales order & product code
            Me.IsiCboProduct(Me.txtPrimaryPO.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim RowDgv As DataGridViewRow

        Try
            If Me.ApakahDataItemBenar = True Then
                Me.txtSupplierCode.Enabled = False
                'Me.btnFindSupplier.Enabled = False

                Dim Index As Integer = Me.GetIndex(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)

                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)

                    ''@rosadinnn
                    ''@MUaT
                    'row.Cells(7).Value = CDbl(row.Cells(7).Value) + CDbl(Me.TxtQtyperBatchMulti.Text)
                    row.Cells(Me.ColQtyMulti.Index).Value = CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatchMulti.Text)
                    Me.SetQtyDefaultOnDgv(row.Index)

                Else

                    ''@rosadinnn
                    ''@MUaT
                    'Me.dgvItem.Rows.Add("", Me.cboCodeNo.Text, Me.txtDesc.Text, 0, _
                    '0, "", 0, Me.TxtQtyperBatchMulti.Text, Me.lblUnitsMulti.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimaryCodeNo.Text, -1, Me.txtPrimaryUnitsDefault.Text)
                    Me.dgvItem.Rows.Add()
                    RowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnit.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColTotalQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryMasterUnit.Index).Value = -1
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                        .Cells(Me.ColUnitCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                        If Me.txtPrimaryUnitsMulti.Text <> Me.txtPrimaryUnitsDefault.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                                    ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) _
                                    & " " & Me.lblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                                    Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)) & " " & Me.lblUnitsDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        .Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                        .Cells(Me.ColPrimaryUnitMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                    End With

                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If Me.dgvItem.RowCount = 1 Then
                        'set nilai acuan untuk pengecekan SO
                        Me.ObjPO.GetPrimaryCurrency_Tax(Me.txtPrimaryPO.Text, _
                        Me.PersenPajakAcuan, Me.PrimaryCurrencySO)
                    End If
                End If

                'kendalikan focus item selanjutnya
                Dim JmlObjLama, JmlPOBBaru, JmlDGV, JmlPOB As Decimal
                JmlDGV = Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text)
                JmlObjLama = Me.ObjPOCItem.GetJmlPOCAktif_Multi(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryPOC.Text)

                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV
                Me.TxtQtyperBatchMulti.Text = "0"
                Me.TxtRemarkItem.Text = ""

                If JmlPOBBaru <= 0 Then
                    'remove row, pada code no ybs
                    Dim rowView As DataRowView = Me.cboCodeNo.SelectedItem
                    Dim row As DseTransaksiPO.orderbalance_report_view0MultiRow = rowView.Row
                    Me.TBL3.Rows.Remove(row)
                    Me.cboCodeNo_SelectedValueChanged(sender, e)
                    Me.cboCodeNo.Focus()
                    If Me.cboCodeNo.Items.Count = 0 Then
                        Me.SetGboItem()
                        Me.cboCodeNo.Focus()
                    End If
                Else
                    'focus pada cbo code no
                    Me.cboCodeNo.Focus()
                End If

                'format dgv
                Me.FormatDGV()

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                'no po harus diset enable false
                Me.txtPONO.Enabled = False
                Me.btnFindPO.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim IsMultiUnit As Boolean = False
        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                'digit grouping
                row.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value)
            Next

            ''@rosadinnn
            ''@MUaT
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnits.Index).Value <> .Cells(Me.ColPrimaryUnitMulti.Index).Value Then
                        IsMultiUnit = True
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.ColTotalQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColTotalQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then

                ''@rosadinnn
                ''@MUaT
                ''@tambah kolom qty multi
                If e.ColumnIndex = Me.ColRemark.Index Or e.ColumnIndex = Me.ColQtyMulti.Index Then
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim JmlDGV As Decimal
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If

                    If e.ColumnIndex <> Me.ColRemark.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit

                        ''@rosadinnn
                        ''@MUaT
                        JmlDGV = 0
                        If e.ColumnIndex = Me.ColQtyMulti.Index Then
                            Me.EndEditQty_Multi(sender, e, JmlDGV, row)
                        ElseIf e.ColumnIndex = Me.ColTotalQty.Index Then
                            Me.EndEditQty_Default(sender, e, JmlDGV, row)
                        End If
                        
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditQty_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal JmlDGV As Decimal, ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'cek apakah sales order balance menjadi minus setelah perubahan data
                Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri As Decimal
                JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(CDbl(Me.txtPrimaryPO.Text), .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlObjLama = Me.ObjPOCItem.GetJmlPOCAktif_Multi(Me.txtPrimaryPO.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, Me.TxtPrimaryPOC.Text)

                ''@rosadinnn
                ''@MUaT
                JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value
                'JmlDGV = Me.GetJmlDGV(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                                        Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                                        Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value))

                'perhitungan menggunakan nilai jumlah dgv yang baru
                ''JmlPOBBaru = JmlPOB - JmlDGV
                JmlPOBBaru = JmlPOB + JmlObjLama - JmlSendiri
                If JmlPOBBaru < 0 Then  '' CStr(JmlPOB - JmlDGV + JmlSendiri)
                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB + JmlObjLama), ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)) + _
                    " " + Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value + ".", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'jika jumlah menurun maka harus load ulang data
                Me.SetGboItem()

                ''@rosadinnn
                ''@MUaT
                Me.SetQtyDefaultOnDgv(e.RowIndex)
            End With
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
                    '2 baris sama apabila semua data sama kecuali: data qty, total qty, qty per master
                    'warehouse

                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = row1.Cells(Me.ColPrimaryCodeNo.Index).Value _
                    Then
                        Return True
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
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = row1.Cells(Me.ColPrimaryCodeNo.Index).Value _
                     Then
                        Me.GabungkanBaris(row, row1)
                        Me.ObjCD.SetNumber(Me.dgvItem)
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

    Private Sub GabungkanBaris(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow)
        Try
            'gabungkan data qty.
            row.Cells(Me.ColTotalQty.Index).Value = CDbl(row.Cells(Me.ColTotalQty.Index).Value) _
            + CDbl(row1.Cells(Me.ColTotalQty.Index).Value)
            'gabungkan data remarks
            If CStr(row.Cells(Me.ColRemark.Index).Value).EndsWith(".") = False Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(Me.ColRemark.Index).Value = CStr(row.Cells(Me.ColRemark.Index).Value) + ". "
            End If
            row.Cells(Me.ColRemark.Index).Value = CStr(row.Cells(Me.ColRemark.Index).Value) + CStr(row1.Cells(Me.ColRemark.Index).Value)
            'hapus baris ybs
            Me.dgvItem.Rows.RemoveAt(row1.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
            If Me.TxtRemark.Enabled = True Then
                Me.txtSupplierCode.Enabled = True
                'Me.btnFindSupplier.Enabled = True
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
            Dim row1 As DseTransaksiPO.orderbalance_report_view0MultiRow

            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)

            'cek sampai dimana perlu load ulang data
            If Me.txtPrimaryPO.Text = "" Or Me.txtPrimaryPO.Text = "-1" Then
                Me.SetGboItem()
            Else
                If Me.SamaPO = True Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL3.Rows
                        If row1.Primary_Product = Me.bufferProduct Then
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'load ulang data cbo product
                        Me.IsiCboProduct(Me.BufferPO)
                    End If
                Else
                    'cari apakah ada no SO yang baru saja dihapuskan pada tabel acuan, untuk mengetahui apakah
                    'perlu load ulang data dari so
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL.Rows
                        If row1.Primary_PO = Me.BufferPO Then
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        Me.SetGboItem()
                    End If
                End If
            End If

            'format tampilan
            Me.FormatDGV()

            'jika dgv kosong, maka po bisa diganti
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.txtPONO.Enabled = False
                    Me.btnFindPO.Enabled = False
                Else
                    Me.txtPONO.Enabled = True
                    Me.btnFindPO.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'tidak perlu cek akibatnya pada sales order balance
            'ketika user akan menghapus data, tidak perlu cek apakah akibatnya akan mengakibatkan ada stock yang minus, karena akan mengakibatkan
            'stock menjadi positif
            'tetapi untuk transaksi yang berkaitan dengan transaksi lainnya (DO, Invoice, GRN atau PO) harus cek validitas jumlah berkaitan dengan mereka
            'tandai apakah pada baris yang sama no SO-nya sama

            'If Me.txtPrimaryPO.Text = CStr(Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.Row.Index).Value) Then
            'Me.SamaPO = True
            'Else
            'Me.SamaPO = False
            'End If

            If Me.txtPrimaryCodeNo.Text = CStr(Me.dgvItem(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) Then
                Me.SamaCodeNo = True
            Else
                Me.SamaCodeNo = False
            End If

            'Me.BufferPO = CStr(Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.Row.Index).Value)
            Me.bufferProduct = CStr(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPONO.KeyDown
        If Me.txtPONO.Enabled = False Or Me.txtPONO.ReadOnly = True Then
            Exit Sub
        End If
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll
        Dim SupName, CPName, PhoneNo As String
        SupName = ""
        CPName = ""
        PhoneNo = ""

        Try
            If e.KeyValue = 13 Then
                If Me.ObjPO.ApakahSudahAda(Me.txtPONO.Text) = True Then
                    Dim PrimaryPO As Double = -1
                    PrimaryPO = Me.ObjPO.GetPrimary(Me.txtPONO.Text)
                    If Me.ObjPO.ApakahSudahOtorisasi(PrimaryPO) = True Then
                        Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONO.Text)
                        Me.txtPONO.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                        Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                        ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = SupName
                        Me.txtContactPerson.Text = CPName
                        Me.txtPhoneNo.Text = PhoneNo
                        Me.txtCancellationReason.Focus()
                        Me.SetGboItem()
                    Else
                        MsgBox("This purchase order has not been authorized, you cannot cancel this puchase order.", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtPONO.Clear()
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtContactPerson.Clear()
                        Me.txtPhoneNo.Clear()
                        Me.txtPONO.Focus()
                    End If
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtPONO.Text <> "" Then
                        i.txtCari.Text = Me.txtPONO.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryPO > 0 Then
                        If Me.ObjPO.ApakahSudahOtorisasi(i.PrimaryPO) = True Then
                            Me.txtPrimaryPO.Text = i.PrimaryPO 'Me.ObjPO.GetPrimary(Me.txtPONO.Text)
                            Me.txtPONO.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                            Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)

                            ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                            Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                            Me.TxtSupplierName.Text = SupName
                            Me.txtContactPerson.Text = CPName
                            Me.txtPhoneNo.Text = PhoneNo
                            Me.txtCancellationReason.Focus()
                            Me.SetGboItem()
                        Else
                            MsgBox("This purchase order has not been authorized, you cannot cancel this puchase order.", _
                            MsgBoxStyle.Information, "Information")
                            Me.txtPONO.Clear()
                            Me.txtSupplierCode.Clear()
                            Me.TxtSupplierName.Clear()
                            Me.txtPrimarySupplier.Text = "-1"
                            Me.txtContactPerson.Clear()
                            Me.txtPhoneNo.Clear()
                            Me.txtPONO.Focus()
                        End If
                    Else
                        Me.txtPONO.Clear()
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtContactPerson.Clear()
                        Me.txtPhoneNo.Clear()
                        Me.txtCancellationReason.Focus()
                        Me.KosongkanItem()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPONO.Leave
        If Me.txtPONO.Enabled = False Or Me.txtPONO.ReadOnly = True Then
            Exit Sub
        End If
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll
        Dim SupName, CPName, PhoneNo As String
        SupName = ""
        CPName = ""
        PhoneNo = ""

        Try
            If Me.txtPONO.Text <> "" Then
                If Me.ObjPO.ApakahSudahAda(Me.txtPONO.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjPO.GetPrimary(Me.txtPONO.Text)
                    If Me.ObjPO.ApakahSudahOtorisasi(PrimaryBaru) = True Then
                        If PrimaryBaru <> Me.txtPrimaryPO.Text Then
                            Me.txtPrimaryPO.Text = Me.ObjPO.GetPrimary(Me.txtPONO.Text)
                            Me.txtPONO.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                            Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                            ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                            Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                            Me.TxtSupplierName.Text = SupName
                            Me.txtContactPerson.Text = CPName
                            Me.txtPhoneNo.Text = PhoneNo
                            Me.txtCancellationReason.Focus()
                            Me.SetGboItem()
                        End If
                    Else
                        MsgBox("This purchase order has not been authorized, you cannot cancel this puchase order.", _
                        MsgBoxStyle.Information, "Information")
                        Me.txtPONO.Clear()
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtContactPerson.Clear()
                        Me.txtPhoneNo.Clear()
                        Me.txtPONO.Focus()
                    End If
                Else
                    If Me.txtPONO.Text <> "" Then
                        i.txtCari.Text = Me.txtPONO.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryPO > 0 Then
                        If Me.ObjPO.ApakahSudahOtorisasi(i.PrimaryPO) = True Then
                            Me.txtPrimaryPO.Text = i.PrimaryPO
                            Me.txtPONO.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                            Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                            ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                            Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                            Me.TxtSupplierName.Text = SupName
                            Me.txtContactPerson.Text = CPName
                            Me.txtPhoneNo.Text = PhoneNo
                            Me.txtCancellationReason.Focus()
                            Me.SetGboItem()
                        Else
                            MsgBox("This purchase order has not been authorized, you cannot cancel this puchase order.", _
                            MsgBoxStyle.Information, "Information")
                            Me.txtPONO.Clear()
                            Me.txtSupplierCode.Clear()
                            Me.TxtSupplierName.Clear()
                            Me.txtPrimarySupplier.Text = "-1"
                            Me.txtContactPerson.Clear()
                            Me.txtPhoneNo.Clear()
                            Me.txtPONO.Focus()
                        End If
                    Else
                        Me.txtPONO.Clear()
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtContactPerson.Clear()
                        Me.txtPhoneNo.Clear()
                        Me.txtCancellationReason.Focus()
                        Me.KosongkanItem()
                    End If
                End If
            Else
                Me.txtPONO.Clear()
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
                Me.txtCancellationReason.Focus()
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPO.Click
        Dim i As New FrmCariTransaksiPO
        i.BackColorAll = Me.BackColorAll
        Dim SupName, CPName, PhoneNo As String
        SupName = ""
        CPName = ""
        PhoneNo = ""

        Try
            '  i.cboCriteria.SelectedIndex = 0
            If Me.txtPONO.Text <> "" Then
                i.txtCari.Text = Me.txtPONO.Text
            End If
            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                If Me.ObjPO.ApakahSudahOtorisasi(i.PrimaryPO) = True Then
                    Me.txtPrimaryPO.Text = i.PrimaryPO
                    Me.txtPONO.Text = Me.ObjPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
                    Me.txtPrimarySupplier.Text = ObjPO.GetPrimarySupplier(Me.txtPrimaryPO.Text)
                    ObjSupplier.Get_SuppName_ContactPerson_PhoneNo(Me.txtPrimarySupplier.Text, SupName, CPName, PhoneNo)
                    Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                    Me.TxtSupplierName.Text = SupName
                    Me.txtContactPerson.Text = CPName
                    Me.txtPhoneNo.Text = PhoneNo
                    Me.txtCancellationReason.Focus()
                    Me.SetGboItem()
                Else
                    MsgBox("This purchase order has not been authorized, you cannot cancel this puchase order.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtPONO.Clear()
                    Me.txtSupplierCode.Clear()
                    Me.TxtSupplierName.Clear()
                    Me.txtPrimarySupplier.Text = "-1"
                    Me.txtContactPerson.Clear()
                    Me.txtPhoneNo.Clear()
                    Me.txtPONO.Focus()
                End If
            Else
                Me.txtPONO.Clear()
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
                Me.txtCancellationReason.Focus()
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    ''Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
    ''    Try
    ''        If Me.cboCodeNo.SelectedIndex <> -1 Then
    ''            Dim RowView As DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
    ''            Dim row2 As DseTransaksiPO.orderbalance_report_view0Row = RowView.Row
    ''            Dim PrimaryCodeNoBaru As Double = row2.Primary_Product

    ''            If Me.txtPrimaryCodeNo.Text <> PrimaryCodeNoBaru Then
    ''                Me.txtPrimaryCodeNo.Text = PrimaryCodeNoBaru
    ''                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
    ''            End If
    ''        Else
    ''            'reset tampilan
    ''            Me.txtPrimaryCodeNo.Text = "-1"
    ''            Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
    ''            Me.txtPrimaryUnitsDefault.Text = "-1"
    ''            Me.TxtQtyperBatchMulti.Text = "0"
    ''        End If

    ''        'format tampilan
    ''        If Me.txtPrimaryCodeNo.Text <> "-1" Then
    ''            Me.lblUnitsMulti.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)
    ''            Me.txtPrimaryUnitsDefault.Text = Me.ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
    ''        Else
    ''            Me.lblUnitsMulti.Text = "Units"
    ''            Me.txtPrimaryUnitsDefault.Text = "-1"
    ''        End If
    ''    Catch ex As Exception
    ''        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    ''    End Try
    ''End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Try
            If Me.cboCodeNo.SelectedIndex <> -1 Then
                Dim RowView As DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
                Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow = RowView.Row
                Dim PrimaryCodeNoBaru As Double = row2.Primary_Product

                If Me.txtPrimaryCodeNo.Text <> PrimaryCodeNoBaru Then
                    Me.txtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                End If

                Me.txtPrimaryUnitsMulti.Text = row2.Primary_Unit_Multi
                Me.lblUnitsMulti.Text = row2.Unit_Code_Multi
                Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))

                Me.txtPrimaryUnitsDefault.Text = row2.Primary_Unit_Default
                Me.lblUnitsDefault.Text = row2.Unit_Code_Default
                Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping("0", Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))

                Me.TxtRatioReference.Text = Me.ObjCD.DigitGrouping(row2.Ratio_Reference, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                Me.TxtRatioDefault.Text = Me.ObjCD.DigitGrouping(row2.Ratio_Default, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))

                Dim idx As Integer = -1
                idx = Me.GetIndexRatio(Me.txtPrimaryCodeNo.Text)
                If Me.txtPrimaryUnitsDefault.Text <> Me.txtPrimaryUnitsMulti.Text Then
                    If idx <> -1 Then
                        Me.LlbQty.Text = "Qty (" & Me.dgvItem.Item(Me.ColRatioDesc.Index, idx).Value & ")"
                        Me.TxtRatioReference.Text = Me.dgvItem.Item(Me.ColRatioReference.Index, idx).Value
                        Me.TxtRatioDefault.Text = Me.dgvItem.Item(Me.ColRatioDefault.Index, idx).Value
                    Else
                        Me.LlbQty.Text = "Qty (" & Me.ObjCD.DigitGrouping(row2.Ratio_Reference, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) & _
                        " " & row2.Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(row2.Ratio_Default, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)) & _
                        " " & row2.Unit_Code_Default & ")"
                    End If
                    Me.TxtQtyperBatchDefault.Visible = True
                    Me.lblUnitsDefault.Visible = True
                Else
                    Me.LlbQty.Text = "Qty"
                    Me.TxtQtyperBatchDefault.Visible = False
                    Me.lblUnitsDefault.Visible = False
                End If

            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = ""

                Me.txtPrimaryUnitsMulti.Text = "-1"
                Me.lblUnitsMulti.Text = "Unit"
                Me.TxtQtyperBatchMulti.Text = "0"

                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.lblUnitsDefault.Text = "Unit"
                Me.TxtQtyperBatchDefault.Text = "0"

                Me.TxtRatioReference.Text = "0"
                Me.TxtRatioDefault.Text = "0"

                Me.LlbQty.Text = "Qty"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjPOC.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjPOC.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtPOCNo.Text = Me.ObjPOC.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtPOCNo.Text = Me.ObjPOC.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtPOCNo.Text = Me.ObjPOC.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Try
            If Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then

                ''@rosadinnn
                ''@MUaT
                Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                Me.SetQtyDefault()

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Function GetIndexRatio(ByVal vPrimaryProduct As Double) As Integer
        Try
            For Each rowDgv As DataGridViewRow In Me.dgvItem.Rows
                With rowDgv
                    If .Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        .Cells(Me.ColRatioDesc.Index).Value <> "-" Then
                        Return rowDgv.Index
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            If CDec(Me.TxtQtyperBatchMulti.Text) > 0 And Me.TxtRatioDefault.Text <> "" And Me.TxtRatioReference.Text <> "" Then
                Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.GotFocus
        TxtQtyperBatchDefault.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtQtyperBatchDefault_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Leave
        Try
            Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(TxtQtyperBatchDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
            SetQtyMulti()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyMulti()
        Try
            Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchDefault.Text) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text), ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
        TxtQtyperBatchDefault.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub SetQtyDefaultOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, vRowIdx).Value) * _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value)), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal JmlDGV As Decimal, ByVal row As DataGridViewRow)
        Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri As Decimal
        Try
            With row
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                        e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                Me.SetQtyMultiOnDgv(e.RowIndex)

                JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(CDbl(Me.txtPrimaryPO.Text), .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlObjLama = Me.ObjPOCItem.GetJmlPOCAktif_Multi(Me.txtPrimaryPO.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, Me.TxtPrimaryPOC.Text)
                JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value
                JmlPOBBaru = JmlPOB + JmlObjLama - JmlSendiri
                If JmlPOBBaru < 0 Then
                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB + JmlObjLama), ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)) + _
                                        " " + Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value + ".", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                Me.SetGboItem()

                Me.SetQtyMultiOnDgv(e.RowIndex)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyMultiOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColQtyMulti.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIdx).Value) * _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value)), _
                                                            ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value, _
                                                            Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStopProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStopProcess.Click
        Try
            If MsgBox("Stop cancellation process?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.vStopCancellation = True
                Me.NonAktifkanDetail()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_POC_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal vTblPrimaryPO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core)

        Try
            Me.baru = False
            Me.ObjPOC.ResetObject(Me.dgvItem)

            Me.ObjPOC.CariDetail(vPrimaryTransaksi, Me.dgvItem)
            Me.ObjPOC.primaryPO = ObjCutOff_Core.GetPrimaryUpdate(Me.ObjPOC.primaryPO, vTblPrimaryPO)
            Me.UserCode = Me.ObjPOC.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_POC_Save()

        Try
            Me.baru = True
            Me.ObjPOC.SimpanBaru(Me.dgvItem, Me.UserCode, False, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class