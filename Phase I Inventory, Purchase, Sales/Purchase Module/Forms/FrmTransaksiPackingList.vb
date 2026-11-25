Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiPackingList
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public StrictEditting As Boolean = True

    Public FormNo As Integer = 0

    Dim baru As Boolean
    Dim ObjPackingList As New ClsTransaksiPackingList
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjSupplier As New clsMasterSupplier
    Dim ObjPackingListItem As New ClsTransaksiPackingListItem
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjPOItem As New ClsTransaksiPOItem

    Dim IsInventoryInstalled As Boolean = False

    'Dim TBL As New DseTransaksiPO.orderbalance_report_view0DataTable
    'Dim TBL2 As New DseTransaksiPackingList.tbltrans_packinglist_item_view3DataTable
    'Dim TBL3 As New DseTransaksiPO.orderbalance_report_view0DataTable
    'Dim TBL4 As New DseTransaksiPackingList.tbltrans_packinglist_item_view3DataTable

    ''@rosadinnn
    ''@MUaT
    Dim TBL As New DseTransaksiPO.orderbalance_report_view0MultiDataTable
    Dim TBL2 As New DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiDataTable
    Dim TBL3 As New DseTransaksiPO.orderbalance_report_view0MultiDataTable
    Dim TBL4 As New DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiDataTable

    Dim PersenPajakAcuan, PrimaryCurrencySO As Double
    Dim SamaPO As Boolean = False
    Dim SamaCodeNo As Boolean = False
    Dim buffer As String = ""
    Dim bufferProduct As String = ""
    Dim BufferPO As String = ""

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryPackingList.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryPackingList.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryPackingList.BackColor, False)

            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryPackingList.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryPackingList.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryPackingList.BackColor, False)
            Me.ChkIgnoreTolerance.Enabled = True

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
                Me.btnFindSupplier.Enabled = False
            End If

            If Me.txtPrimarySupplier.Text <> "-1" And Me.txtPrimarySupplier.Text <> "" Then
                Me.SetGboItem()
            End If

            ''@rosadinnn
            '@MUaT
            Me.FormatDGV()

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
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ChkIgnoreTolerance.Enabled = False

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

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)

            Me.TxtPrimaryPackingList.Text = "-1"
            Me.txtPrimarySupplier.Text = "-1"
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
            Me.txtPrimaryPO.Text = "-1"
            Me.TxtQtyperBatchMulti.Text = "0"
            Me.TxtOrderBalance.Text = "0"
            Me.lblUnitsMulti.Text = "Units"
            Me.LblUnits2.Text = Me.lblUnitsMulti.Text
            Me.cboCodeNo.DataSource = Nothing
            Me.cboPurchaseOrder.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '1. kode product atau no po tidak diisikan
            If Me.txtPrimaryCodeNo.Text = "-1" Or Me.txtPrimaryCodeNo.Text = "" _
            Or Me.txtPrimaryPO.Text = "" Or Me.txtPrimaryPO.Text = "-1" Then
                Return False
            End If

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

            ''@rosadinnn
            ''@MUaT
            JmlPOB += ObjPO.GetToleranceUp(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text)

            JmlObjLama = Me.ObjPackingListItem.GetJmlPackingListAktif_Multi(Me.txtPrimaryPO.Text, _
            Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryPackingList.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV - CDec(Me.TxtQtyperBatchMulti.Text)
            If JmlPOBBaru < 0 Then

                ''@rosadinnn
                ''@MUaT
                MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB + JmlObjLama - JmlDGV), ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) + " " + _
                lblUnitsMulti.Text + ".", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '10. Dalam 1 Packing List bisa ada beberapa PO dengan syarat: primary currency dan pajaknya sama
            If Me.dgvItem.RowCount > 0 Then
                Dim mTaxNew, mprimaryCurrency As Double
                Me.ObjPO.GetPrimaryCurrency_Tax(Me.txtPrimaryPO.Text, mTaxNew, mprimaryCurrency)

                '   a. % pajak yang digunakan sama
                If mTaxNew <> Me.PersenPajakAcuan Then
                    MsgBox("This purchase order item cannot be in one packing list with the existing items. " + _
                    "For several purchase order item written in one packing list their currency and tax must be identical", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '   b. primary currency sama
                If mprimaryCurrency <> Me.PrimaryCurrencySO Then
                    MsgBox("This purchase order item cannot be in one packing list with the existing items. " + _
                    "For several purchase order item written in one packing list their currency and tax must be identical", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'cek bila purchase order belum diotorisasi maka tidak bisa masuk
            If Me.ObjPO.GetAuthorizationLevel(Me.txtPrimaryPO.Text) <> Me.ObjPO.GetAuthorizationLevelNeeded(Me.txtPrimaryPO.Text) Then
                MsgBox("This purchase order require more authorization, please check authorization needed.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml(ByVal vPrimaryCodeNo As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDec(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDec(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDec(row.Cells(Me.ColQtyMulti.Index).Value)
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
            KoleksiNamaData.Add("packing list number")
            KoleksiTextBox.Add(Me.TxtPackingListNo)
            KoleksiNamaData.Add("supplier code")
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

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.ObjPackingList.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjPackingList.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtPackingListNo.Text = Me.ObjPackingList.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjPackingList.TanggalTransaksi
                        Me.TxtPackingListNo.Text = Me.ObjPackingList.NoTransaksi
                        Return False
                    Else
                        Me.TxtPackingListNo.Text = Me.ObjPackingList.GetNoBaru(Me.dtpDate.Value)
                    End If
                End If
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjPackingList.TimeStamp <> Me.ObjPackingList.GetTimestamp(Me.ObjPackingList.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjPackingList.CariDetail(Me.ObjPackingList.Primary, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'Update SMESS 2026 'cek berkaitan nomor sudah digunakan 
            If Me.baru Then
                If Me.ObjPackingList.ApakahSudahAda(Me.TxtPackingListNo.Text) Then
                    Me.TxtPackingListNo.Text = Me.ObjPackingList.GetNoBaru(Me.dtpDate.Value)
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
            With Me.ObjPackingList
                .NoTransaksi = Me.TxtPackingListNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimarySupplier = Me.txtPrimarySupplier.Text
                .SupplierGITNumber = Me.TxtSupplier_GITNo.Text
                .EstimatedArrivalDate = Me.dtpETD.Value
                .ReceivedDate = Me.dtpReceivedDate.Value
                .Remarks = Me.TxtRemark.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjPackingList
                Me.TxtPackingListNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtPrimarySupplier.Text = .PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.TxtSupplier_GITNo.Text = .SupplierGITNumber
                Me.dtpETD.Value = .EstimatedArrivalDate
                Me.dtpReceivedDate.Value = .ReceivedDate
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryPackingList.Text = .Primary

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()

            'set nilai acuan untuk pengecekan SO
            If Me.dgvItem.Rows.Count <> 0 Then
                Dim PrimarySO As Double = Me.dgvItem.Item(Me.ColPrimaryPO.Index, 0).Value
                Me.ObjPO.GetPrimaryCurrency_Tax(PrimarySO, Me.PersenPajakAcuan, Me.PrimaryCurrencySO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjLS As New clsMasterListedSoftware

        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ChkIgnoreTolerance.Text = "Ignore Lower Tolerance"
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
                        If Me.ObjPackingList.SimpanBaru(Me.dgvItem, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPackingList.SimpanPerubahan(Me.dgvItem, Me.UserCode) Then
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
                        Me.KosongkanItem()
                    Else
                        Me.ObjPackingList.CariDetail(Me.TxtPrimaryPackingList.Text, Me.dgvItem)
                        Me.SalinKeField()
                        Me.KosongkanItem()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.KosongkanItem()
            Me.dtpDate.Focus()
            Me.ObjPackingList.ResetObject(Me.dgvItem)
            Me.TxtPackingListNo.Text = Me.ObjPackingList.GetNoBaru(Me.dtpDate.Value)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim i As New FrmCariTransaksiPackingList

        Try
            i.BackColorAll = Me.BackColorAll
            i.ShowDialog()

            If Not i.PrimaryTransaksi = -1 Then
                If Me.ObjPackingList.CariDetail(i.PrimaryTransaksi, Me.dgvItem) = True Then
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
            If Me.TxtPackingListNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPackingListNo.Text <> "" Then
                If Me.ObjPackingList.ApakahBolehEdit(Me.TxtPrimaryPackingList.Text, Me.StrictEditting, Me.UserCode, True) = True Then
                    Me.baru = False
                    Me.ObjPackingList.CariDetail(Me.TxtPrimaryPackingList.Text, Me.dgvItem)
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
            If Me.TxtPackingListNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPackingListNo.Text <> "" Then
                If Me.ObjPackingList.ApakahBolehHapus(Me.TxtPrimaryPackingList.Text, _
                Me.StrictEditting, Me.UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjPackingList.Hapus(Me.TxtPrimaryPackingList.Text, Me.UserCode) = True Then
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
                            If Me.ObjPackingList.SimpanBaru(Me.dgvItem, Me.UserCode) = True Then
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
                            If Me.ObjPackingList.SimpanPerubahan(Me.dgvItem, Me.UserCode) = True Then
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

    Private Sub IsiCboPurchaseOrder(ByVal vPrimarySupplier As Double)
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data awal
            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()

            'isikan dari basis data-persedian terakhir

            ''@rosadinnn
            ''@MUaT
            If Me.ChkIgnoreTolerance.Checked = False Then
                Me.ObjPOItem.IsikanTBLPO_Multi_Tolerance(vPrimarySupplier, Me.TBL, Me.cboPurchaseOrder)
            Else
                Me.ObjPOItem.IsikanTBLPO_Multi(vPrimarySupplier, Me.TBL, Me.cboPurchaseOrder)
            End If

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiRow
            Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow
            Dim ketemu As Boolean

            If Me.baru = False Then
                Me.ObjPackingListItem.IsiTBLPO_Multi(Me.TxtPrimaryPackingList.Text, Me.TBL2)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL2.Rows.Count > 0
                    row1 = Me.TBL2.Rows(0)
                    For Each row2 In Me.TBL.Rows
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
                        'Me.TBL.Rows.Add(row1.Primary_PO, row1.Primary_Supplier, row1.Primary_Product, _
                        'row1.Product_Code, row1.PO_Number, row1.Total_Qty)

                        ''@rosadinnn
                        ''@MUaT
                        row2 = TBL.Neworderbalance_report_view0MultiRow
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
                            .Ratio_Reference = row1.Ratio_Reference
                            .Ratio_Default = row1.Ratio_Default
                            .Primary_Unit_Default = row1.Primary_Unit_Default
                            .Unit_Code_Default = row1.Unit_Code_Default
                            .Tolerance_Down = 0
                            .Tolerance_Up = 0
                        End With
                        Me.TBL.Rows.Add(row2)

                    End If
                    Me.TBL2.Rows.Remove(row1)
                End While
            End If

            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLSO()
            While Me.TBL2.Rows.Count > 0
                row1 = Me.TBL2.Rows(0)
                For Each row2 In Me.TBL.Rows
                    If row1.Primary_Supplier = row2.Primary_Supplier _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_PO = row1.Primary_PO Then
                        row2.OrderBalance = row2.OrderBalance - row1.Total_Qty

                        ''@rosadinnn
                        ''@MUaT
                        ''pindahkan fungsi hapus
                        row2.OrderBalance_Multi = row2.OrderBalance_Multi - row1.Total_Qty_Multi
                        'If row2.OrderBalance_Multi = 0 Then
                        '    Me.TBL.Rows.Remove(row2)
                        'End If

                        'If row2.OrderBalance = 0 Then
                        '    Me.TBL.Rows.Remove(row2)
                        'End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL2.Rows.Remove(row1)
            End While

            ''@rosadinnn
            ''@MUaT
            For i As Integer = TBL.Rows.Count - 1 To 0 Step -1
                row2 = TBL.Rows(i)
                If Me.ChkIgnoreTolerance.Checked = True Then
                    If row2.OrderBalance_Multi + row2.Tolerance_Up <= 0 Then
                        Me.TBL.Rows.Remove(row2)
                    End If
                Else
                    If row2.OrderBalance_Multi - row2.Tolerance_Down <= 0 Then
                        Me.TBL.Rows.Remove(row2)
                    End If
                End If
            Next

            'pastikan hanya 1 so disebutkan hanya 1x
            Dim idx As Integer = 0
            Dim idx2 As Integer = 0
            Dim dummyRow As DseTransaksiPO.orderbalance_report_view0MultiRow
            While idx < Me.TBL.Rows.Count
                row2 = Me.TBL(idx)
                idx2 = idx + 1
                While idx2 <= Me.TBL.Rows.Count - 1
                    dummyRow = Me.TBL.Rows(idx2)
                    If dummyRow.PO_Number = row2.PO_Number Then
                        Me.TBL.Rows.Remove(dummyRow)
                    Else
                        idx2 += 1
                    End If
                End While

                'pindah ke baris selanjutnya
                idx += 1
            End While

            Me.TBL.DefaultView.Sort = "[PO Number] ASC"
            Me.cboPurchaseOrder.DataSource = Me.TBL.DefaultView
            Me.cboPurchaseOrder.ValueMember = "PO Number"
            Me.cboPurchaseOrder.DisplayMember = "PO Number"
            Me.cboPurchaseOrder.Refresh()

            If Me.TBL.Rows.Count = 0 Then
                Me.TxtOrderBalance.Text = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub IsikanTBLSO()
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiRow
            For Each row In Me.dgvItem.Rows
                Dim ketemu As Boolean = False
                For Each row1 In Me.TBL2.Rows
                    If row1.Primary_Supplier = Me.txtPrimarySupplier.Text And _
                    row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                    row1.Primary_PO = row.Cells(Me.ColPrimaryPO.Index).Value Then
                        'kalo ketemu ubah data order balance
                        row1.Total_Qty = row1.Total_Qty + CDec(row.Cells(Me.ColTotalQty.Index).Value)

                        ''@rosadinnn
                        ''@MUaT
                        row1.Total_Qty_Multi = row1.Total_Qty_Multi + (CDec(row.Cells(Me.ColQtyMulti.Index).Value))

                        ketemu = True
                        Exit For
                    End If
                Next
                'kalo ga ketemu tambahkan baris baru
                If ketemu = False Then
                    'Me.TBL2.Rows.Add(Me.TxtPrimaryPackingList.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'row.Cells(Me.ColPrimaryPO.Index).Value, row.Cells(Me.ColTotalQty.Index).Value, Me.txtPrimarySupplier.Text, _
                    'row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColPONo.Index).Value)

                    ''@rosadinnn
                    ''@MUaT
                    row1 = Me.TBL2.Newtbltrans_packinglist_item_view3MultiRow
                    With row1
                        .Primary_Packing_List = Me.TxtPrimaryPackingList.Text
                        .Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                        .Primary_PO = row.Cells(Me.ColPrimaryPO.Index).Value
                        .Total_Qty = row.Cells(Me.ColTotalQty.Index).Value
                        .Primary_Supplier = Me.txtPrimarySupplier.Text
                        .Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                        .PO_Number = row.Cells(Me.ColPONo.Index).Value
                        .Total_Qty_Multi = row.Cells(Me.ColQtyMulti.Index).Value
                        .Unit_Code_Multi = row.Cells(Me.ColUnitCodeMulti.Index).Value
                        .Primary_Unit_Multi = row.Cells(Me.ColPrimaryUnitMulti.Index).Value
                        .Ratio_Reference = row.Cells(Me.ColRatioReference.Index).Value
                        .Ratio_Default = row.Cells(Me.ColRatioDefault.Index).Value
                        .Primary_Unit_Default = row.Cells(Me.ColPrimaryUnits.Index).Value
                        .Unit_Code_Default = row.Cells(Me.ColUnit.Index).Value
                    End With
                    Me.TBL2.Rows.Add(row1)

                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboProduct(ByVal vPrimaryPO As Double)
        Try
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan dari basis data-persedian terakhir
            Me.cboCodeNo.DataSource = Nothing
            Me.ObjPOItem.IsikanTBLProduct_Multi(Me.txtPrimaryPO.Text, Me.TBL3, Me.cboCodeNo, Me.ChkIgnoreTolerance.Checked)

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiRow
            Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjPackingListItem.IsiTBLProductByPO_Multi(Me.TxtPrimaryPackingList.Text, _
                Me.TBL4, Me.txtPrimaryPO.Text)
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
                        'Me.TBL3.Rows.Add(row1.Primary_PO, row1.Primary_Supplier, row1.Primary_Product, _
                        'row1.Product_Code, row1.PO_Number, row1.Total_Qty)

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
                            .Ratio_Reference = row1.Ratio_Reference
                            .Ratio_Default = row1.Ratio_Default
                            .Primary_Unit_Default = row1.Primary_Unit_Default
                            .Unit_Code_Default = row1.Unit_Code_Default
                            .Tolerance_Down = 0
                            .Tolerance_Up = 0
                        End With
                        Me.TBL3.Rows.Add(row2)

                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsiTBLProduct(Me.txtPrimaryPO.Text)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                For Each row2 In Me.TBL3.Rows
                    If row1.Primary_Supplier = row2.Primary_Supplier _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_PO = row1.Primary_PO Then
                        row2.OrderBalance = row2.OrderBalance - row1.Total_Qty
                        'If row2.OrderBalance = 0 Then
                        '    Me.TBL3.Rows.Remove(row2)
                        'End If

                        ''@rosadinnn
                        ''@MUaT
                        ''pindahkan fungsi hapus
                        row2.OrderBalance_Multi = row2.OrderBalance_Multi - row1.Total_Qty_Multi
                        'If row2.OrderBalance_Multi = 0 Then
                        '    Me.TBL3.Rows.Remove(row2)
                        'End If

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

            ''@rosadinnn
            ''@MUaT
            For i As Integer = TBL3.Rows.Count - 1 To 0 Step -1
                row2 = TBL3.Rows(i)
                If Me.ChkIgnoreTolerance.Checked = True Then
                    If row2.OrderBalance_Multi + row2.Tolerance_Up <= 0 Then
                        Me.TBL3.Rows.Remove(row2)
                    End If
                Else
                    If row2.OrderBalance_Multi - row2.Tolerance_Down <= 0 Then
                        Me.TBL3.Rows.Remove(row2)
                    End If
                End If
            Next

            Me.TBL3.DefaultView.Sort = "[Product Code] ASC"
            Me.cboCodeNo.DataSource = Me.TBL3.DefaultView
            Me.cboCodeNo.DisplayMember = "Product Code"
            Me.cboCodeNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub IsiTBLProduct(ByVal vPrimaryPO As Double)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiRow
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Primary_Supplier = Me.txtPrimarySupplier.Text And _
                                row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                                row1.Primary_PO = row.Cells(Me.ColPrimaryPO.Index).Value Then
                            'kalo ketemu ubah data order balance
                            row1.Total_Qty = row1.Total_Qty + CDec(row.Cells(Me.ColTotalQty.Index).Value)

                            ''@rosadinnn
                            ''@MUaT
                            row1.Total_Qty_Multi = row1.Total_Qty_Multi + CDec(row.Cells(Me.ColQtyMulti.Index).Value)

                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        'Me.TBL4.Rows.Add(Me.TxtPrimaryPackingList.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        'row.Cells(Me.ColPrimaryPO.Index).Value, row.Cells(Me.ColTotalQty.Index).Value, _
                        'Me.txtPrimarySupplier.Text, row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColPONo.Index).Value)

                        ''@rosadinnn
                        ''@MUaT
                        row1 = Me.TBL4.Newtbltrans_packinglist_item_view3MultiRow
                        With row1
                            .Primary_Packing_List = Me.TxtPrimaryPackingList.Text
                            .Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                            .Primary_PO = row.Cells(Me.ColPrimaryPO.Index).Value
                            .Total_Qty = row.Cells(Me.ColTotalQty.Index).Value
                            .Primary_Supplier = Me.txtPrimarySupplier.Text
                            .Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                            .PO_Number = row.Cells(Me.ColPONo.Index).Value
                            .Total_Qty_Multi = row.Cells(Me.ColQtyMulti.Index).Value
                            .Primary_Unit_Multi = row.Cells(Me.ColPrimaryUnitMulti.Index).Value
                            .Unit_Code_Multi = row.Cells(Me.ColUnitCodeMulti.Index).Value
                            .Ratio_Reference = row.Cells(Me.ColRatioReference.Index).Value
                            .Ratio_Default = row.Cells(Me.ColRatioDefault.Index).Value
                            .Primary_Unit_Default = row.Cells(Me.ColPrimaryUnits.Index).Value
                            .Unit_Code_Default = row.Cells(Me.ColUnit.Index).Value
                        End With
                        Me.TBL4.Rows.Add(row1)

                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItem()
        Try
            Me.cboPurchaseOrder.DataSource = Nothing
            Me.cboCodeNo.DataSource = Nothing

            'sales order & product code
            Me.IsiCboPurchaseOrder(Me.txtPrimarySupplier.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click

        ''@rosadinnn
        ''@MUaT
        Dim rowDgv As DataGridViewRow

        Try
            If Me.ApakahDataItemBenar = True Then
                Me.txtSupplierCode.Enabled = False
                Me.btnFindSupplier.Enabled = False

                Dim Index As Integer = Me.GetIndex(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)

                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    'row.Cells(Me.ColTotalQty.Index).Value = CDec(row.Cells(Me.ColTotalQty.Index).Value) + CDec(Me.TxtQtyperBatchMulti.Text)

                    ''@rosadinnn
                    ''@MUaT
                    row.Cells(Me.ColQtyMulti.Index).Value = CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatchMulti.Text)
                    Me.SetQtyDefaultOnDgv(row.Index)

                Else
                    'Me.dgvItem.Rows.Add("", Me.cboPurchaseOrder.Text, Me.cboCodeNo.Text, Me.txtDesc.Text, "", _
                    '"", 0, 0, "", 0, Me.TxtQtyperBatchMulti.Text, Me.lblUnitsMulti.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsDefault.Text, -1, Me.txtPrimaryPO.Text)

                    ''@rosadinnn
                    ''@MUaT
                    Me.dgvItem.Rows.Add()
                    rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDgv
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPONo.Index).Value = Me.cboPurchaseOrder.Text
                        .Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColSuppCodeNo.Index).Value = ""
                        .Cells(Me.ColSuppDesc.Index).Value = ""
                        .Cells(Me.ColQtyMaster.Index).Value = "0"
                        .Cells(Me.ColQtyPerMaster.Index).Value = "0"
                        .Cells(Me.ColMasterUnit.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = "0"
                        .Cells(Me.ColTotalQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = "-1"
                        .Cells(Me.ColPrimaryPO.Index).Value = Me.txtPrimaryPO.Text
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
                        Me.ObjPO.GetPrimaryCurrency_Tax(Me.txtPrimaryPO.Text, Me.PersenPajakAcuan, Me.PrimaryCurrencySO)
                    End If
                End If

                'kendalikan focus item selanjutnya
                Dim JmlObjLama, JmlPOBBaru, JmlDGV, JmlPOB As Decimal
                JmlDGV = Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)

                ''@rosadinnn
                ''@MUaT
                JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text)

                JmlObjLama = Me.ObjPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV
                Me.TxtQtyperBatchMulti.Text = "0"
                Me.TxtRemarkItem.Text = ""

                'If JmlPOBBaru <= 0 Then
                '    'remove row, pada code no ybs
                '    Dim rowView As Data.DataRowView = Me.cboCodeNo.SelectedItem
                '    Dim row As DseTransaksiPO.orderbalance_report_view0MultiRow = rowView.Row
                '    Me.TBL3.Rows.Remove(row)
                '    Me.cboCodeNo_SelectedValueChanged(sender, e)
                '    Me.cboCodeNo.Focus()
                '    If Me.cboCodeNo.Items.Count = 0 Then
                '        Me.SetGboItem()
                '        Me.cboPurchaseOrder.Focus()
                '    End If
                'Else
                '    If Me.cboCodeNo.SelectedIndex <> Me.cboCodeNo.Items.Count - 1 Then
                '        Me.cboCodeNo.SelectedIndex += 1
                '        Me.cboCodeNo.Focus()
                '    Else
                '        'focus pada cbo code no
                '        Me.cboCodeNo.Focus()
                '        Me.SetOrderBalance()
                '    End If
                'End If
                Me.SetGboItem()

                'format dgv
                Me.FormatDGV()

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
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

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
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

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
        Dim row As DataGridViewRow

        Dim BufferQtyDefault As String = ""
        Dim BufferRatioReference As String = ""
        Dim BufferRatioDefault As String = ""
        Dim BufferRatioDesc As String = ""
        Dim BufferLblQty As String = ""


        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                If Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value Or _
                    Me.TxtRemark.Enabled = False Then
                    Exit Sub
                End If

                BufferQtyDefault = Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value
                BufferRatioReference = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                BufferRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                BufferRatioDesc = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                BufferLblQty = Me.LlbQty.Text

                frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)
                frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                frmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value
                frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value

                If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                    Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value & _
                                                                            " " & Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value & " = " & _
                                                                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value & " " & _
                                                                            Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    If Me.txtPrimaryPO.Text = Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.RowIndex).Value.ToString And Me.txtPrimaryCodeNo.Text = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value.ToString Then
                        Me.TxtRatioReference.Text = frmRatio.TxtRatioMulti.Text
                        Me.TxtRatioDefault.Text = frmRatio.TxtRatioDefault.Text
                        Me.LlbQty.Text = "Qty (" & Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value & ")"
                    End If
                    Me.SetQtyDefaultOnDgv(e.RowIndex)

                    If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) = 0 Then
                        MsgBox("Ratio value restored to previous value, because ratio you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = BufferRatioDesc
                        Me.LlbQty.Text = BufferLblQty
                        Me.SetQtyDefaultOnDgv(e.RowIndex)
                    End If

                End If
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

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                        e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value))

                'cek apakah sales order balance menjadi minus setelah perubahan data
                Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri As Decimal
                JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlObjLama = Me.ObjPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPO.Index).Value)

                ''@rosadinnn
                ''@MUaT
                JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value
                JmlPOB += ObjPO.GetToleranceUp(.Cells(ColPrimaryPO.Index).Value, .Cells(ColPrimaryCodeNo.Index).Value)

                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlPOBBaru = JmlPOB - JmlSendiri + JmlObjLama
                If JmlPOBBaru < 0 Then
                    MsgBox("Maximum value you can enter is " + CStr(JmlPOB + JmlObjLama) + _
                    " " + CStr(Me.ObjProduct.GetUnitCode(.Cells(Me.ColPrimaryCodeNo.Index).Value)) + ".", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'jika jumlah menurun maka harus load ulang data
                Me.SetGboItem()

                ''@rosadinnn
                ''@MUaT
                Me.SetQtyDefaultOnDgv(e.RowIndex)

                If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) = 0 Then
                    MsgBox("Qty restored to previous value, because qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(e.RowIndex)
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditQty_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal JmlDGV As Decimal, ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                        e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                'cek apakah sales order balance menjadi minus setelah perubahan data
                Me.SetQtyMultiOnDgv(e.RowIndex)
                Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri As Decimal
                JmlPOB = Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlObjLama = Me.ObjPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPO.Index).Value)
                JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value

                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlPOBBaru = JmlPOB - JmlSendiri + JmlObjLama
                If JmlPOBBaru < 0 Then
                    MsgBox("Maximum value you can enter is " + CStr(JmlPOB + JmlObjLama) + " " + CStr(Me.ObjProduct.GetUnitCode(.Cells(Me.ColPrimaryCodeNo.Index).Value)) + ".", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                Me.SetGboItem()
                Me.SetQtyMultiOnDgv(e.RowIndex)
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
                    And row.Cells(Me.ColPrimaryPO.Index).Value = row1.Cells(Me.ColPrimaryPO.Index).Value Then
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
                    And row.Cells(Me.ColPrimaryPO.Index).Value = row1.Cells(Me.ColPrimaryPO.Index).Value Then
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
            row.Cells(Me.ColTotalQty.Index).Value = CDec(row.Cells(Me.ColTotalQty.Index).Value) _
            + CDec(row1.Cells(Me.ColTotalQty.Index).Value)
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
                Me.btnFindSupplier.Enabled = True
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
                    Else
                        Me.SetOrderBalance()
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
            If Me.txtPrimaryPO.Text = CStr(Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.Row.Index).Value) Then
                Me.SamaPO = True
            Else
                Me.SamaPO = False
            End If
            If Me.txtPrimaryCodeNo.Text = CStr(Me.dgvItem(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) Then
                Me.SamaCodeNo = True
            Else
                Me.SamaCodeNo = False
            End If
            Me.BufferPO = CStr(Me.dgvItem.Item(Me.ColPrimaryPO.Index, e.Row.Index).Value)
            Me.bufferProduct = CStr(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        If Me.txtSupplierCode.Enabled = False Or Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                    Me.txtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                    Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                    Me.TxtSupplier_GITNo.Focus()
                    Me.SetGboItem()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplier_GITNo.Focus()
                        Me.SetGboItem()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
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

    Private Sub txtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        If Me.txtSupplierCode.Enabled = False Or Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) Then
                    Dim PrimaryBaru As Double = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                    If PrimaryBaru <> Me.txtPrimarySupplier.Text Then
                        Me.txtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                        Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplier_GITNo.Focus()
                        Me.SetGboItem()
                    End If
                Else
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.TxtSupplier_GITNo.Focus()
                        Me.SetGboItem()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.TxtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
                        Me.KosongkanItem()
                    End If
                End If
            Else
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.TxtSupplier_GITNo.Focus()
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.txtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = i.PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.TxtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.TxtSupplier_GITNo.Focus()
                Me.SetGboItem()
            Else
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Focus()
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub cboPurchaseOrder_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPurchaseOrder.SelectedValueChanged
        Try
            If Me.cboPurchaseOrder.SelectedIndex <> -1 Then
                Dim RowView As Data.DataRowView = Me.cboPurchaseOrder.Items(Me.cboPurchaseOrder.SelectedIndex)
                Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow = RowView.Row
                Dim PrimaryPOBaru As Double = row2.Primary_PO
                If Me.txtPrimaryPO.Text = "" Then
                    Me.txtPrimaryPO.Text = PrimaryPOBaru
                    Me.IsiCboProduct(Me.txtPrimaryPO.Text)
                Else
                    If Me.txtPrimaryPO.Text <> PrimaryPOBaru Then
                        Me.txtPrimaryPO.Text = PrimaryPOBaru
                        Me.IsiCboProduct(Me.txtPrimaryPO.Text)
                    End If
                End If
            Else
                Me.txtPrimaryPO.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Try
            If Me.cboCodeNo.SelectedIndex <> -1 Then
                Dim RowView As Data.DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
                Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow = RowView.Row
                Dim PrimaryCodeNoBaru As Double = row2.Primary_Product

                If Me.txtPrimaryCodeNo.Text = "" Then
                    Me.txtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.txtPrimaryUnitsDefault.Text = Me.ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
                Else
                    If Me.txtPrimaryCodeNo.Text <> PrimaryCodeNoBaru Then
                        Me.txtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.txtPrimaryUnitsDefault.Text = Me.ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
                    End If
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
                idx = Me.GetIndexRatio(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text)
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

                Me.LblUnits2.Text = Me.lblUnitsMulti.Text

                Me.SetOrderBalance()
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = ""

                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.lblUnitsMulti.Text = "Unit"
                Me.TxtQtyperBatchMulti.Text = "0"

                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.lblUnitsDefault.Text = "Unit"
                Me.TxtQtyperBatchDefault.Text = "0"

                Me.TxtRatioReference.Text = "0"
                Me.TxtRatioDefault.Text = "0"

                Me.LlbQty.Text = "Qty"
                Me.TxtOrderBalance.Text = "0"
            End If

            Me.LblUnits2.Text = Me.lblUnitsMulti.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjPackingList.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjPackingList.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtPackingListNo.Text = Me.ObjPackingList.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtPackingListNo.Text = Me.ObjPackingList.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtPackingListNo.Text = Me.ObjPackingList.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        txtSupplierCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        txtSupplierCode.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplier_GITNo.GotFocus
        TxtSupplier_GITNo.BackColor = Me.BackColorAll
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSupplier_GITNo.LostFocus
        TxtSupplier_GITNo.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        TxtRemark.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        TxtRemark.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.GotFocus
        TxtQtyperBatchMulti.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        TxtQtyperBatchMulti.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        TxtRemarkItem.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        TxtRemarkItem.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmCetakTransaksiPackingList

        Try
            If Me.TxtPackingListNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPackingListNo.Text <> "" Then
                Me.ObjPackingList.Cetak(ij.crv1, Me.TxtPrimaryPackingList.Text)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub SetOrderBalance()
        Try
            If IsNumeric(Me.txtPrimaryCodeNo.Text) = True And IsNumeric(Me.txtPrimaryPO.Text) = True Then
                Me.TxtOrderBalance.Text = _
                Me.ObjCD.DigitGrouping(Me.ObjPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text) _
                + Me.ObjPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) - _
                Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")

            ''@rosadinnn
            ''@MUaT
        Finally
            Me.SetFieldToleranceUp()

        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Try
            If Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then
                'Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))

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
    Private Function GetIndexRatio(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Integer
        Try
            For Each rowDgv As DataGridViewRow In Me.dgvItem.Rows
                With rowDgv
                    If .Cells(Me.ColPrimaryPO.Index).Value = vPrimaryPO And .Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
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

    Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
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


    ''@rosadinnn
    ''@MUaT
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

    Private Sub SetQtyDefault()
        Try
            Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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

    Private Sub LlbQty_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbQty.Click
        Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
        Try
            If Me.TxtRemark.Enabled = False Or Me.LlbQty.Text = "Qty" Then
                Exit Sub
            End If
            If Me.GetIndexRatio(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text) <> -1 Then
                MsgBox("There is already ratio defined below. You cannot set different ratio for exist product.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            frmRatio.vRatioMulti = Me.TxtRatioReference.Text
            frmRatio.vUnitMulti = Me.lblUnitsMulti.Text
            frmRatio.vRatioDefault = Me.TxtRatioDefault.Text
            frmRatio.vUnitDefault = Me.lblUnitsDefault.Text
            If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtRatioReference.Text = frmRatio.TxtRatioMulti.Text
                Me.TxtRatioDefault.Text = frmRatio.TxtRatioDefault.Text
                Me.LlbQty.Text = "Qty (" & frmRatio.TxtRatioMulti.Text & " " & frmRatio.LlbUnitMulti.Text & " = " & frmRatio.TxtRatioDefault.Text & " " & frmRatio.LlbUnitDefault.Text & ")"
                SetQtyDefault()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
        TxtQtyperBatchDefault.BackColor = Me.txtPrimaryCodeNo.BackColor
    End Sub

    Private Sub SetFieldToleranceUp()
        Try
            Me.ChkIgnoreTolerance.Enabled = True
            Me.TxtToleranceUp.Text = ObjCD.DigitGrouping(Me.ObjPO.GetToleranceUp(Me.txtPrimaryPO.Text, Me.txtPrimaryCodeNo.Text), _
            ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkIgnoreTolerance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkIgnoreTolerance.CheckedChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.SetGboItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtOrderBalance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtOrderBalance.TextChanged
        Try
            If IsNumeric(Me.TxtOrderBalance.Text) = True Then
                If CDec(Me.TxtOrderBalance.Text) = 0 Then
                    Me.TxtToleranceUp.Text = "0"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_PackingList_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal vTblPrimary As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core)
        Dim TblTempPO As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.ObjPackingList.ResetObject(Me.dgvItem)
            Me.ObjPackingList.CariDetail(vPrimaryTransaksi, Me.dgvItem)

            'Update Primary PO
            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryPO.Index).Value) = False And rowD.Cells(Me.ColPrimaryPO.Index).Value <> "-1" Then
                    TblTempPO.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value)
                    If TblTempPO.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), vTblPrimary)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), PrimaryUpdate, TblTempPO)
                        rowD.Cells(Me.ColPrimaryPO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryPO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPO.Index).Value), TblTempPO)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_PackingList_Save() As Double
        Try
            Me.baru = True
            Me.ObjPackingList.SimpanBaru1(Me.dgvItem, Me.UserCode, False, False, True)
            Return Me.ObjPackingList.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function
End Class