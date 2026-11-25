Option Explicit On
Option Strict Off

Imports Core_Module
Imports Inventory_Module

Public Class FrmTransaksiGRN
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0

    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjSupplier As New clsMasterSupplier
    Dim ObjGRNItem As New ClsTransaksiGRNItem
    Dim ObjGRN As New ClsTransaksiGRN
    Dim ObjStockInOut As New clsTransaksiStockInOut
    Dim ObjDO As New ClsTransaksiDO
    Dim TBL As New Inventory_Module.dseTransaksiStockInOut.tbltransaksistockinout_view1DataTable
    Dim Baru As Boolean
    Dim buffer As String = ""

    Private Sub TxtStockInNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.GotFocus
        Me.TxtStockInNo.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtStockInNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.LostFocus
        Me.TxtStockInNo.BackColor = Me.txtPrimaryStockIn.BackColor
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.txtPrimaryStockIn.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            '      Me.chkTaxReturn.Enabled = True

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.dgvItem.AllowUserToDeleteRows = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.txtPrimaryStockIn.Text = "-1"
            Me.TxtPrimaryGRN.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("GRN  number")
            KoleksiTextBox.Add(Me.TxtGRNNo)
            KoleksiNamaData.Add("stock in number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
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
            If Me.Baru = False Then
                If Me.ObjGRN.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjGRN.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtStockInNo.Text = Me.ObjGRN.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjGRN.TanggalTransaksi
                        Me.TxtStockInNo.Text = Me.ObjGRN.NoTransaksi
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
            With Me.ObjGRN
                .NoTransaksi = Me.TxtGRNNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .primaryGIT = Me.txtPrimaryGIT.Text
                .primaryStockInPurchase = Me.txtPrimaryStockIn.Text
                .primarySupplier = Me.txtPrimarySupplier.Text
                .Remark = Me.TxtRemark.Text
                .SupplierDONumber = Me.txtSupplierDONumber.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDataStockIn()
        Try
            If Me.txtPrimaryStockIn.Text <> "" And Me.txtPrimaryStockIn.Text <> "-1" Then
                Me.TxtStockInNo.Text = Me.ObjStockInOut.GetNoTransaksi(Me.txtPrimaryStockIn.Text)
                '  Dim PrimaryDO As Double = Me.ObjStockInOut.GetPrimaryDO(Me.txtPrimaryStockIn.Text)
                '    Me.txtPrimaryDO.Text = PrimaryDO
                Me.ObjStockInOut.GetInfo(Me.txtPrimaryStockIn.Text, _
                Me.txtSupplierDONumber.Text, Me.txtPrimarySupplier.Text, Me.txtSupplierCode.Text, Me.TxtSupplierName.Text, _
                Me.txtPrimaryGIT.Text, Me.txtPackingListNumber.Text)
            Else
                Me.TxtStockInNo.Text = ""
                Me.txtSupplierDONumber.Text = ""
                Me.txtSupplierCode.Clear()
                Me.TxtSupplierName.Clear()
                Me.txtPackingListNumber.Clear()
                '           Me.dtpDate.Value = #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjGRN
                Me.TxtGRNNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remark
                Me.TxtPrimaryGRN.Text = .primary
                Me.txtPrimaryStockIn.Text = .primaryStockInPurchase
                Me.txtSupplierDONumber.Text = .SupplierDONumber
                Me.txtPrimarySupplier.Text = .primarySupplier

                Me.IsiDataStockIn()

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
        Try
            Me.Baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
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
                If Baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjGRN.SimpanBaru(Me.dgvItem, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
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
                        If Me.ObjGRN.SimpanPerubahan(Me.dgvItem, Me.UserCode) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
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
                Me.btnFind.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjGRN.CariDetail(Me.TxtPrimaryGRN.Text, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Me.Baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.ObjGRN.ResetObject(Me.dgvItem)
            Me.TxtGRNNo.Text = Me.ObjGRN.GetNoBaru(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiGRN
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()

            If Not ij.PrimaryGRN = -1 Then
                If Me.ObjGRN.CariDetail(ij.PrimaryGRN, Me.dgvItem) = True Then
                    Me.SalinKeField()
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

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtGRNNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtGRNNo.Text <> "" Then
                If Me.ObjGRN.ApakahBolehHapus(Me.TxtPrimaryGRN.Text, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjGRN.Hapus(Me.TxtPrimaryGRN.Text, Me.UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnFind.Focus()
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
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If Baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjGRN.SimpanBaru(Me.dgvItem, Me.UserCode) = True Then
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
                            If Me.ObjGRN.SimpanPerubahan(Me.dgvItem, Me.UserCode) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
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

    Private Sub FormatDGV()
        Dim row As DataGridViewRow

        Try
            Dim AdaSec As Boolean = False
            Dim AdaPO As Boolean = False
            Dim AdaExpiry As Boolean = False
            Dim AdaBatchNo As Boolean = False

            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
                row.Cells(Me.ColQtyMaster.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyMaster.Index).Value)
                row.Cells(Me.ColSecQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecQty.Index).Value)
                row.Cells(Me.ColQtyPerMaster.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyPerMaster.Index).Value)
                row.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColTotalQty.Index).Value)
                If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
                    AdaSec = True
                End If
                If row.Cells(Me.colPrimaryPO.Index).Value <> -1 Then
                    AdaPO = True
                End If

                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = True Then
                    If row.Cells(Me.ColExpiryDate.Index).Value <> #1/1/1950# Then
                        AdaExpiry = True
                        row.Cells(Me.ColExpiryDate.Index).Value = Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy")
                    Else
                        row.Cells(Me.ColExpiryDate.Index).Value = "-"
                    End If
                End If
                If row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    AdaBatchNo = True
                End If
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                    row.Cells(Me.ColSecQty.Index).Value = "-"
                    row.Cells(Me.ColSecUnit.Index).Value = "-"
                Else
                    If row.Cells(Me.ColSecQty.Index).Value = "-1" Then
                        row.Cells(Me.ColSecQty.Index).Value = "-"
                        row.Cells(Me.ColSecUnit.Index).Value = "-"
                    End If
                End If
            Next

            If AdaSec = True Then
                Me.ColSecQty.Visible = True
                Me.ColSecUnit.Visible = True
            Else
                Me.ColSecQty.Visible = False
                Me.ColSecUnit.Visible = False
            End If

            If AdaPO = True Then
                Me.ColPONO.Visible = True
            Else
                Me.ColPONO.Visible = False
            End If
            If AdaExpiry = True Then
                Me.ColExpiryDate.Visible = True
            Else
                Me.ColExpiryDate.Visible = False
            End If
            If AdaBatchNo = True Then
                Me.ColBatchNo.Visible = True
            Else
                Me.ColBatchNo.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '2. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = 14 Then
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
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
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
                Me.TxtStockInNo.Enabled = True
                Me.btnFindStockIn.Enabled = True
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
            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)

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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub TxtStockOutNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.GotFocus
        TxtStockInNo.BackColor = BackColorAll
    End Sub

    Private Sub TxtStockOutNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.LostFocus
        TxtStockInNo.BackColor = Me.TxtPrimaryGRN.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        TxtRemark.BackColor = BackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        TxtRemark.BackColor = Me.TxtPrimaryGRN.BackColor
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.GotFocus
        TxtQtyperBatch.BackColor = BackColorAll
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.LostFocus
        TxtQtyperBatch.BackColor = Me.TxtPrimaryGRN.BackColor
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        TxtRemarkItem.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        TxtRemarkItem.BackColor = Me.TxtPrimaryGRN.BackColor
    End Sub

    'Khusus untuk CMSI, bila sudah memiliki invoice maka mata uang (yang digunakan untuk perhitungan total
    'dan sub total), exchange rate to IDR, dan tax rate digunakan data dari invoice dari mana SR ybs berasal.
    'Lebih jauh lagi bila buat nota retur pajak, maka 

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '        Try
        ' Dim obj As New ClsTransaksiSalesReturn
        ' obj.CetakSalesReturnArtha(FrmCetakTransaksiSalesReturn.crv1, Me.TxtPrimaryGRN.Text)
        ' obj = Nothing
        ' FrmCetakTransaksiSalesReturn.ShowDialog()
        ' Catch ex As Exception
        ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        ' End Try
    End Sub

    Private Sub IsikanDGV()
        Dim row As Inventory_Module.dseTransaksiStockInOut.tbltransaksistockinout_view1Row
        Try
            ObjStockInOut.IsikanTBLStockInOut(TBL, Me.txtPrimaryStockIn.Text)
            Me.dgvItem.Rows.Clear()
            For Each row In TBL.Rows
                Dim ExpiryDate As String = Format(row.Expiry_Date, "dd MMM yyyy")
                Me.dgvItem.Rows.Add("0", row.Product_Code, row.Product_Description, "-1", "-1", _
                row.No_PO, row.Batch_No, ExpiryDate, "0", "0", "-1", "0", row.Qty, _
                row.Unit_Code_Qty, row.Secondary_Qty, row.Unit_Code_Sec_Qty, row.Remarks, _
                row.Primary_Product, "-1", row.Primary_Units, row.Primary_Secondary_Units, _
                row.Primary_PO_Purchase)
            Next
            ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, _
    ByVal LotNo As String, ByVal SecQty As Double) As Integer
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvItem.Rows
                If CDbl(row.Cells(Me.ColSecQty.Index).Value) = SecQty And row.Cells(Me.ColPrimaryCodeNo.Index).Value _
                = vPrimaryProduct And row.Cells(Me.colPrimaryPO.Index).Value = vPrimarySO And _
                LotNo = row.Cells(Me.ColBatchNo.Index).Value Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtStockInNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtStockInNo.KeyDown
        Dim i As New FrmCariTransaksiStockInOut
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Try
            If e.KeyValue = 13 Then
                If Me.ObjStockInOut.ApakahSudahAda(Me.TxtStockInNo.Text) = True Then
                    Me.txtPrimaryStockIn.Text = Me.ObjStockInOut.GetPrimary(Me.TxtStockInNo.Text)
                    'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                    Dim NoTransaksiLain As String = Me.ObjGRN.GetNoGRNByPrimaryStockInOut(Me.txtPrimaryStockIn.Text)
                    If NoTransaksiLain <> "" Then
                        MsgBox("This stock in sales return referred by sales return " _
                        + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                        Me.TxtStockInNo.Clear()
                        Me.txtPrimaryStockIn.Text = "-1"
                        Me.IsiDataStockIn()
                        Exit Sub
                    End If
                    Me.IsiDataStockIn()
                    Me.IsikanDGV()
                    '                    Me.TxtReasonOfReturn.Focus()
                Else
                    'ada 2 stock in dengan kode yang sama atau tidak ada stock in yang dimaksud
                    If Me.TxtStockInNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtStockInNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksiStockInOut > 0 Then
                        'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                        Dim NoTransaksiLain As String = Me.ObjGRN.GetNoGRNByPrimaryStockInOut(i.PrimaryTransaksiStockInOut)
                        If NoTransaksiLain <> "" Then
                            MsgBox("This stock in sales return referred by sales return " _
                            + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                            Me.TxtStockInNo.Clear()
                            Me.txtPrimaryStockIn.Text = "-1"
                            Me.IsiDataStockIn()
                            Exit Sub
                        End If

                        Me.txtPrimaryStockIn.Text = i.PrimaryTransaksiStockInOut
                        Me.IsiDataStockIn()
                        Me.IsikanDGV()
                        '     Me.TxtReasonOfReturn.Focus()
                    Else
                        Me.TxtStockInNo.Clear()
                        Me.txtPrimaryStockIn.Text = "-1"
                        Me.IsiDataStockIn()
                        Me.dgvItem.Rows.Clear()
                        Me.TxtStockInNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtStockInNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStockInNo.Leave
        Dim i As New FrmCariTransaksiStockInOut
        i.BackColorAll = Me.BackColorAll
        i.JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Try
            If Me.TxtStockInNo.Text <> "" Then
                If Me.ObjStockInOut.ApakahSudahAda(Me.TxtStockInNo.Text) = True Then
                    Me.txtPrimaryStockIn.Text = Me.ObjStockInOut.GetPrimary(Me.TxtStockInNo.Text)
                    Me.IsiDataStockIn()
                    Me.IsikanDGV()
                    '       Me.TxtReasonOfReturn.Focus()
                Else
                    If Me.TxtStockInNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtStockInNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksiStockInOut > 0 Then
                        'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                        Dim NoTransaksiLain As String = Me.ObjGRN.GetNoGRNByPrimaryStockInOut(i.PrimaryTransaksiStockInOut)
                        If NoTransaksiLain <> "" And NoTransaksiLain <> Me.TxtGRNNo.Text Then
                            MsgBox("This stock in sales return referred by sales return " _
                            + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                            Me.TxtStockInNo.Clear()
                            Me.txtPrimaryStockIn.Text = "-1"
                            Me.IsiDataStockIn()
                            Exit Sub
                        End If

                        Me.txtPrimaryStockIn.Text = i.PrimaryTransaksiStockInOut
                        Me.IsiDataStockIn()
                        Me.IsikanDGV()
                        '               Me.TxtReasonOfReturn.Focus()
                    Else
                        Me.TxtStockInNo.Clear()
                        Me.txtPrimaryStockIn.Text = "-1"
                        Me.IsiDataStockIn()
                        Me.TxtStockInNo.Focus()
                        Me.dgvItem.Rows.Clear()
                    End If
                End If
            Else
                Me.txtPrimaryStockIn.Text = "-1"
                Me.IsiDataStockIn()
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindStockIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindStockIn.Click
        Dim i As New FrmCariTransaksiStockInOut
        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll
            i.JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            If Me.TxtStockInNo.Text <> "" Then
                i.txtCari.Text = Me.TxtStockInNo.Text
            End If
            i.ShowDialog()

            If i.PrimaryTransaksiStockInOut > 0 Then
                'cek apakah sudah ada sales return yang merujuk ke stock in sales return ini
                Dim NoTransaksiLain As String = Me.ObjGRN.GetNoGRNByPrimaryStockInOut(i.PrimaryTransaksiStockInOut)
                If NoTransaksiLain <> "" Then
                    MsgBox("This GRN referred by GRN: " _
                    + NoTransaksiLain, MsgBoxStyle.Information, "Information")
                    Me.TxtStockInNo.Clear()
                    Me.txtPrimaryStockIn.Text = "-1"
                    Me.IsiDataStockIn()
                    Exit Sub
                End If

                Me.txtPrimaryStockIn.Text = i.PrimaryTransaksiStockInOut
                Me.IsiDataStockIn()
                Me.IsikanDGV()
                '         Me.TxtReasonOfReturn.Focus()
            Else
                Me.TxtStockInNo.Clear()
                Me.txtPrimaryStockIn.Text = "-1"
                '                Me.IsiDataStockIn()
                Me.TxtStockInNo.Focus()
                Me.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.Baru = False Then
                    If Me.ObjGRN.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjGRN.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtGRNNo.Text = Me.ObjGRN.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtGRNNo.Text = Me.ObjGRN.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtGRNNo.Text = Me.ObjGRN.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class