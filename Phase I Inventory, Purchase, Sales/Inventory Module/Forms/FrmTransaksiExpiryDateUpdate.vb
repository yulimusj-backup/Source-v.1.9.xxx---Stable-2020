Option Explicit On
Option Strict Off

Imports Core_Module

''' <summary>
''' form ini digunakan untuk melakukan update data expiry date.
''' Data expiry date yang diupdate adalah di:
''' '1. Data di tbl product qty
''' '2. Data disemua item transaksi inventory
''' '3. Data disemua transaksi sales/purchase yang menggunakan batch no ybs
''' Hanya berlaku untuk product yang kontrol batch no dan kontrol expiry date
''' </summary>
''' <remarks></remarks>
Public Class FrmTransaksiExpiryDateUpdate
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjProduct As New clsMasterProduct
    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                If MsgBox("Exit without save?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                    Me.Close()
                End If
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiExpiryDateUpdate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DTA As New DseTransaksiBatchNumberUpdateTableAdapters.tblmaster_productTableAdapter

        Try
            'cek terlebih dahulu ke database apakah ada minimal 1 code no yang 
            'control batch no dan expiry date secara bersamaan
            If DTA.ScalarQuery_GetJmlBatchNumberControlled.HasValue = True Then
                If DTA.ScalarQuery_GetJmlBatchNumberControlled.Value > 0 Then
                    Exit Sub
                End If
            End If

            MsgBox("There are no product with batch number or expiry date controlled. You cannot use this feature.", MsgBoxStyle.Information, "Information")
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
       Handles TxtCodeNo.KeyDown
        Dim objCode As New clsMasterProduct

        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeNo.Text <> "" Then
                    If objCode.ApakahSudahAda(Me.TxtCodeNo.Text) = True Then
                        Me.TxtPrimaryCodeNo.Text = objCode.GetPrimary(Me.TxtCodeNo.Text)
                        Me.TxtCodeNo.Text = objCode.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.cboBatchNo.Focus()
                        Me.IsiCBOCodeNo()
                    Else
                        'ada 2 GroupCode dengan kode yang sama atau tidak ada GroupCode yang dimaksud
                        Me.btnFindCodeNo.PerformClick()
                    End If
                Else
                    Me.btnFindCodeNo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objCode = Nothing
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct
        Dim objCode As New clsMasterProduct
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCodeNo.Text <> "" Then
                If objCode.ApakahSudahAda(Me.TxtCodeNo.Text) = True Then
                    Me.TxtPrimaryCodeNo.Text = objCode.GetPrimary(Me.TxtCodeNo.Text)
                    Me.TxtCodeNo.Text = objCode.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = objCode.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.TxtBatchNew.Clear()
                        Me.dtpNew.Value = Today
                        Me.dtpOld.Value = Today
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtBatchNew.Clear()
                Me.dtpNew.Value = Today
                Me.dtpOld.Value = Today
            End If
            Me.IsiCBOCodeNo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
  Handles btnFindCodeNo.Click
        Dim i As New frmCariMasterProduct
        Dim objCode As New clsMasterProduct
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = objCode.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.cboBatchNo.Focus()
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtBatchNew.Clear()
                Me.dtpNew.Value = Today
                Me.dtpOld.Value = Today
                Me.TxtCodeNo.Focus()
            End If
            Me.IsiCBOCodeNo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub IsiCBOCodeNo()
        Try
            Me.cboBatchNo.SelectedIndex = -1

            If Me.TxtPrimaryCodeNo.Text = "-1" Then
                Me.cboBatchNo.DataSource = Nothing
                Me.cboBatchNo.Items.Clear()
                Me.chkExpiry.Checked = True
            Else
                'isikah chk expiry date kalo controlled expiry
                Me.chkExpiry.Checked = Me.ObjProduct.ApakahExpiryDateControlled(Me.TxtPrimaryCodeNo.Text) = True
                If Me.chkExpiry.Checked = True Then
                    Me.dtpNew.Enabled = True
                Else
                    Me.dtpNew.Enabled = False
                End If

                'isikan data batch number
                Me.ObjProduct.IsiCboBatchNo_UntukUpdateBatch(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo, Me.TBL)

                Me.TBL.DefaultView.Sort = "[Lot No] ASC"
                Me.cboBatchNo.DataSource = Me.TBL.DefaultView
                Me.cboBatchNo.ValueMember = "Lot No"
                Me.cboBatchNo.DisplayMember = "Lot No"
                Me.cboBatchNo.Refresh()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim DTA1 As New DseTransaksiBatchNumberUpdateTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim dummy As Integer = 0
        Dim rowView As Data.DataRowView
        Dim row As dseMasterProduct.tblmaster_productqty1Row

        Try
            Me.TxtBatchNew.Text = Me.TxtBatchNew.Text.Trim

            If Me.TxtBatchNew.Text = "" Then
                MsgBox("Please specify new batch number.", MsgBoxStyle.Information, "Information")
                Me.TxtBatchNew.Focus()
                Exit Sub
            End If

            'pastikan batch number, tidak diubah menjadi batch number yang sudah ada
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                If i <> Me.cboBatchNo.SelectedIndex Then
                    rowView = TBL.DefaultView(i)
                    row = rowView.Row
                    If row.Lot_No.ToLower = Me.TxtBatchNew.Text.ToLower Then
                        MsgBox("Duplication in batch number. Please change your new batch number.", MsgBoxStyle.Information, "Information")
                        Me.TxtBatchNew.Focus()
                        Exit Sub
                    End If
                End If
            Next

            'jika menjadi batch number yang sudah, maka menjadi penggabungan batch, proses ini irreversible, sehingga user harus diberikan 
            'warning yang keras --> nanti pada SMESS aja

            If MsgBox("Update data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'update batch number, jika batch numbernya berbeda
                If Me.cboBatchNo.Text <> Me.TxtBatchNew.Text Then
                    '1. Inventory
                    '   - Transaksi Stock In Out
                    DTA1.UpdateQuery_UpdateBatch(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Transaksi Bundling In
                    DTA1.UpdateQuery_UpdateBatch_BundlingIn(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Transaksi Bundling Out
                    DTA1.UpdateQuery_UpdateBatch_BundlingOut(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Transaksi Stock Transfer
                    DTA1.UpdateQuery_UpdateBatchStockTransfer(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Master Product Qty
                    DTA1.UpdateQuery_UpdateBatchProductQty(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Transaksi Booking Inv Item
                    DTA1.UpdateQuery_UpdateBatchBookingInv(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    'untuk yang stock taking dipikirkan ulang, karena mereka tidak bisa dikembalikan lagi, jadi sebenarnya tidak
                    'perlu di lakukan update pada bagian ini
                    '   - Transaksi Stock Taking Tag Item
                    '   - Transaksi Stock Taking Item Buffer
                    '   - Transaksi Stock Taking Item Start
                    '   - Transkasi Stock Taking Item Penyesuaian

                    '2. Sales
                    '   - Transaksi DO
                    DTA1.UpdateQuery_UpdateBatchDO(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Sales Invoice Item
                    DTA1.UpdateQuery_UpdateBatchInvoice(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Sales Invoice Obselete Item
                    DTA1.UpdateQuery_UpdateBatchInvoiceObselete(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Transaksi Sales Return
                    DTA1.UpdateQuery_UpdateBatchSalesReturn(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '3. Purchase
                    '   - Transaksi GRN Item
                    DTA1.UpdateQuery_UpdateBatchGRN(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    '   - Transaksi Purchase Return
                    DTA1.UpdateQuery_UpdateBatchPurchaseReturn(Me.TxtBatchNew.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                End If

                'update expiry date jika, menggunakan kontrol expiry date dan expiry date baru yang dimasukkan berbeda
                If Me.dtpNew.Value.Date <> Me.dtpOld.Value.Date And Me.chkExpiry.Checked = True Then
                    '1. Inventory
                    '   - Transaksi Stock In Out
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryStockInOut(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryStockInOut(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Transaksi Bundling In
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryBundlingIn(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryBundlingIn(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Transaksi Bundling Out
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryBundlingOut(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryBundlingOut(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Transaksi Stock Transfer
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryStockTransfer(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryStockTransfer(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Master Product Qty
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryProductQty(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryProductQty(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Transaksi Booking Inv Item
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryBookingInv(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryBookingInv(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Transaksi Stock Taking Tag Item
                    '   - Transaksi Stock Taking Item Buffer
                    '   - Transaksi Stock Taking Item Start
                    '   - Transkasi Stock Taking Item Penyesuaian

                    '2. Sales
                    '   - Transaksi DO
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryDO(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryDO(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Sales Invoice Item
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryInvoice(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryInvoice(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Sales Invoice Obselete Item
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiry_ObseleteInvoice(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiry_ObseleteInvoice(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Transaksi Sales Return
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpirySalesReturnItem(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpirySalesReturnItem(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '3. Purchase
                    '   - Transaksi GRN Item
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiryGRN(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiryGRN(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                    '   - Transaksi Purchase Return
                    If Me.TxtBatchNew.Text <> Me.cboBatchNo.Text Then
                        DTA1.UpdateQuery_UpdateExpiry_PurchaseReturn(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.TxtBatchNew.Text)
                    Else
                        DTA1.UpdateQuery_UpdateExpiry_PurchaseReturn(Me.dtpNew.Value, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                    End If
                End If

                MsgBox("Data changed", MsgBoxStyle.Information, "Information")
                dummy = Me.cboBatchNo.SelectedIndex

                Me.TxtCodeNo.Text = ""
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.TxtBatchNew.Text = ""
                Me.dtpNew.Value = Today
                Me.dtpOld.Value = Today

                Me.IsiCBOCodeNo()
                If dummy < Me.cboBatchNo.Items.Count Then
                    Me.cboBatchNo.SelectedIndex = dummy
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
        End Try
    End Sub

    Private Sub cboBatchNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboBatchNo.SelectedIndexChanged
        Dim row As dseMasterProduct.tblmaster_productqty1Row

        Try
            If Me.cboBatchNo.SelectedIndex <> -1 Then
                If Me.chkExpiry.Checked = True Then
                    If Me.TBL.Rows.Count > 0 Then
                        row = Me.TBL.DefaultView(Me.cboBatchNo.SelectedIndex).Row
                        Me.dtpOld.Value = row.Expiry_Date
                    Else
                        Me.dtpOld.Value = Today
                    End If
                Else
                    Me.dtpOld.Value = Today
                End If
                Me.dtpNew.Value = Me.dtpOld.Value
                Me.TxtBatchNew.Text = Me.cboBatchNo.Text
            Else
                Me.dtpOld.Value = Today
                Me.dtpNew.Value = Me.dtpOld.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class