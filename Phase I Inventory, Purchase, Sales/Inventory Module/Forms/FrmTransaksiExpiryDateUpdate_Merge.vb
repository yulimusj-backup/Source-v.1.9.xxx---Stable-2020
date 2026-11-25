Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiExpiryDateUpdate_Merge

    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjProduct As New clsMasterProduct
    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.tblmaster_productqty1DataTable

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                If MsgBox("Exit without save?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
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
            If DTA.ScalarQuery_GetJmlBatchNumberControlled.HasValue = True Then
                If DTA.ScalarQuery_GetJmlBatchNumberControlled.Value > 0 Then
                    Exit Sub
                End If
            End If
            MsgBox("There are no product with batch number. You cannot use this feature.", MsgBoxStyle.Information, "Information")
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            Me.TxtCodeNo.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
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

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
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
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
            End If
            Me.IsiCBOCodeNo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            objCode = Nothing
        End Try
    End Sub

    Private Sub btnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
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
            Me.cboBatchNoDestination.SelectedIndex = -1
            If Me.TxtPrimaryCodeNo.Text = "-1" Then
                Me.cboBatchNo.DataSource = Nothing
                Me.cboBatchNo.Items.Clear()
                Me.cboBatchNoDestination.DataSource = Nothing
                Me.cboBatchNoDestination.Items.Clear()
            Else
                'isikan data batch number
                Me.ObjProduct.IsiCboBatchNo_UntukUpdateBatch(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo, Me.TBL)
                Me.ObjProduct.IsiCboBatchNo_UntukUpdateBatch(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNoDestination, Me.TBL2)

                Me.TBL.DefaultView.Sort = "[Lot No] ASC"
                Me.cboBatchNo.DataSource = Me.TBL.DefaultView
                Me.cboBatchNo.ValueMember = "Lot No"
                Me.cboBatchNo.DisplayMember = "Lot No"
                Me.cboBatchNo.Refresh()

                Me.cboBatchNoDestination.DataSource = Me.TBL2.DefaultView
                Me.cboBatchNoDestination.ValueMember = "Lot No"
                Me.cboBatchNoDestination.DisplayMember = "Lot No"
                Me.cboBatchNoDestination.Refresh()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            Me.TxtCodeNo.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim dta As New DseTransaksiBatchNumberUpdateTableAdapters.tblmaster_productqtyTableAdapter
        Dim nol As New Nullable(Of Long)
        Try
            If Not Me.IsDataValid() Then
                Return
            End If
            If MsgBox("Are you sure to process?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                dta.Connection.ConnectionString = Core_Module.GetConstring
                nol = dta.ScalarQuery_GetLoginCount(Me.UserCode)
                If nol.HasValue Then
                    If nol.Value > 0 Then
                        If MsgBox("There are users that login to SMESS except you, do you want to continue?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                            Return
                        End If
                    End If
                End If
                MergeBatch()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function IsDataValid() As Boolean
        Try
            If Me.TxtPrimaryCodeNo.Text = "-1" Then
                MsgBox("Select product code first.", MsgBoxStyle.Information, "Information")
                Me.TxtPrimaryCodeNo.Focus()
                Return False
            End If
            If Me.cboBatchNo.Items.Count = 0 Then
                MsgBox("There is no batch to merge for this product code.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.cboBatchNo.Text = Me.cboBatchNoDestination.Text Then
                MsgBox("Select different destination batch number.", MsgBoxStyle.Information, "Information")
                Me.cboBatchNoDestination.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Sub MergeBatch()
        Dim DTA1 As New DseTransaksiBatchNumberUpdateTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim dtaProductQty As New DseTransaksiBatchNumberUpdateTableAdapters.tblmaster_productqtyTableAdapter
        Dim tblProductQtySource As New DseTransaksiBatchNumberUpdate.tblmaster_productqtyDataTable
        Dim tblProductQtyDestination As New DseTransaksiBatchNumberUpdate.tblmaster_productqtyDataTable
        Dim rowProductQtySource As DseTransaksiBatchNumberUpdate.tblmaster_productqtyRow
        Dim rowProductQtyDestination As DseTransaksiBatchNumberUpdate.tblmaster_productqtyRow
        Try
            Me.Cursor = Cursors.WaitCursor
            ''update product qty
            dtaProductQty.Connection.ConnectionString = Core_Module.GetConstring
            dtaProductQty.Fill(tblProductQtySource, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            dtaProductQty.Fill(tblProductQtyDestination, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNoDestination.Text)
            For Each rowProductQtySource In tblProductQtySource
                tblProductQtyDestination.DefaultView.RowFilter = "[Primary Storage Area] = " & rowProductQtySource.Primary_Storage_Area.ToString
                If tblProductQtyDestination.DefaultView.Count > 0 Then
                    rowProductQtyDestination = tblProductQtyDestination.DefaultView(0).Row
                    dtaProductQty.UpdateQuery_Qty(rowProductQtySource.Qty + rowProductQtyDestination.Qty, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNoDestination.Text, rowProductQtySource.Primary_Storage_Area)
                    dtaProductQty.DeleteQuery(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text, rowProductQtySource.Primary_Storage_Area)
                Else
                    dtaProductQty.UpdateQuery_Batch(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, rowProductQtySource.Primary_Storage_Area, Me.cboBatchNo.Text)
                End If
            Next

            ''update others table
            '   - Transaksi Stock In Out
            DTA1.UpdateQuery_UpdateBatch(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi Bundling In
            DTA1.UpdateQuery_UpdateBatch_BundlingIn(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi Bundling Out
            DTA1.UpdateQuery_UpdateBatch_BundlingOut(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi Stock Transfer
            DTA1.UpdateQuery_UpdateBatchStockTransfer(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi Booking Inv Item
            DTA1.UpdateQuery_UpdateBatchBookingInv(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi DO
            DTA1.UpdateQuery_UpdateBatchDO(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Sales Invoice Item
            DTA1.UpdateQuery_UpdateBatchInvoice(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Sales Invoice Obselete Item
            DTA1.UpdateQuery_UpdateBatchInvoiceObselete(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi Sales Return
            DTA1.UpdateQuery_UpdateBatchSalesReturn(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi GRN Item
            DTA1.UpdateQuery_UpdateBatchGRN(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            '   - Transaksi Purchase Return
            DTA1.UpdateQuery_UpdateBatchPurchaseReturn(Me.cboBatchNoDestination.Text, Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)

            MsgBox("Data changed", MsgBoxStyle.Information, "Information")
            Me.TxtCodeNo.Text = ""
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.IsiCBOCodeNo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class