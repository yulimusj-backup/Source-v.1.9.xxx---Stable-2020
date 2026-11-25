Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmTransaksiStockInAdj_CoverMinus
    Dim Lanjut As Boolean = True
    Dim Successfull As Boolean = False

    Public TblQty As dseTransaksiStockInOut_CoverMinus.tblmaster_productqtyDataTable
    Public RefDate As Date = #1/1/1950#

    Dim DTAQty As New dseTransaksiStockInOut_CoverMinusTableAdapters.tblmaster_productqtyTableAdapter
    Dim rowQty As dseTransaksiStockInOut_CoverMinus.tblmaster_productqtyRow
    Dim DTAInOut As New dseTransaksiStockInOut_CoverMinusTableAdapters.tbltrans_stock_in_outTableAdapter
    Dim TBLInOut As New dseTransaksiStockInOut_CoverMinus.tbltrans_stock_in_outDataTable
    Dim rowInOut As dseTransaksiStockInOut_CoverMinus.tbltrans_stock_in_outRow
    Dim DTABundling_After As New dseTransaksiStockInOut_CoverMinusTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
    Dim TBLBundling_After As New dseTransaksiStockInOut_CoverMinus.tbltransaksi_stockbundling_item_afterDataTable
    Dim rowBundling_After As dseTransaksiStockInOut_CoverMinus.tbltransaksi_stockbundling_item_afterRow
    Dim DTABundling_Before As New dseTransaksiStockInOut_CoverMinusTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
    Dim TBLBundling_Before As New dseTransaksiStockInOut_CoverMinus.tbltransaksi_stockbundling_item_beforeDataTable
    Dim rowBundling_Before As dseTransaksiStockInOut_CoverMinus.tbltransaksi_stockbundling_item_beforeRow
    Dim DTATransfer As New dseTransaksiStockInOut_CoverMinusTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
    Dim TBLTransfer As New dseTransaksiStockInOut_CoverMinus.tbltransaksi_stocktransfer_itemDataTable
    Dim rowTransfer As dseTransaksiStockInOut_CoverMinus.tbltransaksi_stocktransfer_itemRow
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTransaksiStockInAdj_CoverMinus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'kosongkan tbl qty
            Me.TblQty.Rows.Clear()
            Me.TBLBundling_After.Rows.Clear()
            Me.TBLBundling_Before.Rows.Clear()
            Me.TBLInOut.Rows.Clear()
            Me.TBLTransfer.Rows.Clear()

            Me.DTABundling_After.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTABundling_Before.Connection.ConnectionString = Me.DTABundling_After.Connection.ConnectionString
            Me.DTAInOut.Connection.ConnectionString = Me.DTABundling_After.Connection.ConnectionString
            Me.DTAQty.Connection.ConnectionString = Me.DTABundling_After.Connection.ConnectionString
            Me.DTATransfer.Connection.ConnectionString = Me.DTABundling_After.Connection.ConnectionString

            'isikan tabel qty, stock in out, bundling after, bundling before, transfer
            Me.DTABundling_After.Fill(Me.TBLBundling_After, Me.RefDate)
            Me.DTABundling_Before.Fill(Me.TBLBundling_Before, Me.RefDate)
            Me.DTAInOut.Fill(Me.TBLInOut, Me.RefDate)
            Me.DTAQty.Fill(Me.TblQty)
            Me.DTATransfer.Fill(Me.TBLTransfer, Me.RefDate)
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim TipeStock As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust

        Try
            Me.ObjCD.UpdateThreading()

            'step 1, update menggunakan stock in out
            Me.prb1.Value = 0
            Me.prb1.Maximum = Me.TBLInOut.Rows.Count
            For Each Me.rowInOut In Me.TBLInOut
                If Me.Lanjut = True Then
                    Me.prb1.Value += 1
                    Me.LblProgres.Text = "Step 1 of 4: reversing stock in out - data " + Me.prb1.Value.ToString + " of " + Me.prb1.Maximum.ToString

                    With Me.rowInOut
                        TipeStock = .Tipe_Transaksi
                        If TipeStock = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                        TipeStock = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                        TipeStock = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                        TipeStock = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                        TipeStock = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                            Me.TambahQty(.Primary_Product, .Qty, .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                        Else
                            Me.KurangiQty(.Primary_Product, .Qty, .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                        End If
                    End With
                Else
                    Me.Successfull = False
                    Exit Sub
                End If
            Next
            'step 2, update menggunakan stock transfer
            Me.prb1.Value = 0
            Me.prb1.Maximum = Me.TBLTransfer.Rows.Count
            For Each Me.rowTransfer In Me.TBLTransfer
                If Me.Lanjut = True Then
                    Me.prb1.Value += 1
                    Me.LblProgres.Text = "Step 2 of 4: reversing stock transfer - data " + Me.prb1.Value.ToString + " of " + Me.prb1.Maximum.ToString

                    With Me.rowTransfer
                        Me.TambahQty(.Primary_Product, .Qty, .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                        Me.KurangiQty(.Primary_Product, .Qty, .Batch_No, .Expiry_Date, .primary_destination_area)
                    End With
                Else
                    Me.Successfull = False
                    Exit Sub
                End If
            Next

            'step 3a, update menggunakan item stock bundling after
            Me.prb1.Value = 0
            Me.prb1.Maximum = Me.TBLBundling_After.Rows.Count + Me.TBLBundling_Before.Rows.Count
            For Each Me.rowBundling_After In Me.TBLBundling_After
                If Me.Lanjut = True Then
                    Me.prb1.Value += 1
                    Me.LblProgres.Text = "Step 3 of 4: reversing stock bundling - data " + Me.prb1.Value.ToString + " of " + Me.prb1.Maximum.ToString

                    With Me.rowBundling_After
                        Me.KurangiQty(.Primary_Product, .Qty, .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                    End With
                Else
                    Me.Successfull = False
                    Exit Sub
                End If
            Next
            'step 3b, update menggunakan item stock bundling before
            For Each Me.rowBundling_Before In Me.TBLBundling_Before
                If Me.Lanjut = True Then
                    Me.prb1.Value += 1
                    Me.LblProgres.Text = "Step 3 of 4: reversing stock bundling - data " + Me.prb1.Value.ToString + " of " + Me.prb1.Maximum.ToString

                    With Me.rowBundling_Before
                        Me.TambahQty(.Primary_Product, .Qty, .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                    End With
                Else
                    Me.Successfull = False
                    Exit Sub
                End If
            Next

            'step 4, hanya pertahankan yang minus saja
            Me.prb1.Value = 0
            Me.prb1.Maximum = Me.TblQty.Rows.Count
            For ij As Double = Me.TblQty.Rows.Count - 1 To 0 Step -1
                Me.prb1.Value += 1
                Me.LblProgres.Text = "Step 4 of 4: filtering negatives only - data " + Me.prb1.Value.ToString + " of " + Me.prb1.Maximum.ToString
                Me.rowQty = Me.TblQty.Rows(ij)
                If Me.rowQty.Qty > 0 Then
                    Me.TblQty.Rows.Remove(Me.rowQty)
                End If
            Next
            'klo sampe sini, tandai sukses
            Me.Successfull = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.Successfull = True Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahQty(ByVal vPrimaryProduct As Double, ByVal vQty As Decimal, ByVal vBatchNo As String, ByVal vExpiry As Date, ByVal vPrimaryLocation As Double)
        Try
            Me.TblQty.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString + " AND [Lot No] = '" + vBatchNo + "' AND [Primary Storage Area] = " + vPrimaryLocation.ToString
            If Me.TblQty.DefaultView.Count = 0 Then
                'tambahkan baris baru
                Me.rowQty = Me.TblQty.Newtblmaster_productqtyRow
                Me.rowQty.Expiry_Date = vExpiry
                Me.rowQty.Lot_No = vBatchNo
                Me.rowQty.Primary = -1
                Me.rowQty.Primary_Product = vPrimaryProduct
                Me.rowQty.Primary_Storage_Area = vPrimaryLocation
                Me.rowQty.Qty = vQty
                Me.rowQty.Qty_Secondary = -1
                Me.TblQty.Rows.Add(Me.rowQty)
            Else
                Me.rowQty = Me.TblQty.DefaultView(0).Row
                Me.rowQty.Qty += vQty
                If Me.rowQty.Qty = 0 Then
                    Me.TblQty.Rows.Remove(Me.rowQty)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangiQty(ByVal vPrimaryProduct As Double, ByVal vQty As Decimal, ByVal vBatchNo As String, ByVal vExpiry As Date, ByVal vPrimaryLocation As Double)
        Try
            Me.TblQty.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString + " AND [Lot No] = '" + vBatchNo + "' AND [Primary Storage Area] = " + vPrimaryLocation.ToString
            If Me.TblQty.DefaultView.Count = 0 Then
                'tambahkan baris baru
                Me.rowQty = Me.TblQty.Newtblmaster_productqtyRow
                Me.rowQty.Expiry_Date = vExpiry
                Me.rowQty.Lot_No = vBatchNo
                Me.rowQty.Primary = -1
                Me.rowQty.Primary_Product = vPrimaryProduct
                Me.rowQty.Primary_Storage_Area = vPrimaryLocation
                Me.rowQty.Qty = vQty * -1
                Me.rowQty.Qty_Secondary = -1
                Me.TblQty.Rows.Add(Me.rowQty)
            Else
                Me.rowQty = Me.TblQty.DefaultView(0).Row
                Me.rowQty.Qty -= vQty
                If Me.rowQty.Qty = 0 Then
                    Me.TblQty.Rows.Remove(Me.rowQty)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel data creation?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.Lanjut = False
                Me.btnCancel.Text = "Cancelling"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
