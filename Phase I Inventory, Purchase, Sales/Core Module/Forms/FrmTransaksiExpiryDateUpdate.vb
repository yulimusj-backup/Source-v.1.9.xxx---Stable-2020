Option Explicit On
Option Strict Off

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
    Dim objProduct As New clsMasterProduct
    Dim TBL As New DseTransaksiExpiryDateUpdate.TblFilterDataTable

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiExpiryDateUpdate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'cek terlebih dahulu ke database apakah ada minimal 1 code no yang 
            'controll batch no dan expiry date secara bersamaan

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsiCboBatchNo(ByVal VprimaryProduct As Double) As Boolean
        Try
            ''  Me.KosongkanItem()

            Me.IsiDataByStockInOutItem(VprimaryProduct)
            Me.IsiDataByMasterProductQty(VprimaryProduct)
            Me.IsiDataByTransaksiBookingGITItem(VprimaryProduct)
            Me.IsiDataByTransaksiBookingInvItem(VprimaryProduct)
            Me.IsiDataByTransaksiDOItem(VprimaryProduct)
            Me.IsiDataByTransaksiDOlossitem(VprimaryProduct)
            Me.IsiDataByTransaksiEmailItem(VprimaryProduct)
            Me.IsiDataByTransaksiGRNItem(VprimaryProduct)
            Me.IsiDataByTransaksiPurchaseReturnItem(VprimaryProduct)
            Me.IsiDataByTransaksiSalesReturnItem(VprimaryProduct)
            Me.IsiDataByTransaksiStockTakingItemPenyesuaian(VprimaryProduct)
            Me.IsiDataByTransaksiStockTakingTagItem(VprimaryProduct)
            Me.IsiDataByTransaksiStockBundlingitemAfter(VprimaryProduct)
            Me.IsiDataByTransaksiStockBundlingitemBefore(VprimaryProduct)
            Me.IsiDataByTransaksiStockTransferItem(VprimaryProduct)

            Me.cboBatchNo.DataSource = Me.TBL.DefaultView
            Me.TBL.DefaultView.Sort = "[BatchNo] ASC"
            Me.cboBatchNo.DisplayMember = "BatchNo"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function IsiDataByStockInOutItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltransaksi_stockinout_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltransaksi_stockinout_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Function

    Private Function IsiDataByMasterProductQty(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tblmaster_productqty1TableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tblmaster_productqty1DataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tblmaster_productqty1Row
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            '' Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Lot_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Function

    Private Function IsiDataByTransaksiBookingGITItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_booking_gititemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_booking_gititemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_booking_gititemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()

        End Try
    End Function

    Private Function IsiDataByTransaksiBookingInvItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_booking_invitemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_booking_invitemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_booking_invitemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function IsiDataByTransaksiDOItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_do_item1TableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_do_item1DataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_do_item1Row
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Private Function IsiDataByTransaksiDOlossitem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_do_loss_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_do_loss_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_do_loss_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function IsiDataByTransaksiEmailItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_email_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_email_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_email_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function IsiDataByTransaksiGRNItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_grn_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_grn_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_grn_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function IsiDataByTransaksiPurchaseReturnItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_purchasereturn_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_purchasereturn_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function IsiDataByTransaksiSalesReturnItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_salesreturn_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_salesreturn_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function IsiDataByTransaksiStockTakingItemPenyesuaian(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_stock_taking_itempenyesuaianTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_stock_taking_itempenyesuaianDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_stock_taking_itempenyesuaianRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Private Function IsiDataByTransaksiStockTakingTagItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_stock_taking_tag_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltrans_stock_taking_tag_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltrans_stock_taking_tag_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Private Function IsiDataByTransaksiStockBundlingitemAfter(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltransaksi_stockbundling_item_afterDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltransaksi_stockbundling_item_afterRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function IsiDataByTransaksiStockBundlingitemBefore(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltransaksi_stockbundling_item_beforeDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltransaksi_stockbundling_item_beforeRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Private Function IsiDataByTransaksiStockTransferItem(ByVal vprimaryproduct As Double) As Boolean
        Dim DTA As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Dim TBlInOut As DseTransaksiExpiryDateUpdate.tbltransaksi_stocktransfer_itemDataTable
        Dim Row As DseTransaksiExpiryDateUpdate.tbltransaksi_stocktransfer_itemRow
        Dim rowTransaksi, DummyRow As DseTransaksiExpiryDateUpdate.TblFilterRow
        Dim SudahAda As Boolean = False

        Try
            ''  Me.TBL.Rows.Clear()
            TBlInOut = DTA.GetDataBy_PrimaryProduct(vprimaryproduct)
            For Each Row In TBlInOut.Rows
                rowTransaksi = Me.TBL.NewTblFilterRow
                rowTransaksi.PrimaryProduct = Row.Primary_Product
                rowTransaksi.BatchNo = Row.Batch_No
                rowTransaksi.ExpiryDate = Row.Expiry_Date

                'cek apakah sudah ada
                SudahAda = False
                For Each DummyRow In Me.TBL.Rows
                    If DummyRow.PrimaryProduct = rowTransaksi.PrimaryProduct And DummyRow.BatchNo = rowTransaksi.BatchNo Then
                        SudahAda = True
                    End If
                Next
                If SudahAda = False Then
                    Me.TBL.Rows.Add(rowTransaksi)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub btnFindCodeNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCodeNo.Click
        Dim i As New frmCariMasterProduct
        Try
            If Me.objProduct.ApakahSudahAda(Me.txtPrimaryCode.Text) = True Then
                'cek apakah product yang bersangkutan expiry date controlled
                Me.txtPrimaryCode.Text = objProduct.GetPrimary(Me.TxtProductCode.Text)
                If Me.objProduct.ApakahExpiryDateControlled(Me.txtPrimaryCode.Text) = True Then
                    Me.TxtProductCode.Text = objProduct.GetProductCode(Me.txtPrimaryCode.Text)
                    Me.txtDesc.Text = objProduct.GetDescription(Me.txtPrimaryCode.Text)
                    Me.IsiCboBatchNo(Me.txtPrimaryCode.Text)
                    Me.cboBatchNo.Focus()
                    Me.dtpOldExpiry.Value = Me.objProduct.GetExpiryDate(Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                Else
                    MsgBox("Please Select Expiry Date Controlled Product", MsgBoxStyle.Information, "Information")
                    Me.TxtProductCode.Clear()
                    Me.txtDesc.Clear()
                    Me.txtPrimaryCode.Text = "-1"
                    Me.TxtProductCode.Focus()
                    Me.cboBatchNo.DataSource = Nothing
                End If

                ''  Me.SetDefaultPrecision()
                Me.cboBatchNo.Focus()
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.TxtProductCode.Text <> "" Then
                    i.txtCari.Text = Me.TxtProductCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryProduct > 0 Then
                    'cek apakah product yang bersangkutan expiry date controlled
                    If Me.objProduct.ApakahExpiryDateControlled(i.PrimaryProduct) = True Then
                        Me.txtPrimaryCode.Text = i.PrimaryProduct
                        Me.TxtProductCode.Text = objProduct.GetProductCode(Me.txtPrimaryCode.Text)
                        Me.txtDesc.Text = objProduct.GetDescription(Me.txtPrimaryCode.Text)
                        Me.IsiCboBatchNo(Me.txtPrimaryCode.Text)
                        ''  Me.SetDefaultPrecision()
                        Me.cboBatchNo.Focus()
                        Me.dtpOldExpiry.Value = Me.objProduct.GetExpiryDate(Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    Else
                        MsgBox("Please Select Expiry Date Controlled Product", MsgBoxStyle.Information, "Information")
                        Me.TxtProductCode.Clear()
                        Me.txtDesc.Clear()
                        Me.txtPrimaryCode.Text = "-1"
                        Me.TxtProductCode.Focus()
                        Me.cboBatchNo.DataSource = Nothing
                        ''    Me.SetDefaultPrecision()
                    End If
                Else
                    Me.TxtProductCode.Clear()
                    Me.txtDesc.Clear()
                    Me.txtPrimaryCode.Text = "-1"
                    Me.TxtProductCode.Focus()
                    Me.cboBatchNo.DataSource = Nothing
                    ''    Me.SetDefaultPrecision()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtProductCode.KeyDown
        Dim i As New frmCariMasterProduct
        Try
            If e.KeyValue = 13 Then
                If objProduct.ApakahSudahAda(Me.TxtProductCode.Text) = True Then
                    'cek apakah product yang bersangkutan expiry date controlled
                    Me.txtPrimaryCode.Text = objProduct.GetPrimary(Me.TxtProductCode.Text)
                    If Me.objProduct.ApakahExpiryDateControlled(Me.txtPrimaryCode.Text) = True Then
                        Me.TxtProductCode.Text = objProduct.GetProductCode(Me.txtPrimaryCode.Text)
                        Me.txtDesc.Text = objProduct.GetDescription(Me.txtPrimaryCode.Text)
                        Me.IsiCboBatchNo(Me.txtPrimaryCode.Text)
                        ''  Me.SetDefaultPrecision()
                        Me.cboBatchNo.Focus()
                        Me.dtpOldExpiry.Value = Me.objProduct.GetExpiryDate(Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    Else
                        MsgBox("Please Select Expiry Date Controlled Product", MsgBoxStyle.Information, "Information")
                        Me.TxtProductCode.Clear()
                        Me.txtDesc.Clear()
                        Me.txtPrimaryCode.Text = "-1"
                        Me.TxtProductCode.Focus()
                        Me.cboBatchNo.DataSource = Nothing
                        ''    Me.SetDefaultPrecision()
                    End If
                    
                Else
                    If Me.TxtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtProductCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        'cek apakah product yang bersangkutan expiry date controlled
                        If Me.objProduct.ApakahExpiryDateControlled(i.PrimaryProduct) = True Then
                            Me.txtPrimaryCode.Text = i.PrimaryProduct
                            Me.TxtProductCode.Text = objProduct.GetProductCode(Me.txtPrimaryCode.Text)
                            Me.txtDesc.Text = objProduct.GetDescription(Me.txtPrimaryCode.Text)
                            Me.IsiCboBatchNo(Me.txtPrimaryCode.Text)
                            ''  Me.SetDefaultPrecision()
                            Me.cboBatchNo.Focus()
                            Me.dtpOldExpiry.Value = Me.objProduct.GetExpiryDate(Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                        Else
                            MsgBox("Please Select Expiry Date Controlled Product", MsgBoxStyle.Information, "Information")
                            Me.TxtProductCode.Clear()
                            Me.txtDesc.Clear()
                            Me.txtPrimaryCode.Text = "-1"
                            Me.TxtProductCode.Focus()
                            Me.cboBatchNo.DataSource = Nothing
                            ''    Me.SetDefaultPrecision()
                        End If
                    Else
                        Me.TxtProductCode.Clear()
                        Me.txtDesc.Clear()
                        Me.txtPrimaryCode.Text = "-1"
                        Me.TxtProductCode.Focus()
                        Me.cboBatchNo.DataSource = Nothing
                        ''    Me.SetDefaultPrecision()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.Leave
        Dim i As New frmCariMasterProduct

        Try
            Me.KosongkanItem()
            If Me.TxtProductCode.Text <> "" Then
                If objProduct.ApakahSudahAda(Me.TxtProductCode.Text) = True Then
                    'cek apakah product yang bersangkutan expiry date controlled)
                    Me.txtPrimaryCode.Text = objProduct.GetPrimary(Me.TxtProductCode.Text)
                    If Me.objProduct.ApakahExpiryDateControlled(Me.txtPrimaryCode.Text) = True Then
                        Me.TxtProductCode.Text = objProduct.GetProductCode(Me.txtPrimaryCode.Text)
                        Me.txtDesc.Text = objProduct.GetDescription(Me.txtPrimaryCode.Text)
                        Me.IsiCboBatchNo(Me.txtPrimaryCode.Text)

                        ''  Me.SetDefaultPrecision()
                        Me.cboBatchNo.Focus()
                        Me.dtpOldExpiry.Value = Me.objProduct.GetExpiryDate(Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    Else
                        MsgBox("Please Select Expiry Date Controlled Product", MsgBoxStyle.Information, "Information")
                        Me.TxtProductCode.Clear()
                        Me.txtPrimaryCode.Text = "-1"
                        Me.TxtProductCode.Focus()
                        Me.txtDesc.Clear()
                        Me.cboBatchNo.DataSource = Nothing
                        ''  Me.SetDefaultPrecision()
                    End If
                Else
                    If Me.TxtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtProductCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        'cek apakah product yang bersangkutan expiry date controlled
                        If Me.objProduct.ApakahExpiryDateControlled(i.PrimaryProduct) = True Then
                            Me.txtPrimaryCode.Text = i.PrimaryProduct
                            Me.TxtProductCode.Text = objProduct.GetProductCode(Me.txtPrimaryCode.Text)
                            Me.txtDesc.Text = objProduct.GetDescription(Me.txtPrimaryCode.Text)
                            Me.IsiCboBatchNo(Me.txtPrimaryCode.Text)
                            ''  Me.SetDefaultPrecision()
                            Me.cboBatchNo.Focus()
                            Me.dtpOldExpiry.Value = Me.objProduct.GetExpiryDate(Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                        Else
                            MsgBox("Please Select Expiry Date Controlled Product", MsgBoxStyle.Information, "Information")
                            Me.TxtProductCode.Clear()
                            Me.txtPrimaryCode.Text = "-1"
                            Me.TxtProductCode.Focus()
                            Me.txtDesc.Clear()
                            Me.cboBatchNo.DataSource = Nothing
                            ''  Me.SetDefaultPrecision()
                        End If
                    Else
                        Me.TxtProductCode.Clear()
                        Me.txtPrimaryCode.Text = "-1"
                        Me.TxtProductCode.Focus()
                        Me.txtDesc.Clear()
                        Me.cboBatchNo.DataSource = Nothing
                        ''  Me.SetDefaultPrecision()
                    End If
                End If
            Else
                Me.TxtProductCode.Clear()
                Me.txtPrimaryCode.Text = "-1"
                Me.TxtProductCode.Focus()
                Me.txtDesc.Clear()
                Me.cboBatchNo.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Text = ""
            Me.dtpOldExpiry.Value = Now

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim dta As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stockinout_item1TableAdapter
        Dim dta1 As New DseTransaksiExpiryDateUpdateTableAdapters.tblmaster_productqtyTableAdapter
        Dim dta2 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_booking_gititem1TableAdapter
        Dim dta3 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_booking_invitem1TableAdapter
        Dim dta4 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_do_itemTableAdapter
        Dim dta5 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_do_loss_item1TableAdapter
        Dim dta6 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_email_item1TableAdapter
        Dim dta7 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_grn_item1TableAdapter
        Dim dta8 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_purchasereturn_item1TableAdapter
        Dim dta9 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_salesreturn_item1TableAdapter
        Dim dta10 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_stock_taking_itempenyesuaian1TableAdapter
        Dim dta11 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltrans_stock_taking_tag_item1TableAdapter
        Dim dta12 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stockbundling_item_after1TableAdapter
        Dim dta13 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stockbundling_item_before1TableAdapter
        Dim dta14 As New DseTransaksiExpiryDateUpdateTableAdapters.tbltransaksi_stocktransfer_item1TableAdapter
        Try
            If Me.cboBatchNo.Text <> "" Then
                hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    dta.UpdateQuery_ExpiryDateStockinout(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta1.UpdateQuery_ExpiryDateMasterQty(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta2.UpdateQuery_ExpiryDateBookingGITItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta3.UpdateQuery_ExpiryDateBookingInvItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta4.UpdateQuery_ExpiryDateTransaksiDOItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta5.UpdateQuery_ExpiryDateTransaksiDOLossItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta6.UpdateQuery_ExpiryDateTransaksiEmailItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta7.UpdateQuery_ExpiryDateTransaksiGRNItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta8.UpdateQuery_TransaksiPurchaseReturnItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta9.UpdateQuery_ExpiryDateTransaksiSalesReturnItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta10.UpdateQuery_ExpiryTransaksiStockTakingItemPenyesuaian(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta11.UpdateQuery_ExpiryDateTransaksiStockTagItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta12.UpdateQuery_ExpiryDateTransaksiStockBundlingItemAfter(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta13.UpdateQuery_ExpiryDateTransaksiStockBundlingItemBefore(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                    dta14.UpdateQuery_ExpiryDateTransaksiStockTransferItem(Me.DtpNewExpiry.Value, Me.txtPrimaryCode.Text, Me.cboBatchNo.Text)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            dta1.Dispose()
            dta2.Dispose()
            dta3.Dispose()
            dta4.Dispose()
            dta5.Dispose()
            dta6.Dispose()
            dta7.Dispose()
            dta8.Dispose()
            dta9.Dispose()
            dta10.Dispose()
            dta11.Dispose()
            dta12.Dispose()
            dta13.Dispose()
            dta14.Dispose()
        End Try
    End Sub

    Private Sub TxtProductCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductCode.TextChanged
        Try
            Me.TxtProductCode.CharacterCasing = CharacterCasing.Upper

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class