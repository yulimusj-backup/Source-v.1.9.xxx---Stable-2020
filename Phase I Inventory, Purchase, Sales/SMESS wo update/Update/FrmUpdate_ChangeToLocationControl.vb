Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate
    Dim DTA_tblmaster_product As New DseUpdate_ChangeToLocationContrTableAdapters.tblmaster_productTableAdapter

    Private Sub btnLocation_Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLocation_Start.Click
        Try
            Me.prgLocationCont.Value = 0
            Me.prgLocationCont.Maximum = 20

            'tblmaster_product
            Me.DTA_tblmaster_product.UpdateQuery_tblmaster_product()
            Me.chkLocation_tblmaster_product.Checked = True
            Me.prgLocationCont.Value += 1
            'tblmaster_productqty
            Me.DTA_tblmaster_product.UpdateQuery_tblmaster_productqty()
            Me.ChkLocation_tblmaster_productqty.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockinout_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stockinout_item()
            Me.ChkLocation_tbltransaksi_stockinout_item.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockbundling_item_after
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stockbundling_item_after()
            Me.chkLocation_tbltransaksi_stockbundling_item_after.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockbundling_item_before()
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stockbundling_item_before()
            Me.ChkLocation_tbltransaksi_stockbundling_item_before.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stocktransfer_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stocktransfer_item1()
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stocktransfer_item2()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltransaksi_stocktransfer_item.Checked = True
            'tbltrans_aop_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_aop_item()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_aop_item.Checked = True
            'tbltrans_booking_invitem
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_booking_invitem()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_booking_invitem.Checked = True
            'tbltrans_purchasereturn_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_purchasereturn_item()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_purchasereturn_item.Checked = True
            'tbltrans_salesinvoice_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesinvoice_item()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item.Checked = True
            'tbltrans_salesinvoice_item_copy
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesinvoice_item_copy()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item_copy.Checked = True
            'tbltrans_salesinvoice_item_import
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesinvoice_item_import()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item_import.Checked = True
            'tbltrans_salesreturn_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesreturn_item()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item.Checked = True
            'tbltrans_salesreturn_item_copy
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesreturn_item_copy()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item_copy.Checked = True
            'tbltrans_salesreturn_item_import
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesreturn_item_import()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item_import.Checked = True
            'tbltrans_stock_taking_itembuffer
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_itembuffer()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itembuffer.Checked = True
            'tbltrans_stock_taking_itempenyesuaian
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_itempenyesuaian()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.Checked = True
            'tbltrans_stock_taking_itemstart
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_itemstart()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itemstart.Checked = True
            'tbltransaksi_stock_card
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stock_card()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltransaksi_stock_card.Checked = True
            'tbltrans_stock_taking_tag
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_tag()
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_tag.Checked = True

            msgbox("Data updated",MsgBoxStyle.Information,"Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
