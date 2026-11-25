Option Explicit On
Option Strict Off

Partial Public Class FrmUpdate
    Dim DTA_tblmaster_product As New DseUpdate_ChangeToLocationContrTableAdapters.tblmaster_productTableAdapter
    Dim DTA_tblmaster_product1 As New DseUpdate_ChangeToLocationContrTableAdapters.tblmaster_product1TableAdapter
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjLoc As New Core_Module.clsMasterStorageArea

    Private Sub UpdateAll_LocationControl()
        Try
            Me.prgLocationCont.Value = 0
            Me.prgLocationCont.Maximum = 20

            'tblmaster_product
            Me.DTA_tblmaster_product.UpdateQuery_tblmaster_product()
            Me.chkLocation_tblmaster_product.Checked = True
            Me.prgLocationCont.Value += 1
            'tblmaster_productqty
            Me.DTA_tblmaster_product.UpdateQuery_tblmaster_productqty(Me.TxtLocControl_PrimaryLocation.Text)
            Me.ChkLocation_tblmaster_productqty.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockinout_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stockinout_item(Me.TxtLocControl_PrimaryLocation.Text)
            Me.ChkLocation_tbltransaksi_stockinout_item.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockbundling_item_after
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stockbundling_item_after(Me.TxtLocControl_PrimaryLocation.Text)
            Me.chkLocation_tbltransaksi_stockbundling_item_after.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockbundling_item_before()
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stockbundling_item_before(Me.TxtLocControl_PrimaryLocation.Text)
            Me.ChkLocation_tbltransaksi_stockbundling_item_before.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stocktransfer_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stocktransfer_item1(Me.TxtLocControl_PrimaryLocation.Text)
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stocktransfer_item2(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltransaksi_stocktransfer_item.Checked = True
            'tbltrans_aop_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_aop_item(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_aop_item.Checked = True
            'tbltrans_booking_invitem
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_booking_invitem(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_booking_invitem.Checked = True
            'tbltrans_purchasereturn_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_purchasereturn_item(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_purchasereturn_item.Checked = True
            'tbltrans_salesinvoice_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesinvoice_item(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item.Checked = True
            'tbltrans_salesinvoice_item_copy
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesinvoice_item_copy(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item_copy.Checked = True
            'tbltrans_salesinvoice_item_import
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesinvoice_item_import(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item_import.Checked = True
            'tbltrans_salesreturn_item
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesreturn_item(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item.Checked = True
            'tbltrans_salesreturn_item_copy
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesreturn_item_copy(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item_copy.Checked = True
            'tbltrans_salesreturn_item_import
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_salesreturn_item_import(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item_import.Checked = True
            'tbltrans_stock_taking_itembuffer
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_itembuffer(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itembuffer.Checked = True
            'tbltrans_stock_taking_itempenyesuaian
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_itempenyesuaian(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.Checked = True
            'tbltrans_stock_taking_itemstart
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_itemstart(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itemstart.Checked = True
            'tbltransaksi_stock_card
            Me.DTA_tblmaster_product.UpdateQuery_tbltransaksi_stock_card(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltransaksi_stock_card.Checked = True
            'tbltrans_stock_taking_tag
            Me.DTA_tblmaster_product.UpdateQuery_tbltrans_stock_taking_tag(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_tag.Checked = True

            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateByPrimaryProduct_LocationControl(ByVal vPrimaryProduct As Double)
        Try
            Me.prgLocationCont.Value = 0
            Me.prgLocationCont.Maximum = 20

            'tblmaster_product
            Me.DTA_tblmaster_product1.UpdateQuery_tblmaster_product(vPrimaryProduct)
            Me.chkLocation_tblmaster_product.Checked = True
            Me.prgLocationCont.Value += 1
            'tblmaster_productqty
            Me.DTA_tblmaster_product1.UpdateQuery_tblmaster_productqty(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.ChkLocation_tblmaster_productqty.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockinout_item
            Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stockinout_item(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.ChkLocation_tbltransaksi_stockinout_item.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockbundling_item_after
            Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stockbundling_item_after(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.chkLocation_tbltransaksi_stockbundling_item_after.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stockbundling_item_before(vPrimaryProduct)
            Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stockbundling_item_before(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.ChkLocation_tbltransaksi_stockbundling_item_before.Checked = True
            Me.prgLocationCont.Value += 1
            'tbltransaksi_stocktransfer_item
            Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stocktransfer_item1(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stocktransfer_item2(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltransaksi_stocktransfer_item.Checked = True
            'tbltrans_aop_item
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_aop_item(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_aop_item.Checked = True
            'tbltrans_booking_invitem
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_booking_invitem(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_booking_invitem.Checked = True
            'tbltrans_purchasereturn_item
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_purchasereturn_item(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_purchasereturn_item.Checked = True
            'tbltrans_salesinvoice_item
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesinvoice_item(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item.Checked = True
            'tbltrans_salesinvoice_item_copy
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesinvoice_item_copy(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item_copy.Checked = True
            'tbltrans_salesinvoice_item_import
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesinvoice_item_import(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesinvoice_item_import.Checked = True
            'tbltrans_salesreturn_item
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesreturn_item(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item.Checked = True
            'tbltrans_salesreturn_item_copy
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesreturn_item_copy(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item_copy.Checked = True
            'tbltrans_salesreturn_item_import
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_salesreturn_item_import(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_salesreturn_item_import.Checked = True
            'tbltrans_stock_taking_itembuffer
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_stock_taking_itembuffer(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itembuffer.Checked = True
            'tbltrans_stock_taking_itempenyesuaian
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_stock_taking_itempenyesuaian(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.Checked = True
            'tbltrans_stock_taking_itemstart
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_stock_taking_itemstart(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_itemstart.Checked = True
            'tbltransaksi_stock_card
            Me.DTA_tblmaster_product1.UpdateQuery_tbltransaksi_stock_card(Me.TxtLocControl_PrimaryLocation.Text, vPrimaryProduct)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltransaksi_stock_card.Checked = True
            'tbltrans_stock_taking_tag
            Me.DTA_tblmaster_product1.UpdateQuery_tbltrans_stock_taking_tag(Me.TxtLocControl_PrimaryLocation.Text)
            Me.prgLocationCont.Value += 1
            Me.ChkLocation_tbltrans_stock_taking_tag.Checked = True

            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnLocation_Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLocation_Start.Click
        Try
            If Me.TxtLocControl_PrimaryLocation.Text = "-1" Or Me.TxtLocControl_LocationCode.Text.Trim = "" Then
                MsgBox("Please specify reference location.", MsgBoxStyle.Information, "Information")
                Me.TxtLocControl_LocationCode.Focus()
                Exit Sub
            End If

            If Me.TxtLocControl_TxtPrimaryProduct.Text = -1 Then
                Me.UpdateAll_LocationControl()
            Else
                If Me.ObjProd.ApakahExpiryDateControlled(TxtLocControl_TxtPrimaryProduct.Text) = False Then
                    Me.UpdateByPrimaryProduct_LocationControl(TxtLocControl_TxtPrimaryProduct.Text)
                Else
                    If MsgBox("This product is already location controlled, proceeding this process will reset all location data. Continue anyway?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.UpdateByPrimaryProduct_LocationControl(TxtLocControl_TxtPrimaryProduct.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocControl_CodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocControl_CodeNo.KeyDown
        Dim i As New Core_Module.frmCariMasterProduct
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        Dim UnitPrice, Diskon As Decimal

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtLocControl_CodeNo.Text)
                If Jml = 1 Then
                    Me.TxtLocControl_TxtPrimaryProduct.Text = Me.ObjProduct.GetPrimary(Me.TxtLocControl_CodeNo.Text)
                    Me.TxtLocControl_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtLocControl_TxtPrimaryProduct.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtLocControl_TxtPrimaryProduct.Text = i.PrimaryProduct
                        Me.TxtLocControl_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtLocControl_TxtPrimaryProduct.Text)
                        Me.TxtLocControl_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtLocControl_TxtPrimaryProduct.Text)
                    Else
                        Me.TxtLocControl_CodeNo.Clear()
                        Me.TxtLocControl_TxtPrimaryProduct.Text = "-1"
                        Me.TxtLocControl_CodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtLocControl_CodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocControl_CodeNo.Leave
        Dim i As New Core_Module.frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtLocControl_CodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtLocControl_CodeNo.Text) = 1 Then
                    Me.TxtLocControl_TxtPrimaryProduct.Text = Me.ObjProduct.GetPrimary(Me.TxtLocControl_CodeNo.Text)
                    Me.TxtLocControl_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtLocControl_TxtPrimaryProduct.Text)
                Else
                    If Me.TxtLocControl_CodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocControl_CodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtLocControl_TxtPrimaryProduct.Text = i.PrimaryProduct
                        Me.TxtLocControl_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtLocControl_TxtPrimaryProduct.Text)
                        Me.TxtLocControl_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtLocControl_TxtPrimaryProduct.Text)
                    Else
                        Me.TxtLocControl_CodeNo.Clear()
                        Me.TxtLocControl_TxtPrimaryProduct.Text = "-1"
                        Me.TxtLocControl_CodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtLocControl_TxtPrimaryProduct.Text = "-1"
                Me.TxtLocControl_Desc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnLocControl_FindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLocControl_FindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct 

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll

            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocControl_CodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtLocControl_CodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtLocControl_TxtPrimaryProduct.Text = i.PrimaryProduct
                Me.TxtLocControl_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtLocControl_TxtPrimaryProduct.Text)
                Me.TxtLocControl_Desc.Text = Me.ObjProduct.GetDescription(Me.TxtLocControl_TxtPrimaryProduct.Text)
            Else
                Me.TxtLocControl_CodeNo.Clear()
                Me.TxtLocControl_TxtPrimaryProduct.Text = "-1"
                Me.TxtLocControl_CodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtLocControl_LocationCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocControl_LocationCode.KeyDown
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        Dim UnitPrice, Diskon As Decimal

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjLoc.GetJml_ByCode(Me.TxtLocControl_LocationCode.Text)
                If Jml = 1 Then
                    Me.TxtLocControl_PrimaryLocation.Text = Me.ObjLoc.GetPrimary(Me.TxtLocControl_LocationCode.Text)
                    Me.TxtLocControl_LocationDesc.Text = Me.ObjLoc.GetDescription(Me.TxtLocControl_PrimaryLocation.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.TxtLocControl_PrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocControl_LocationCode.Text = Me.ObjLoc.GetStorageAreaCode(Me.TxtLocControl_PrimaryLocation.Text)
                        Me.TxtLocControl_LocationDesc.Text = Me.ObjLoc.GetDescription(Me.TxtLocControl_PrimaryLocation.Text)
                    Else
                        Me.TxtLocControl_LocationCode.Clear()
                        Me.TxtLocControl_PrimaryLocation.Text = "-1"
                        Me.TxtLocControl_LocationCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtLocControl_LocationCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocControl_LocationCode.Leave
        Dim i As New Core_Module.FrmCariMasterStorageArea
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtLocControl_LocationCode.Text <> "" Then
                If Me.ObjLoc.GetJml_ByCode(Me.TxtLocControl_LocationCode.Text) = 1 Then
                    Me.TxtLocControl_PrimaryLocation.Text = Me.ObjLoc.GetPrimary(Me.TxtLocControl_LocationCode.Text)
                    Me.TxtLocControl_LocationDesc.Text = Me.ObjLoc.GetDescription(Me.TxtLocControl_PrimaryLocation.Text)
                Else
                    If Me.TxtLocControl_LocationCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocControl_LocationCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.TxtLocControl_PrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocControl_LocationCode.Text = Me.ObjLoc.GetStorageAreaCode(Me.TxtLocControl_PrimaryLocation.Text)
                        Me.TxtLocControl_LocationDesc.Text = Me.ObjLoc.GetDescription(Me.TxtLocControl_PrimaryLocation.Text)
                    Else
                        Me.TxtLocControl_LocationCode.Clear()
                        Me.TxtLocControl_PrimaryLocation.Text = "-1"
                        Me.TxtLocControl_LocationCode.Focus()
                    End If
                End If
            Else
                Me.TxtLocControl_PrimaryLocation.Text = "-1"
                Me.TxtLocControl_LocationDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnLocControl_FindLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLocControl_FindLocation.Click
        Dim i As New Core_Module.FrmCariMasterStorageArea 

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll

            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocControl_LocationCode.Text <> "" Then
                i.txtCari.Text = Me.TxtLocControl_LocationCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.TxtLocControl_PrimaryLocation.Text = i.PrimaryStorageArea
                Me.TxtLocControl_LocationCode.Text = Me.ObjLoc.GetStorageAreaCode(Me.TxtLocControl_PrimaryLocation.Text)
                Me.TxtLocControl_LocationDesc.Text = Me.ObjLoc.GetDescription(Me.TxtLocControl_PrimaryLocation.Text)
            Else
                Me.TxtLocControl_LocationCode.Clear()
                Me.TxtLocControl_PrimaryLocation.Text = "-1"
                Me.TxtLocControl_LocationCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub
End Class
