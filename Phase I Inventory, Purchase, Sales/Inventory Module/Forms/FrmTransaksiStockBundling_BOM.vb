Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockBundling_BOM
    Dim TBLHasil As Core_Module.dseMasterProduct_BOM.tblmaster_product_item_bom1DataTable
    Public UserCode As String = ""
    Public ObjStockBundling As New ClsTransaksiStockBundling
    Dim ObjItemInOut As New clsTransaksiStockInOut
    Dim ObjItemBundling As New clsTransaksiStockBundlingItem
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim objStockBundlingItem As New clsTransaksiStockBundlingItem
    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL3 As New dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4 As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before1DataTable
    Dim TBL5 As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before2DataTable
    Dim TBL6 As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before3DataTable
    Dim buffer As String = ""

    Private Sub TxtCodeNoAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNoAfter.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0
        Dim PrimaryProduct As Double = -1

        Try
            i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNoAfter.Text)
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)

                If Jml = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                    Me.TxtPrimaryCodeNoAfter.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)
                    Me.txtPrimaryCategoryAfter.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.SetGboItemAfter()
                    If Me.txtBatchNoAfter.Visible = True Then
                        Me.txtBatchNoAfter.Focus()
                    ElseIf Me.TxtLocationAfter.Visible = True Then
                        Me.TxtLocationAfter.Focus()
                    Else
                        Me.TxtQtyperBatchAfter.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNoAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                        Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.SetGboItemAfter()
                        If Me.txtBatchNoAfter.Visible = True Then
                            Me.txtBatchNoAfter.Focus()
                        ElseIf Me.TxtLocationAfter.Visible = True Then
                            Me.TxtLocationAfter.Focus()
                        Else
                            Me.TxtQtyperBatchAfter.Focus()
                        End If
                    Else
                        Me.TxtCodeNoAfter.Clear()
                        Me.TxtPrimaryCodeNoAfter.Text = "-1"
                        Me.txtPrimaryCategoryAfter.Text = "-1"
                        Me.KosongkanItemAfter()
                        Me.TxtCodeNoAfter.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNoAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoAfter.Leave
        Dim i As New frmCariMasterProduct
        Dim PrimaryProduct As Double = -1

        Try
            i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
            If Me.TxtCodeNoAfter.Text <> "" Then
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)
                If Me.ObjProduct.GetJml(Me.TxtCodeNoAfter.Text) = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) Then
                    Me.TxtPrimaryCodeNoAfter.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)
                    Me.txtPrimaryCategoryAfter.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.SetGboItemAfter()
                    If Me.txtBatchNoAfter.Visible = True Then
                        Me.txtBatchNoAfter.Focus()
                    ElseIf Me.TxtLocationAfter.Visible = True Then
                        Me.TxtLocationAfter.Focus()
                    Else
                        Me.TxtQtyperBatchAfter.Focus()
                    End If
                Else
                    If Me.TxtCodeNoAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                        Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.SetGboItemAfter()
                        If Me.txtBatchNoAfter.Visible = True Then
                            Me.txtBatchNoAfter.Focus()
                        ElseIf Me.TxtLocationAfter.Visible = True Then
                            Me.TxtLocationAfter.Focus()
                        Else
                            Me.TxtQtyperBatchAfter.Focus()
                        End If
                    Else
                        Me.TxtCodeNoAfter.Clear()
                        Me.TxtPrimaryCodeNoAfter.Text = "-1"
                        Me.txtPrimaryCategoryAfter.Text = "-1"
                        Me.KosongkanItemAfter()
                        Me.TxtCodeNoAfter.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNoAfter.Text = "-1"
                Me.txtDescAfter.Clear()
                Me.txtPrimaryCategoryAfter.Text = "-1"
                Me.KosongkanItemAfter()
            End If
            Me.SetGboItemAfter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProductAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNoAfter.Click
        Dim i As New frmCariMasterProduct

        Try
            i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNoAfter.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNoAfter.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                Me.SetGboItemAfter()
                If Me.txtBatchNoAfter.Visible = True Then
                    Me.txtBatchNoAfter.Focus()
                ElseIf Me.TxtLocationAfter.Visible = True Then
                    Me.TxtLocationAfter.Focus()
                Else
                    Me.TxtQtyperBatchAfter.Focus()
                End If
            Else
                Me.TxtCodeNoAfter.Clear()
                Me.TxtPrimaryCodeNoAfter.Text = "-1"
                Me.txtPrimaryCategoryAfter.Text = "-1"
                Me.KosongkanItemAfter()
                Me.TxtCodeNoAfter.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub KosongkanItemAfter()
        Try
            ModCekData.KosongkanControl(Me.gboItemAfter.Controls, True, True, False)
            Me.TxtPrimaryCodeNoAfter.Text = "-1"
            Me.txtPrimaryUnitsAfter.Text = "-1"
            Me.txtPrimarySecondaryUnitsAfter.Text = "-1"
            Me.txtPrimaryLocationAfter.Text = "-1"
            Me.txtPrimaryCategoryAfter.Text = "-1"
            Me.TxtQtyperBatchAfter.Text = "0"
            Me.txtSecondaryQtyAfter.Text = "0"
            Me.txtFaktorPemberat.Text = "0"
            Me.chkBatchNoControlledAfter.Checked = False
            Me.chkExpiryDateControlledAfter.Checked = False
            Me.chkLocationControlledAfter.Checked = False
            Me.chkPdoControlledAfter.Checked = False
            Me.chkUseSecondaryUnitAfter.Checked = False
            Me.txtSecondaryQtyAfter.Visible = False
            Me.lblSecondaryUnitsAfter.Visible = False
            Me.TxtQtyperBatchAfter.Width = 239
            Me.lblUnitsAfter.Left = Me.lblSecondaryUnitsAfter.Left
            Me.lblSecQtyAfter.Visible = False
            Me.lblExpiryDateAfter.Visible = False
            Me.DtpExpiredAfter.Value = Today
            Me.DtpExpiredAfter.Visible = True
            Me.lblLocationAfter.Visible = False
            Me.lblWarehouseAfter.Visible = False
            Me.TxtLocationAfter.Visible = True
            Me.TxtWarehouseAfter.Visible = True
            Me.btnFindLocationAfter.Visible = True
            Me.lblBatchNoAfter.Visible = False
            Me.txtBatchNoAfter.Visible = True
            Me.txtBatchNoAfter.TabStop = True
            Me.txtBatchNoAfter.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItemB4()
        Try
            ModCekData.KosongkanControl(Me.gboItemb4.Controls, True, True, False)
            Me.TxtPrimaryCodeNob4.Text = "-1"
            Me.txtPrimaryUnitsb4.Text = "-1"
            Me.txtPrimarySecondaryUnitsb4.Text = "-1"
            Me.txtPrimaryLocationb4.Text = "-1"
            Me.txtPrimaryCategoryb4.Text = "-1"
            Me.TxtQtyperBatchb4.Text = "0"
            Me.txtSecondaryQtyb4.Text = "0"
            Me.chkBatchNoControlledb4.Checked = False
            Me.chkExpiryDateControlledb4.Checked = False
            Me.chkLocationControlledb4.Checked = False
            Me.chkPdoControlledb4.Checked = False
            Me.chkUseSecondaryUnitb4.Checked = False
            Me.lblSecondaryUnitsb4.Visible = False
            Me.TxtQtyperBatchb4.Width = 239
            Me.lblUnitsb4.Left = Me.lblSecondaryUnitsb4.Left
            Me.lblSecQtyb4.Visible = False
            Me.lblExpiryDateb4.Visible = False
            Me.DtpExpiredb4.Visible = True
            Me.DtpExpiredb4.Value = Today
            Me.lblLocationb4.Visible = False
            Me.lblWarehouseb4.Visible = False
            Me.cboLocationb4.DataSource = Nothing
            Me.cboLocationb4.Visible = True
            Me.cboLocationb4.Items.Clear()
            Me.TxtWarehouseb4.Visible = True
            Me.lblBatchNob4.Visible = False
            Me.cboBatchNob4.Visible = True
            Me.cboBatchNob4.TabStop = True
            Me.cboBatchNob4.DataSource = Nothing
            Me.cboBatchNob4.Items.Clear()
            Me.cboSecondaryUnitsb4.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItemAfter()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNoAfter.Text, Me.chkLocationControlledAfter, _
            Me.chkBatchNoControlledAfter, Me.chkExpiryDateControlledAfter, Me.chkUseSecondaryUnitAfter, _
            Me.chkPdoControlledAfter, Me.txtPrimaryUnitsAfter, Me.txtPrimarySecondaryUnitsAfter)

            Me.txtPrimaryPDOAfter.Text = "-1"
            If Me.chkLocationControlledAfter.Checked = False Then
                Me.txtPrimaryLocationAfter.Text = "-1"
                Me.TxtLocationAfter.Text = ""
                Me.TxtWarehouseAfter.Clear()
                Me.TxtLocationAfter.ReadOnly = True
                Me.lblLocationAfter.Visible = True
                Me.lblWarehouseAfter.Visible = True
                Me.TxtLocationAfter.Visible = False
                Me.TxtWarehouseAfter.Visible = False
                Me.btnFindLocationAfter.Visible = False
            Else
                Me.TxtLocationAfter.ReadOnly = False
                Me.lblLocationAfter.Visible = False
                Me.lblWarehouseAfter.Visible = False
                Me.TxtLocationAfter.Visible = True
                Me.TxtWarehouseAfter.Visible = True
                Me.btnFindLocationAfter.Visible = True
            End If
            If Me.chkExpiryDateControlledAfter.Checked = True Then
                Me.lblExpiryDateAfter.Visible = False
                Me.DtpExpiredAfter.Visible = True
                Me.DtpExpiredAfter.Value = Today
            Else
                Me.lblExpiryDateAfter.Visible = True
                Me.DtpExpiredAfter.Visible = False
                Me.DtpExpiredAfter.Value = #1/1/1950#
            End If
            If Me.chkBatchNoControlledAfter.Checked = True Then
                Me.txtBatchNoAfter.ReadOnly = False
                Me.txtBatchNoAfter.TabStop = True
                Me.lblBatchNoAfter.Visible = False
                Me.txtBatchNoAfter.Visible = True
            Else
                Me.txtBatchNoAfter.ReadOnly = True
                Me.txtBatchNoAfter.TabStop = False
                Me.lblBatchNoAfter.Visible = True
                Me.txtBatchNoAfter.Visible = False
            End If
            If Me.chkUseSecondaryUnitAfter.Checked = False Then
                Me.txtSecondaryQtyAfter.Visible = False
                Me.lblSecondaryUnitsAfter.Visible = False
                Me.lblUnitsAfter.Left = Me.lblSecondaryUnitsAfter.Left
                Me.TxtQtyperBatchAfter.Width = Me.lblUnitsAfter.Left - Me.TxtQtyperBatchAfter.Left - 6
                Me.txtSecondaryQtyAfter.Text = "-1"
                Me.lblSecQtyAfter.Visible = False
            Else
                Me.txtSecondaryQtyAfter.Visible = True
                Me.lblSecondaryUnitsAfter.Visible = True
                Me.lblSecQtyAfter.Visible = True
                Me.lblUnitsAfter.Left = Me.lblSecQtyAfter.Left - 37
                Me.TxtQtyperBatchAfter.Width = Me.lblUnitsAfter.Left - Me.TxtQtyperBatchAfter.Left - 6
                Me.txtSecondaryQtyAfter.Text = "0"
            End If
            If Me.TxtPrimaryCodeNoAfter.Text <> "-1" Then
                Me.lblUnitsAfter.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNoAfter.Text)
                If Me.chkUseSecondaryUnitAfter.Checked = True Then
                    Me.lblSecondaryUnitsAfter.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNoAfter.Text)
                End If
            Else
                Me.lblUnitsAfter.Text = "Units"
                Me.lblSecondaryUnitsAfter.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItemb4()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNob4.Text, Me.chkLocationControlledb4, _
            Me.chkBatchNoControlledb4, Me.chkExpiryDateControlledb4, Me.chkUseSecondaryUnitb4, _
            Me.chkPdoControlledb4, Me.txtPrimaryUnitsb4, Me.txtPrimarySecondaryUnitsb4)

            Me.cboBatchNob4.DataSource = Nothing
            Me.cboLocationb4.DataSource = Nothing
            Me.cboSecondaryUnitsb4.DataSource = Nothing

            Me.txtPrimaryPDOb4.Text = "-1"
            If Me.chkBatchNoControlledb4.Checked = True Then
                Me.cboBatchNob4.TabStop = True
                Me.lblBatchNob4.Visible = False
                Me.cboBatchNob4.Visible = True
                'isikan data batch no dari basis data
                Me.IsiCboItem(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockIn.Text)
            Else
                Me.cboBatchNob4.TabStop = False
                Me.lblBatchNob4.Visible = True
                Me.cboBatchNob4.Visible = False
                Me.cboBatchNob4.Items.Clear()
            End If
            If Me.chkLocationControlledb4.Checked = False Then
                Me.txtPrimaryLocationb4.Text = "-1"
                Me.cboLocationb4.Items.Clear()
                Me.TxtWarehouseb4.Clear()
                Me.lblLocationb4.Visible = True
                Me.lblWarehouseb4.Visible = True
                Me.cboLocationb4.Visible = False
                Me.TxtWarehouseb4.Visible = False
            Else
                Me.lblLocationb4.Visible = False
                Me.lblWarehouseb4.Visible = False
                Me.cboLocationb4.Visible = True
                Me.TxtWarehouseb4.Visible = True
                If Me.chkLocationControlledb4.Checked = True Then
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockIn.Text, Me.cboBatchNob4.Text)
                End If
            End If
            If Me.chkExpiryDateControlledb4.Checked = True Then
                Me.lblExpiryDateb4.Visible = False
                Me.DtpExpiredb4.Visible = True
            Else
                Me.lblExpiryDateb4.Visible = True
                Me.DtpExpiredb4.Visible = False
                Me.DtpExpiredb4.Value = #1/1/1950#
            End If

            If Me.chkUseSecondaryUnitb4.Checked = False Then
                Me.lblSecondaryUnitsb4.Visible = False
                Me.lblUnitsb4.Left = Me.lblSecondaryUnitsb4.Left
                Me.TxtQtyperBatchb4.Width = Me.lblUnitsb4.Left - Me.TxtQtyperBatchb4.Left - 6
                Me.txtSecondaryQtyb4.Text = "-1"
                Me.lblSecQtyb4.Visible = False
                Me.cboSecondaryUnitsb4.Visible = False
                Me.cboSecondaryUnitsb4.DataSource = Nothing
                Me.cboSecondaryUnitsb4.Items.Clear()
            Else
                Me.lblSecondaryUnitsb4.Visible = True
                Me.lblSecQtyb4.Visible = True
                Me.lblUnitsb4.Left = Me.lblSecQtyb4.Left - 37
                Me.TxtQtyperBatchb4.Width = Me.lblUnitsb4.Left - Me.TxtQtyperBatchb4.Left - 6
                Me.cboSecondaryUnitsb4.Visible = True
                If Me.chkLocationControlledb4.Checked = False Then
                    'isikan data secondary jika perlu
                    Me.IsiCboSecondary(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockIn.Text, _
                    Me.cboBatchNob4.Text, Me.txtPrimaryLocationb4.Text)
                End If
            End If
            If Me.TxtPrimaryCodeNob4.Text <> "-1" Then
                Me.lblUnitsb4.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text)
                If Me.chkUseSecondaryUnitb4.Checked = True Then
                    Me.lblSecondaryUnitsb4.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNob4.Text)
                End If
            Else
                Me.lblUnitsb4.Text = "Units"
                Me.lblSecondaryUnitsb4.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboLokasi(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, _
   ByVal vBatchNo As String)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboLocation(vPrimaryProduct, Me.cboLocationb4, _
            Me.TBL2, vBatchNo)
            Dim row1 As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before2Row
            Dim row2 As dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockBundlingItem.IsikanTBLLocation(Me.TBL5, vPrimaryProduct, vPrimaryTransaksi, _
                vBatchNo)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL5.Rows.Count > 0
                    row1 = Me.TBL5.Rows(0)
                    For Each row2 In Me.TBL2.Rows
                        If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL2.Rows.Add(row1.Lot_No, row1.Location, row1.Qty)
                    End If
                    Me.TBL5.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLLocation(vPrimaryProduct, vBatchNo)
            While Me.TBL5.Rows.Count > 0
                row1 = Me.TBL5.Rows(0)
                For Each row2 In Me.TBL2.Rows
                    If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL2.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL2.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL5.Rows.Remove(row1)
            End While

            Me.TBL2.DefaultView.Sort = "Location ASC"
            Me.cboLocationb4.DataSource = Me.TBL2.DefaultView
            Me.cboLocationb4.ValueMember = "Location"
            Me.cboLocationb4.DisplayMember = "Location"
            Me.cboLocationb4.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboItem(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Try
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNob4, Me.TBL)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before1Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockBundlingItem.IsikanTBLBatchNo(Me.TBL4, vPrimaryProduct, vPrimaryTransaksi)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    For Each row2 In Me.TBL.Rows
                        If row1.Lot_No = row2.Lot_No Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL.Rows.Add(row1.Lot_No, row1.Expiry_Date, row1.Qty)
                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLBatchNo(vPrimaryProduct)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                For Each row2 In Me.TBL.Rows
                    If row1.Lot_No = row2.Lot_No Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4.Rows.Remove(row1)
            End While
            Me.TBL.DefaultView.Sort = "[Lot No] ASC"
            Me.cboBatchNob4.DataSource = Me.TBL.DefaultView
            Me.cboBatchNob4.ValueMember = "Lot No"
            Me.cboBatchNob4.DisplayMember = "Lot No"
            Me.cboBatchNob4.Refresh()

            If Me.chkLocationControlledb4.Checked = False And _
            Me.chkUseSecondaryUnitb4.Checked = True And Me.cboBatchNob4.Text <> "" Then
                Me.IsiCboSecondary(vPrimaryProduct, vPrimaryTransaksi, _
                Me.cboBatchNob4.Text, Me.txtPrimaryLocationb4.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboSecondary(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, _
  ByVal vBatchNo As String, ByVal vPrimaryLokasi As Double)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboSecondaryUnit(vPrimaryProduct, Me.cboSecondaryUnitsb4, _
            Me.TBL3, vBatchNo, vPrimaryLokasi)
            Dim row1 As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before3Row
            Dim row2 As dseMasterProduct.tblmaster_productqty2Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockBundlingItem.IsikanTBLSecondaryUnit(Me.TBL6, vPrimaryProduct, _
                vPrimaryTransaksi, vBatchNo, vPrimaryLokasi)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL6.Rows.Count > 0
                    row1 = Me.TBL6.Rows(0)
                    For Each row2 In Me.TBL3.Rows
                        If row1.Secondary_Qty = row2.Qty_Secondary Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL3.Rows.Add(row1.Secondary_Qty, row1.Qty)
                    End If
                    Me.TBL6.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLSecondary(vPrimaryProduct, vPrimaryLokasi, vBatchNo)
            While Me.TBL6.Rows.Count > 0
                row1 = Me.TBL6.Rows(0)
                For Each row2 In Me.TBL3.Rows
                    If row1.Secondary_Qty = row2.Qty_Secondary Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL3.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL3.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL6.Rows.Remove(row1)
            End While

            Me.TBL3.DefaultView.Sort = "[Qty Secondary] ASC"
            Me.cboSecondaryUnitsb4.DataSource = Me.TBL3.DefaultView
            Me.cboSecondaryUnitsb4.ValueMember = "Qty Secondary"
            Me.cboSecondaryUnitsb4.DisplayMember = "Qty Secondary"
            Me.cboSecondaryUnitsb4.Refresh()
            'pilih salah satu
            If Me.cboSecondaryUnitsb4.Items.Count > 0 Then
                Me.cboSecondaryUnitsb4.SelectedIndex = -1
                Me.cboSecondaryUnitsb4.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLSecondary(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before3Row
            For Each row In Me.dgvItemb4.Rows
                If CDbl(row.Cells(12).Value) = vPrimaryProduct And CStr(row.Cells(7).Value) = vBatchNo And _
                CDbl(row.Cells(13).Value) = vPrimaryLocation And IsNumeric(row.Cells(5).Value) = True Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL6.Rows
                        If row1.Secondary_Qty = CDbl(row.Cells(5).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(3).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(8).Value) = True Then
                            ED = row.Cells(8).Value
                        End If
                        Me.TBL6.Rows.Add(row.Cells(5).Value, row.Cells(3).Value, _
                        CStr(row.Cells(9).Value), CStr(row.Cells(7).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLLocation(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before2Row
            For Each row In Me.dgvItemb4.Rows
                If CDbl(row.Cells(12).Value) = vPrimaryProduct And CStr(row.Cells(7).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL5.Rows
                        If row1.Lot_No = CStr(row.Cells(7).Value) And row1.Location = row.Cells(9).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(3).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL5.Rows.Add(CStr(row.Cells(7).Value), row.Cells(9).Value, _
                        CDbl(row.Cells(3).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Double)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before1Row
            For Each row In Me.dgvItemb4.Rows
                If CDbl(row.Cells(12).Value) = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Lot_No = CStr(row.Cells(7).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(3).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(8).Value) = True Then
                            ED = row.Cells(8).Value
                        End If
                        Me.TBL4.Rows.Add(CStr(row.Cells(7).Value), CDbl(row.Cells(3).Value), ED)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtCodeNoAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtCodeNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtCodeNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoAfter.LostFocus
        Try
            Me.TxtCodeNoAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNoAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNoAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.txtBatchNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.txtBatchNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNoAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNoAfter.Leave
        Try
            Me.txtBatchNoAfter.Text = Me.txtBatchNoAfter.Text.Trim
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNoAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNoAfter.LostFocus
        Try
            Me.txtBatchNoAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtLocationAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtLocationAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationAfter.LostFocus
        Try
            Me.TxtLocationAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtQtyperBatchAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtQtyperBatchAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.LostFocus
        Try
            Me.TxtQtyperBatchAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.txtSecondaryQtyAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.txtSecondaryQtyAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyAfter.LostFocus
        Try
            Me.txtSecondaryQtyAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtRemarkItemAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtRemarkItemAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemAfter.LostFocus
        Try
            Me.TxtRemarkItemAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.Enter
        Try
            Me.ObjCD.Number_Enter(Me.txtFaktorPemberat, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.GotFocus
        Try
            Me.txtFaktorPemberat.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFaktorPemberat.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtFaktorPemberat, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFaktorPemberat.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtFaktorPemberat, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.Leave
        Try
            Me.txtFaktorPemberat.Text = Me.ObjCD.DigitGrouping(Me.txtFaktorPemberat.Text, 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.LostFocus
        Try
            Me.txtFaktorPemberat.BackColor = Me.txtPrimaryCategoryAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.Leave
        Try
            If Me.TxtPrimaryCodeNoAfter.Text <> "" And Me.TxtPrimaryCodeNoAfter.Text <> "" Then
                Me.TxtQtyperBatchAfter.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchAfter.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNoAfter.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocationAfter.KeyDown
        Dim i As New FrmCariMasterStorageArea
        Dim Obj As New clsMasterStorageArea

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtLocationAfter.Text) = True Then
                    Me.txtPrimaryLocationAfter.Text = Obj.GetPrimary(Me.TxtLocationAfter.Text)
                    Me.TxtWarehouseAfter.Text = Obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                    Me.TxtQtyperBatchAfter.Focus()
                Else
                    If Me.TxtLocationAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocationAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocationAfter.Text = i.PrimaryStorageArea
                        Me.TxtWarehouseAfter.Text = Obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtLocationAfter.Text = Obj.GetStorageAreaCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtQtyperBatchAfter.Focus()
                    Else
                        Me.TxtLocationAfter.Clear()
                        Me.txtPrimaryLocationAfter.Text = "-1"
                        Me.TxtWarehouseAfter.Clear()
                        Me.TxtLocationAfter.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtLocationAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationAfter.Leave
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            If Me.TxtLocationAfter.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtLocationAfter.Text) = True Then
                    Me.TxtQtyperBatchAfter.Focus()
                    Me.txtPrimaryLocationAfter.Text = obj.GetPrimary(Me.TxtLocationAfter.Text)
                    Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                Else
                    If Me.TxtLocationAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocationAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocationAfter.Text = i.PrimaryStorageArea
                        Me.TxtLocationAfter.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtQtyperBatchAfter.Focus()
                    Else
                        Me.TxtLocationAfter.Clear()
                        Me.TxtWarehouseAfter.Clear()
                        Me.txtPrimaryLocationAfter.Text = "-1"
                        Me.TxtLocationAfter.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryLocationAfter.Text = "-1"
                Me.TxtWarehouseAfter.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindLocationAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocationAfter.Click
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocationAfter.Text <> "" Then
                i.txtCari.Text = Me.TxtLocationAfter.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.txtPrimaryLocationAfter.Text = i.PrimaryStorageArea
                Me.TxtLocationAfter.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocationAfter.Text)
                Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                Me.TxtQtyperBatchAfter.Focus()
            Else
                Me.TxtWarehouseAfter.Clear()
                Me.TxtLocationAfter.Clear()
                Me.TxtLocationAfter.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub FrmTransaksiStockBundling_BOM_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.KosongkanItemAfter()
            Me.KosongkanItemB4()
            Me.SetGboItemAfter()
            Me.SetGboItemb4()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class