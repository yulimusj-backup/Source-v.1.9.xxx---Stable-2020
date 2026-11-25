Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTransaksiSalesInvoiceQuick_AddItembyCustom
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    '<i> jgn kelewat diisi dari form utama
    Public PrimaryInvoice As Double = -1
    Public PrimaryCustomer As Double = -1
    Public PrimaryCurrency As Double = -1
    Public Tax As Decimal = 0
    Public IsTax As Boolean = False
    Public TblDummy As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable
    Public DefaultPrecision As Decimal = 0
    Public DefaultPrecision_DefCur As Decimal = 0
    '
    Dim baru As Boolean
    Dim obj As New ClsTransaksiSalesInvoiceQuick_AddByCustomProduct
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCustomProd As New ClsMasterProductCustom
    Dim ObjProduct As New Core_Module.clsMasterProduct
    'Dim ObjInvoiceItem As New ClsTransaksiSalesInvoiceItem
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim objDO As New ClsTransaksiDO
    Dim ObjLocation As New clsMasterStorageArea
    'Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim TBL_Item As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2_Item As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL3_Item As New dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4_Item As New DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item2DataTable
    Dim TBL5_Item As New DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item3DataTable

    Public MetodeHitungInclude As EnumHitungInclude = EnumHitungInclude.NettPriceInclude
    Dim Buffer2 As String = ""
    Dim Buffer As String = ""
    Dim mTblMultiUnit As New DseTransaksiSO.tblmaster_product_item_multiunitDataTable
    Dim tblCboUnit As New DseTransaksiSO.CboUnitDataTable
    Dim Exit_ As Boolean = True
    Public AktifForm As Boolean = True
    Public Enum EnumHitungInclude As Integer
        UnitPriceSO
        UnitPriceInclude
        NettPriceInclude
    End Enum

    Public Function ApakahDataBenar() As Boolean
        Try
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtCustomCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomCode.KeyDown
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomProd.ApakahSudahAda(Me.TxtCustomCode.Text) = True Then
                    Me.TxtPrimaryCustomCode.Text = Me.ObjCustomProd.GetPrimary(Me.TxtCustomCode.Text)
                    Me.TxtCustomCode.Text = Me.ObjCustomProd.GetCode(Me.TxtPrimaryCustomCode.Text)
                    Me.TxtPrimaryCodeNo.Text = Me.obj.GetPrimaryProductByCustom(Me.TxtPrimaryCustomCode.Text)
                    Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.ProductCodeLeave()
                Else
                    Me.BtnFindCustomProduct.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub TxtCustomCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomCode.Leave
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            If Me.TxtCustomCode.Text <> "" Then
                If Me.ObjCustomProd.ApakahSudahAda(Me.TxtCustomCode.Text) = True Then
                    Me.TxtPrimaryCustomCode.Text = Me.ObjCustomProd.GetPrimary(Me.TxtCustomCode.Text)
                    Me.TxtCustomCode.Text = Me.ObjCustomProd.GetCode(Me.TxtPrimaryCustomCode.Text)
                    Me.TxtPrimaryCodeNo.Text = Me.obj.GetPrimaryProductByCustom(Me.TxtPrimaryCustomCode.Text)
                    Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.ProductCodeLeave()
                Else
                    Me.BtnFindCustomProduct.PerformClick()
                End If
            Else
                Me.TxtPrimaryCustomCode.Text = -1
                Me.TxtCustomCode.Text = ""
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.KosongkanItem()
            End If
            Me.TxtCustomCode.BackColor = Me.TxtPrimaryCustomCode.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub BtnFindCustomProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCustomProduct.Click
        Dim FrmCariGroup As New FrmCariMasterProductCustom
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimaryCustomCode.Text = -1 And Me.TxtCustomCode.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.TxtCustomCode.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDbl(FrmCariGroup.Primary) > 0 Then
                Me.TxtPrimaryCustomCode.Text = FrmCariGroup.Primary
                Me.TxtCustomCode.Text = Me.ObjCustomProd.GetCode(FrmCariGroup.Primary)
                Me.TxtPrimaryCodeNo.Text = Me.obj.GetPrimaryProductByCustom(Me.TxtPrimaryCustomCode.Text)
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.ProductCodeLeave()
            Else
                Me.TxtPrimaryCustomCode.Text = -1
                Me.TxtCustomCode.Text = ""
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Sub ProductCodeLeave()
        Try
            Dim i As New frmCariMasterProduct
            Dim UnitPrice, Diskon As Decimal
            Try
                If Me.TxtCodeNo.Text <> "" Then
                    If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                        Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.cboBatchNo.Visible = True Then
                            Me.cboBatchNo.Focus()
                        ElseIf Me.cboLocation.Visible = True Then
                            Me.cboLocation.Focus()
                        Else
                            If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                Me.TxtQtyDefault.Focus()
                            Else
                                Me.TxtCurrencyCode1.Focus()
                            End If
                        End If
                        Me.ObjProduct.GetUnitPriceNDiskon(PrimaryCustomer, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.Hitung_Amount2()
                    Else
                        If Me.TxtCodeNo.Text <> "" Then
                            i.txtCari.Text = Me.TxtCodeNo.Text
                        End If

                        i.BackColorAll = Me.BackColorAll
                        i.ShowDialog()
                        If i.PrimaryProduct > 0 Then
                            Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                            Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                            Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                            Me.SetGboItem()
                            If Me.cboBatchNo.Visible = True Then
                                Me.cboBatchNo.Focus()
                            ElseIf Me.cboLocation.Visible = True Then
                                Me.cboLocation.Focus()
                            Else
                                If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                    Me.TxtQtyDefault.Focus()
                                Else
                                    Me.TxtCurrencyCode1.Focus()
                                End If
                            End If
                            Me.ObjProduct.GetUnitPriceNDiskon(PrimaryCustomer, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                            Me.TxtDiscFromPerct.Text = "1"
                            Me.Hitung_Amount2()
                        Else
                            Me.TxtCodeNo.Clear()
                            Me.TxtPrimaryCodeNo.Text = "-1"
                            Me.KosongkanItem()
                            Me.TxtCodeNo.Focus()
                        End If
                    End If
                Else
                    Me.TxtPrimaryCodeNo.Text = "-1"
                    Me.txtDesc.Clear()
                    Me.KosongkanItem()
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                i.Dispose()
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.TxtPrimaryUnitsDefault)

            Me.cboBatchNo.DataSource = Nothing
            Me.cboLocation.DataSource = Nothing

            If Me.chkBatchNoControlled.Checked = True Then
                Me.cboBatchNo.TabStop = True
                Me.lblBatchNo.Visible = False
                Me.cboBatchNo.Visible = True
                'isikan data batch no dari basis data
                Me.IsiCboItem(Me.TxtPrimaryCodeNo.Text, Me.PrimaryInvoice)
            Else
                Me.cboBatchNo.TabStop = False
                Me.lblBatchNo.Visible = True
                Me.cboBatchNo.Visible = False
                Me.cboBatchNo.Items.Clear()
            End If

            If Me.chkLocationControlled.Checked = False Then
                Me.txtPrimaryLocation.Text = "-1"
                Me.cboLocation.Items.Clear()
                Me.lblLocation.Visible = True
                Me.cboLocation.Visible = False
            Else
                Me.lblLocation.Visible = False
                Me.cboLocation.Visible = True
                If Me.chkLocationControlled.Checked = True Then
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNo.Text, Me.PrimaryInvoice, Me.cboBatchNo.Text)
                End If
            End If

            If Me.chkExpiryDateControlled.Checked = True Then
                Me.lblExpiryDate.Visible = False
                Me.DtpExpired.Visible = True
            Else
                Me.lblExpiryDate.Visible = True
                Me.DtpExpired.Visible = False
                Me.DtpExpired.Value = #1/1/1950#
            End If

            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.lblUnitsDefault.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
            Else
                Me.lblUnitsDefault.Text = "Units"
            End If

            Me.TxtQtyDefault.Text = "0"
            Me.TxtQtyMulti.Text = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboItem(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Try
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNo, Me.TBL_Item)

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item2Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean

            If Me.baru = False Then
                Me.IsikanTBLBatchNo(Me.TBL4_Item, vPrimaryProduct, vPrimaryTransaksi)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4_Item.Rows.Count > 0
                    row1 = Me.TBL4_Item.Rows(0)
                    For Each row2 In Me.TBL_Item.Rows
                        If row1.Lot_No = row2.Lot_No Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL_Item.Rows.Add(row1.Lot_No, row1.Expiry_Date, row1.Qty)
                    End If
                    Me.TBL4_Item.Rows.Remove(row1)
                End While
            End If

            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLBatchNo(vPrimaryProduct)
            While Me.TBL4_Item.Rows.Count > 0
                row1 = Me.TBL4_Item.Rows(0)
                For Each row2 In Me.TBL_Item.Rows
                    If row1.Lot_No = row2.Lot_No Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL_Item.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL_Item.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4_Item.Rows.Remove(row1)
            End While

            If Me.chkExpiryDateControlled.Checked = False Then
                Me.TBL_Item.DefaultView.Sort = "[Lot No] ASC"
            Else
                Me.TBL_Item.DefaultView.Sort = "[Expiry Date] ASC"
            End If

            Me.cboBatchNo.DataSource = Me.TBL_Item.DefaultView
            Me.cboBatchNo.ValueMember = "Lot No"
            Me.cboBatchNo.DisplayMember = "Lot No"
            Me.cboBatchNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboLokasi(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboLocation(vPrimaryProduct, Me.cboLocation, _
            Me.TBL2_Item, vBatchNo)
            Dim row1 As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item3Row
            Dim row2 As dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.IsikanTBLLocation(Me.TBL5_Item, vPrimaryProduct, vPrimaryTransaksi, vBatchNo)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL5_Item.Rows.Count > 0
                    row1 = Me.TBL5_Item.Rows(0)
                    For Each row2 In Me.TBL2_Item.Rows
                        If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL2_Item.Rows.Add(row1.Lot_No, row1.Location, row1.Qty)
                    End If
                    Me.TBL5_Item.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLLocation(vPrimaryProduct, vBatchNo)
            While Me.TBL5_Item.Rows.Count > 0
                row1 = Me.TBL5_Item.Rows(0)
                For Each row2 In Me.TBL2_Item.Rows
                    If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL2_Item.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL2_Item.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL5_Item.Rows.Remove(row1)
            End While

            Me.TBL2_Item.DefaultView.Sort = "Location ASC"
            Me.cboLocation.DataSource = Me.TBL2_Item.DefaultView
            Me.cboLocation.ValueMember = "Location"
            Me.cboLocation.DisplayMember = "Location"
            Me.cboLocation.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Double)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item2Row

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4_Item.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColLotNumber.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next

                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(Me.ColExpiry.Index).Value) = True Then
                            ED = row.Cells(Me.ColExpiry.Index).Value
                        End If
                        Me.TBL4_Item.Rows.Add(CStr(row.Cells(Me.ColLotNumber.Index).Value), _
                        CDbl(row.Cells(Me.ColTotalQty.Index).Value), ED)
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
            Dim row1 As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item3Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                CStr(row.Cells(Me.ColLotNumber.Index).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL5_Item.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColLotNumber.Index).Value) _
                        And row1.Location = row.Cells(Me.ColLocationCode.Index).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next

                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL5_Item.Rows.Add(CStr(row.Cells(Me.ColLotNumber.Index).Value), _
                        row.Cells(Me.ColLocationCode.Index).Value, _
                        CDbl(row.Cells(Me.ColTotalQty.Index).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Amount2()
        Dim DecimalPlace As Integer
        Dim Amount As Decimal = 0
        Dim NettPrice As Decimal = 0
        Dim DiscPerct, DiscAmount As Decimal

        Try
            If Me.TxtPrimaryCurrency1.Text <> "" And Me.TxtPrimaryCurrency1.Text <> "-1" Then
                DecimalPlace = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)
            Else
                DecimalPlace = 0
            End If

            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, DecimalPlace)

            If Me.TxtDiscFromPerct.Text = "1" Then
                DiscAmount = Math.Round((CDec(Me.TxtUnitPrice.Text) * (CDec(Me.TxtDiscount.Text)) / 100), DecimalPlace, MidpointRounding.AwayFromZero)
                Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(DiscAmount, DecimalPlace)
            Else
                DiscPerct = Math.Round((CDec(Me.TxtDiscAmount1.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2, MidpointRounding.AwayFromZero)
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(DiscPerct, 2)
            End If

            NettPrice = Math.Round(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount1.Text), DecimalPlace, MidpointRounding.AwayFromZero)
            Amount = Math.Round(NettPrice * CDec(Me.TxtQtyMulti.Text), DecimalPlace, MidpointRounding.AwayFromZero)

            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, DecimalPlace)
            Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Amount, DecimalPlace)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.RowCount = 0 Then
                If IsNumeric(PrimaryBaseCurrency) = True Then
                    Me.TxtPrimaryCurrency1.Text = PrimaryBaseCurrency
                Else
                    Me.TxtPrimaryCurrency1.Text = "1"
                End If
            Else
                row = Me.dgvItem.Rows(0)
                Me.TxtPrimaryCurrency1.Text = row.Cells(ColPrimaryCurrency.Index).Value
            End If
            Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)

            Me.TxtUnitPrice.Text = "0"
            Me.TxtDiscount.Text = "0"
            Me.txtNettPrice.Text = "0"
            Me.TxtQtyDefault.Text = "0"
            Me.TxtQtyMulti.Text = "0"
            Me.txtAmount.Text = "0"
            Me.TxtDiscAmount1.Text = "0"
            Me.TxtDiscFromPerct.Text = "1"
            Me.TxtRemarkItem.Text = ""
            Me.TxtCodeNo.Text = ""
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.TxtPrimaryCustomCode.Text = "-1"
            Me.txtDesc.Text = ""
            Me.TxtCustomCode.Text = ""
            Me.SetGboItem()

            'kosongkan juga cbo lokasi dan cbo batch number
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Refresh()
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCurrency2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCurrency1.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                Me.TxtUnitPrice.Focus()
                Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.TxtCurrencyCode1.Text <> "" Then
                    i.txtCari.Text = Me.TxtCurrencyCode1.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                    Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                    Me.TxtUnitPrice.Focus()
                Else
                    Me.TxtCurrencyCode1.Clear()
                    Me.TxtCurrencyCode1.Focus()
                End If
            End If

            If Me.TxtRemarkItem.Enabled = True Then
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCurrencyCode1.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
                    Me.TxtUnitPrice.Focus()
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
                'Me.SetCalculateBy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.Leave
        Dim i As New FrmCariMasterCurrency

        Try
            If Me.TxtCurrencyCode1.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCurrency1.Text = "-1"
            End If

            Hitung_Amount2()
            'Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try

        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '<i> diblock dulu calculatenya.
    'Private Sub SetCalculateBy()
    '    Try
    '        If Me.TxtRemarks.Enabled = True Then
    '            'reset data item exchange rate
    '            Me.dgvExchageRates.Rows.Clear()
    '            If Me.cboCalculatesBy.SelectedIndex = 0 Then
    '                'hitung berdasarkan sales invoice
    '                If Me.dgvItem.RowCount = 0 Then
    '                    'kalo ga ada anak, maka exc
    '                    Me.dgvExchageRates.Rows.Add("0", Me.txtInvoiceNo.Text, Me.txtCurrencyCode.Text, _
    '                    1, "", Me.txtPrimaryInvoice.Text, Me.txtPrimaryCurrency.Text)
    '                Else
    '                    Dim rowDGV As DataGridViewRow
    '                    rowDGV = Me.dgvItem.Rows(0)
    '                    Me.dgvExchageRates.Rows.Add("0", Me.txtInvoiceNo.Text, rowDGV.Cells(Me.ColCurrency.Index).Value, _
    '                    1, "", Me.txtPrimaryInvoice.Text, rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
    '                End If
    '            ElseIf Me.cboCalculatesBy.SelectedIndex = 1 Then
    '                'hitung berdasarkan delivery order
    '                Dim Ketemu As Boolean = False
    '                Dim rowER, rowItem As DataGridViewRow

    '                For Each rowItem In Me.dgvItem.Rows
    '                    Ketemu = False
    '                    For Each rowER In Me.dgvExchageRates.Rows
    '                        If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimaryDO.Index).Value Then
    '                            Ketemu = True
    '                            Exit For
    '                        End If
    '                    Next
    '                    If Ketemu = False Then
    '                        'tambahkan exchange rate baru 
    '                        Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(Me.ColDONO.Index).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
    '                        1, "", rowItem.Cells(Me.ColPrimaryDO.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
    '                    End If
    '                Next
    '            ElseIf Me.cboCalculatesBy.SelectedIndex = 2 Then
    '                'hitung berdasarkan sales order
    '                Dim Ketemu As Boolean = False
    '                Dim rowER, rowItem As DataGridViewRow

    '                For Each rowItem In Me.dgvItem.Rows
    '                    Ketemu = False
    '                    For Each rowER In Me.dgvExchageRates.Rows
    '                        If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimarySO.Index).Value Then
    '                            Ketemu = True
    '                            Exit For
    '                        End If
    '                    Next
    '                    If Ketemu = False Then
    '                        'tambahkan exchange rate baru 
    '                        Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(Me.ColSONO.Index).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
    '                        1, "", rowItem.Cells(Me.ColPrimarySO.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
    '                    End If
    '                Next
    '            End If
    '            If Me.txtCurrencyCode.Text <> "" Then
    '                Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate to " + Me.txtCurrencyCode.Text
    '            Else
    '                Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate"
    '            End If
    '            'hitung ulang
    '            Me.Hitung_Total()
    '            'set number
    '            Me.ObjCD.SetNumber(Me.dgvExchageRates)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub TxtDownTimeCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.GotFocus, txtDesc.GotFocus, TxtUnitPrice.GotFocus, _
    TxtDiscount.GotFocus, TxtDiscAmount1.GotFocus, TxtRemarkItem.GotFocus, TxtCustomCode.GotFocus, TxtQtyMulti.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDownTimeCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.LostFocus, txtDesc.LostFocus, TxtUnitPrice.LostFocus, _
    TxtDiscount.LostFocus, TxtDiscAmount1.LostFocus, TxtRemarkItem.LostFocus, TxtCustomCode.LostFocus, TxtQtyMulti.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryCodeNo.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceQuick_AddByCustomProduct_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.Exit_ = False Then
                Me.DialogResult = Windows.Forms.DialogResult.No
                If Me.AktifForm = True Then
                    If MsgBox("Exit without save?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceQuick_AddByCustomProduct_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.TxtCustomCode.Select()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.BtnFindCustomProduct.Left = Me.SplitContainer5.Panel1.Width - Me.BtnFindCustomProduct.Width
            Me.BtnFindCodeNo.Left = Me.SplitContainer5.Panel1.Width - Me.BtnFindCodeNo.Width
            Me.TxtCustomCode.Width = Me.BtnFindCodeNo.Left - 5 - Me.TxtCodeNo.Left
            Me.TxtCodeNo.Width = Me.SplitContainer5.Panel1.Width - Me.TxtCodeNo.Left
            Me.txtDesc.Width = Me.SplitContainer5.Panel1.Width - Me.txtDesc.Left
            Me.cboBatchNo.Width = Me.txtDesc.Width
            Me.DtpExpired.Width = Me.cboBatchNo.Width
            Me.cboLocation.Width = Me.SplitContainer5.Panel1.Width - Me.cboLocation.Left

            Me.txtAmount.Width = Me.SplitContainer5.Panel2.Width - Me.txtAmount.Left
            Me.btnInput.Left = Me.txtAmount.Left + Me.txtAmount.Width - Me.btnInput.Width
            Me.TxtRemarkItem.Width = Me.btnInput.Left - 5 - Me.TxtRemarkItem.Left
            Me.txtNettPrice.Left = Me.txtAmount.Left + Me.txtAmount.Width - Me.txtNettPrice.Width
            Me.TxtDiscount.Left = Me.txtNettPrice.Left
            Me.TxtQtyMulti.Left = Me.txtNettPrice.Left
            Me.LblDisc.Left = Me.TxtDiscount.Left - 5 - Me.LblDisc.Width
            Me.LblQty.Left = Me.LblDisc.Left
            Me.Label33.Left = Me.LblQty.Left
            Me.CboUnit.Left = Me.txtAmount.Left + Me.txtAmount.Width - Me.CboUnit.Width
            Me.btnFindCurrency1.Left = Me.LblDisc.Left - Me.btnFindCurrency1.Width - 5
            Me.TxtCurrencyCode1.Width = Me.btnFindCurrency1.Left - 5 - Me.TxtCurrencyCode1.Left
            Me.TxtUnitPrice.Width = Me.Label33.Left - 5 - Me.TxtUnitPrice.Left
            Me.TxtDiscAmount1.Width = Me.TxtUnitPrice.Width

            Me.ColVAT.Visible = False
            Me.ColVAT2.HeaderText = "VAT"

            If Me.IsTax = True Then
                Me.ColVAT2.Visible = True
                Me.ColAmount2.Visible = True
            Else
                Me.ColVAT2.Visible = False
                Me.ColAmount2.Visible = False
            End If
            Me.KosongkanItem()
            If Me.TblDummy.Rows.Count > 0 Then
                Me.obj.IsiDGVFromTblDummy(Me.dgvItem, Me, Me.PrimaryInvoice, Me.TblDummy)
                Me.FormatDGV2()
            End If

            If Me.AktifForm = False Then
                Me.GroupBox2.Enabled = False
                Me.dgvItem.AllowUserToDeleteRows = False
                Me.dgvItem.ReadOnly = True
                Me.BtnOK.Visible = False
                Me.BtnCancel.Text = "Close"
                'Me.TxtRemarkItem.Width = Me.txtAmount.Width
            End If
            Me.CancelButton = Me.BtnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1

        Try
            If Me.ApakahDataItemBenar = True Then
                indexSama = Me.GetIndex
                Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
                If indexSama = -1 Then
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                    AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)
                    With row
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColPrimaryCustomProduct.Index).Value = Me.TxtPrimaryCustomCode.Text
                        .Cells(Me.ColCustomCodeNo.Index).Value = Me.TxtCustomCode.Text
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColCustCodeNo.Index).Value = ""
                        .Cells(Me.ColCustDesc.Index).Value = ""
                        .Cells(Me.ColSONO.Index).Value = ""
                        .Cells(Me.ColCustPONo.Index).Value = ""
                        .Cells(Me.ColDONO.Index).Value = ""
                        .Cells(Me.ColStockOutNo.Index).Value = ""
                        .Cells(Me.ColLotNumber.Index).Value = Me.cboBatchNo.Text
                        If Me.DtpExpired.Visible = False Then
                            .Cells(Me.ColExpiry.Index).Value = #1/1/1950#
                        Else
                            .Cells(Me.ColExpiry.Index).Value = Me.DtpExpired.Value
                        End If
                        .Cells(Me.ColLocationCode.Index).Value = Me.cboLocation.Text
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyDefault.Text), AngkaDecimal)
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColSecQty.Index).Value = -1
                        .Cells(Me.ColSecUnit.Index).Value = ""
                        .Cells(Me.ColCurrency.Index).Value = Me.TxtCurrencyCode1.Text
                        .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscount.Index).Value = CDec(Me.TxtDiscount.Text)
                        .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtAmount.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColRemarksItem.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryDO.Index).Value = -1
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimarySO.Index).Value = -1
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryUnit.Index).Value = Me.TxtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                        .Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency1.Text
                        .Cells(Me.ColAliasDesc.Index).Value = Me.obj.GetAliasDesc(Me.TxtPrimaryCodeNo.Text, Me.PrimaryCustomer)
                        .Cells(Me.ColAliasCode.Index).Value = Me.obj.GetAliasCode(Me.TxtPrimaryCodeNo.Text, Me.PrimaryCustomer)

                        .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtDiscAmount1.Text, AngkaDecimalCurrency)
                        .Cells(Me.ColDiscFromPrect.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtDiscFromPerct.Text, AngkaDecimalCurrency)

                        .Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, DefPrecMulti)
                        .Cells(Me.ColUnitsMulti.Index).Value = Me.CboUnit.Text
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.TxtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(Me.obj.GetRatioDefault(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), AngkaDecimal)
                        .Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(Me.obj.GetRatioReference(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecMulti)
                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((CDec(Me.txtNettPrice.Text) * Me.Tax) / 100, AngkaDecimalCurrency)
                        If AngkaDecimalCurrency = 0 Then
                            If Me.txtNettPrice.Text.EndsWith("55") Then
                                .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT.Index).Value) - 1, AngkaDecimalCurrency)
                            End If
                        End If
                        .Cells(Me.ColAngkaDecimal.Index).Value = AngkaDecimal
                        .Cells(Me.ColDefPrecCurr.Index).Value = AngkaDecimalCurrency
                        .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), AngkaDecimalCurrency)
                        Me.SetRatioOnDgv(row)

                    End With
                Else
                    row = Me.dgvItem.Rows(indexSama)
                    'tambahkan qty multi, 
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(Me.TxtQtyMulti.Text), DefPrecMulti)
                    'hitung ulang amount
                    row = Me.dgvItem.Rows(indexSama)

                    Me.SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, row.Index).Value))
                    Me.Hitung_Amount(row)
                End If

                PrimaryCurrenct = Me.TxtPrimaryCurrency1.Text
                Me.KosongkanItem()
                Me.TxtCustomCode.Focus()

                Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
                Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                Me.FormatDGV2()

                'Me.cboCalculatesBy.SelectedIndex = -1
                'Me.cboCalculatesBy.SelectedIndex = 0

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                '<i> yang dibawah dilakukan setelah form add close
                'If Me.ApakahAdaAlias = True Then
                '    Me.chkAlias.Checked = True
                'Else
                '    Me.chkAlias.Checked = False
                'End If

                'Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim rowDGV As DataGridViewRow
        Dim QtyGudang, QtyDGV, QtyDatabaseItem, JmlMaks, QtyPending As Decimal
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim AngkaDecimal As Integer = 0

        Try
            'data salah jika:
            '1. tidak diisikan, kode product, currency
            'cek data yang harus diisi
            KoleksiNamaData.Add("Custom Code")
            KoleksiTextBox.Add(Me.TxtCustomCode)
            KoleksiNamaData.Add("Code no")
            KoleksiTextBox.Add(Me.TxtCodeNo)
            KoleksiNamaData.Add("Currency code")
            KoleksiTextBox.Add(Me.TxtCurrencyCode1)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. jika sudah ada item, maka kode currencynya harus sama
            If Me.dgvItem.Rows.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(0)
                If Me.TxtPrimaryCurrency1.Text <> rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value Then
                    MsgBox("Currency must be identical for each item", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '3. berkaitan dengan data angka:
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            '   a. unit price, qty harus positif
            KoleksiNamaData.Add("Unit Price")
            KoleksiTextBox.Add(Me.TxtUnitPrice)
            KoleksiNamaData.Add("Qty")
            KoleksiTextBox.Add(Me.TxtQtyMulti)
            If ModCekData.CekBilanganPositif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            '   3.b. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyDefault.Text) = False Then
                MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyMulti.Focus()
                Return False
            End If
            '   b. discount paling besar 100% --> otomatis + non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            KoleksiNamaData.Add("discount")
            KoleksiTextBox.Add(Me.TxtDiscount)
            If ModCekData.CekNonNegatif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '4. Remarks maksimal 150 karakter
            If Me.TxtRemarkItem.Text.Length > 150 Then
                MsgBox("Maximum remarks data length is 150 character", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan data stock tidak minus
            'lakukan pengecekan
            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNo.Text) = True Then
                QtyGudang = Me.ObjProduct.GetJml2(Me.TxtPrimaryCodeNo.Text, DTA1, -1, _
                            Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                QtyDGV = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text, _
                Me.txtPrimaryLocation.Text, Me.TxtPrimaryCustomCode.Text)
                QtyDatabaseItem = Me.obj.GetJml_Default(Me.PrimaryInvoice, Me.TxtPrimaryCodeNo.Text, _
                Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                QtyPending = Me.obj.GetJmlPending_Default(Me.TxtPrimaryCodeNo.Text)

                JmlMaks = QtyGudang + QtyDatabaseItem - QtyDGV - QtyPending
                If JmlMaks < 0 Then
                    JmlMaks = 0
                End If

                If CDec(Me.TxtQtyDefault.Text) > JmlMaks Then
                    'Dim Max_ As String = ObjCD.DigitGrouping(JmlMaks * Me.ObjSO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text) / Me.ObjSO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                    Dim Max_ As String = JmlMaks * Me.obj.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text) / Me.obj.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
                    Max_ = Me.ObjCD.RoundDown1(Max_, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))

                    MsgBox("Maximum qty data is " & Max_ & " " + Me.ObjUnit.GetUnitCode(Me.TxtPrimaryUnitsMulti.Text), _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtQtyMulti.Focus()
                    Return False
                End If
            End If


            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            DTA1.Dispose()
        End Try
    End Function

    Private Function GetIndex() As Integer
        Try
            Return Me.GetIndex(Me.TxtUnitPrice.Text, Me.cboBatchNo.Text, Me.TxtDiscount.Text, _
            Me.txtPrimaryLocation.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryCustomCode.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetIndex(ByVal vUnitPrice As Decimal, ByVal vBatchNumber As String, ByVal vDiscount As Decimal, _
   ByVal vPrimaryLocation As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryCustomerCode As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And vBatchNumber = row.Cells(Me.ColLotNumber.Index).Value _
                            And vDiscount = row.Cells(Me.ColDiscount.Index).Value And vPrimaryLocation = row.Cells(Me.ColPrimaryLocation.Index).Value And _
                            vPrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value And vPrimaryCustomerCode = row.Cells(Me.ColPrimaryCustomProduct.Index).Value Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetRatioOnDgv(ByVal vDgvRow As DataGridViewRow)
        Try
            With vDgvRow
                If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnit.Index).Value Then
                    .Cells(Me.ColRatioDesc.Index).Value = .Cells(Me.ColRatioReference.Index).Value & " " & .Cells(Me.ColUnitsMulti.Index).Value & " = " & _
                    .Cells(Me.ColRatioDefault.Index).Value & " " & .Cells(Me.ColUnit.Index).Value
                Else
                    .Cells(Me.ColRatioDesc.Index).Value = "-"
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgv1(ByVal vRow As DataGridViewRow, ByVal vDefPrecDefault As Integer)
        Try
            vRow.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(vRow.Cells(Me.ColTotalQtyMulti.Index).Value) * CDec(vRow.Cells(Me.ColRatioDefault.Index).Value) / CDec(vRow.Cells(Me.ColRatioReference.Index).Value), vDefPrecDefault)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Hitung_Amount(ByVal row As DataGridViewRow)
        Dim DefaultPrecisionRow As Integer
        Dim DiscPerct As Integer = 0

        Try
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            DiscPerct = row.Cells(Me.ColDiscFromPrect.Index).Value
            If DiscPerct = 1 Then
                row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecisionRow))
            Else
                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)), 2))
            End If
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(row.Cells(Me.ColUnitPrice.Index).Value) - CDec(row.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecisionRow))

            'hitung amount
            If Me.MetodeHitungInclude = EnumHitungInclude.UnitPriceInclude Then
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                    (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), _
                    DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
                    DefaultPrecisionRow))
                Else
                    If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100) _
                        * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow, MidpointRounding.AwayFromZero), DefaultPrecisionRow))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
                        DefaultPrecisionRow))
                    End If
                End If
            ElseIf Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, _
                    DefaultPrecisionRow))
                Else
                    If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value _
                        * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value, _
                        DefaultPrecisionRow))
                    End If
                End If
            End If

            With row
                .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((CDec(.Cells(Me.ColNettPrice.Index).Value) * Me.Tax) / 100, DefaultPrecisionRow)
                If DefaultPrecisionRow = 0 Then
                    If .Cells(Me.ColNettPrice.Index).Value.ToString.EndsWith("55") Then
                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT.Index).Value) - 1, DefaultPrecisionRow)

                    End If
                End If
                .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecisionRow)
                .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), DefaultPrecisionRow)
            End With


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV2()
        Dim row As DataGridViewRow
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim IsMultiUnit As Boolean = False

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiry.Index).Value) = False Then
                    row.Cells(Me.ColExpiry.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiry.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiry.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiry.Index).Value = Format(CDate(row.Cells(Me.ColExpiry.Index).Value), "dd MMM yyyy")
                    End If
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(Me.ColPrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.ColLocationCode.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColLotNumber.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False

            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(Me.ColLocationCode.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocationCode.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColLotNumber.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLotNumber.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiry.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiry.Index).Visible = True
            End If

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryUnitsMulti.Index).Value <> row.Cells(Me.ColPrimaryUnit.Index).Value Then
                    IsMultiUnit = True
                    Exit For
                End If
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

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlDGV_Default(ByVal vPrimaryProduct As Double, ByVal vBatchNumber As String, ByVal vPrimaryLocation As Double, ByVal vPrimaryCustom As Double) As Decimal
        Dim row As DataGridViewRow
        Dim retVal As Double = 0

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row.Cells(Me.ColLotNumber.Index).Value = vBatchNumber And _
                row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                row.Cells(Me.ColPrimaryCustomProduct.Index).Value = vPrimaryCustom Then
                    retVal += CDec(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtQtyMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyMulti, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Leave
        Dim UnitPrice, Diskon As Decimal

        Try
            If Me.TxtRemarkItem.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                End If

                Me.ObjProduct.GetUnitPriceNDiskon(Me.PrimaryCustomer, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                                               Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)

                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
                Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                Me.TxtDiscFromPerct.Text = "1"
                Me.SetQtyDefault()
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If IsNumeric(Me.TxtQtyMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                SetQtyDefault()
                If CDec(TxtUnitPrice.Text) > 0 Then
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Dim RatioDefault As Decimal = 0
        Dim RatioReference As Decimal = 0
        Try
            RatioDefault = Me.obj.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
            RatioReference = Me.obj.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
            Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyMulti.Text) * RatioDefault / RatioReference, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.Buffer2 = Me.TxtUnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e, 9999999999)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, Me.DefaultPrecision)
                If CDec(Me.TxtDiscAmount1.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtUnitPrice.Text = Me.Buffer2
                    Exit Sub
                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            Me.PrepareMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrepareMultiUnit()
        Dim Dta As New DseTransaksiSOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Try
            Me.CboUnit.DataSource = Nothing
            If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                Dta.Connection.ConnectionString = Core_Module.GetConstring
                Dta.FillBy_PrimaryProduct(Me.mTblMultiUnit, Me.TxtPrimaryCodeNo.Text)
                Me.IsiCboUnit(Me.TxtPrimaryCodeNo.Text, Me.CboUnit, Me.tblCboUnit, Me.mTblMultiUnit)
                Me.CboUnit.SelectedIndex = -1
                Me.CboUnit.SelectedIndex = 0
            Else
                TxtPrimaryUnitsDefault.Text = "-1"
                TxtPrimaryUnitsMulti.Text = "-1"
                lblUnitsDefault.Text = "Units"
                'lblUnitsMulti.Text = "Units"
                TxtQtyDefault.Text = "0"
                TxtQtyMulti.Text = "0"
                LblUnitPrice.Text = "Unit Price"
                CboUnit.DataSource = Nothing
                CboUnit.Refresh()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsiCboUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
           ByVal vTbl As DseTransaksiSO.CboUnitDataTable, _
           ByVal vTblMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitDataTable) As Boolean
        Dim rowTbl As DseTransaksiSO.CboUnitRow
        Dim rowMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitRow
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objUnit As New Core_Module.ClsMasterUnit
        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vTbl.Rows.Clear()
            vTblMultiUnit.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
            rowTbl = vTbl.NewCboUnitRow
            rowTbl.Primary = objProduct.GetPrimaryUnit(vPrimaryProduct)
            rowTbl.Code = objUnit.GetUnitCode(rowTbl.Primary)
            vTbl.AddCboUnitRow(rowTbl)
            If vTblMultiUnit.DefaultView.Count > 0 Then
                For i As Integer = 0 To vTblMultiUnit.DefaultView.Count - 1
                    rowMultiUnit = vTblMultiUnit.DefaultView(i).Row
                    rowTbl = vTbl.NewCboUnitRow
                    rowTbl.Primary = rowMultiUnit.Primary_Unit
                    rowTbl.Code = objUnit.GetUnitCode(rowMultiUnit.Primary_Unit)
                    vTbl.AddCboUnitRow(rowTbl)
                Next
            End If
            vCbo.DataSource = vTbl.DefaultView
            vTbl.DefaultView.Sort = "[Code] ASC"
            vCbo.DisplayMember = "Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub CboUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnit.SelectedIndexChanged
        Dim UnitPrice, Discount As Decimal
        Dim row As DseTransaksiSO.CboUnitRow
        Try
            If Me.CboUnit.SelectedIndex <> -1 Then
                row = Me.tblCboUnit.DefaultView(Me.CboUnit.SelectedIndex).Row
                Me.TxtPrimaryUnitsMulti.Text = row.Primary
                'Me.lblUnitsMulti.Text = Me.CboUnit.Text
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnit.Text
                Me.TxtQtyMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                Me.ObjProduct.GetUnitPriceNDiskon(Me.PrimaryCustomer, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                Me.PrimaryCurrency, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Discount)
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.PrimaryCurrency))
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.SetQtyDefault()
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLocation.SelectedIndexChanged
        Dim rowLocation As dseMasterProduct.stock_location_viewRow
        Try
            If Me.cboLocation.SelectedIndex <> -1 Then
                rowLocation = TryCast(Me.TBL2_Item.DefaultView(Me.cboLocation.SelectedIndex).Row, _
                dseMasterProduct.stock_location_viewRow)
                Me.txtPrimaryLocation.Text = Me.ObjLocation.GetPrimary(rowLocation.Location)
            Else
                Me.txtPrimaryLocation.Text = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim objItem As New ClsTransaksiSalesInvoiceQuick_AddByCustomProduct
        Try
            Me.Exit_ = True
            If Me.ApakahDataBenar() = True Then
                If MsgBox("Save Data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    'If baru = True Then
                    '    Me.obj.Simpan(Me.dgvItem, Me, Me.PrimaryInvoice)
                    'Else
                    '    Me.obj.HapusItem(Me.PrimaryInvoice)
                    '    Me.obj.Simpan(Me.dgvItem, Me, Me.PrimaryInvoice)
                    'End If
                    objItem.IsiTblDummyFromDgv(Me.dgvItem, Me, Me.PrimaryInvoice, Me.TblDummy)
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            Me.Exit_ = False
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang boleh diedit hanya qty, unit price, remarks item, dan discount
            If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or _
            e.ColumnIndex = Me.ColRemarksItem.Index Or e.ColumnIndex = Me.ColDiscount.Index Or _
            e.ColumnIndex = Me.ColAliasDesc.Index Or e.ColumnIndex = Me.ColAliasCode.Index Or _
            e.ColumnIndex = Me.ColDiscAmount.Index Then
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                'buang nol di belakang decimal jika edit kolom qty
                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or _
                e.ColumnIndex = Me.ColDiscAmount.Index Or _
                e.ColumnIndex = Me.ColDiscount.Index Or _
                e.ColumnIndex = Me.ColUnitPrice.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Try

            If Me.AktifForm = False Then
                Exit Sub
            End If
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                Dim FrmRatio As New FrmTransaksiQuotation_SetQtyRatio
                Dim row As DataGridViewRow

                If Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value Then
                    Exit Sub
                End If
                Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value.ToString
                Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value.ToString
                FrmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                FrmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                FrmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                FrmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                FrmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                FrmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                If FrmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = FrmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = FrmRatio.TxtRatioDefault.Text
                    row = Me.dgvItem.Rows(e.RowIndex)
                    Me.SetRatioOnDgv(row)
                    Me.SetQtyDefaultOnDgv1(row, FrmRatio.vDefPrecDefault)

                    ''batal jika qty default jadi 0
                    If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Ratio restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.SetRatioOnDgv(row)
                        Me.SetQtyDefaultOnDgv1(row, FrmRatio.vDefPrecDefault)
                        Exit Sub
                    End If

                    ''batal jika perubahan ratio menyebabkan stok gudang minus
                    If Me.ApakahEditSah(row.Index, False) = False Then
                        MsgBox("Ratio restored to previous value, because it make qty will be minus.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.SetRatioOnDgv(row)
                        Me.SetQtyDefaultOnDgv1(row, FrmRatio.vDefPrecDefault)
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim row As DataGridViewRow
        Dim AngkaDecimal, IndexTujuan As Integer

        Try
            row = Me.dgvItem.Rows(e.RowIndex)
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'bila tidak ada perubahan maka tidak perlu lakukan apa2
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If

                    'bila ada perubahan:
                    If Me.ColTotalQtyMulti.Index = e.ColumnIndex Or Me.ColUnitPrice.Index = e.ColumnIndex Or _
                    Me.ColDiscount.Index = e.ColumnIndex Or Me.ColDiscAmount.Index = e.ColumnIndex Then
                        If e.ColumnIndex = Me.ColDiscount.Index Then
                            'pastikan nilai discount non negatif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                            'pastikan nilai diskon maksimum = 100
                            If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) >= 100 Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            If CDec(Me.buffer) = CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            row.Cells(Me.ColDiscFromPrect.Index).Value = "1"
                        ElseIf e.ColumnIndex = Me.ColDiscAmount.Index Then
                            'untuk nilai disc amount, pastikan nilainya positif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for discount amount.", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            'digit grouping
                            AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                            Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)

                            If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                                'jika unit price < nett price maka kembalikan ke nilai buffer
                                'hanya berlaku klo menentukan diskon berdasarkan amount
                                If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                                'tandai diskon dihitung berdasarkan amount
                                row.Cells(Me.ColDiscFromPrect.Index).Value = "0"
                                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                                               ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)), 2))
                            Else
                                Exit Sub
                            End If
                        ElseIf e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            Else
                                'pastikan nilai angka tidak lebih dari 9 M
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 9999999999 Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            'pastikan nilai unit price dan qty positif
                            If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            Else
                                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
                                Else
                                    AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                End If
                                'digit grouping
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)
                            End If
                        End If

                        'jika qty maka cek berkaitan dengan stock
                        If e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                            Me.SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                            If CDec(row.Cells(Me.ColTotalQty.Index).Value) <= 0 Then
                                MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                                Exit Sub
                            End If
                            If Me.ApakahEditSah(e.RowIndex) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                                Exit Sub
                            End If
                        End If

                        'jika unit price, pastikan disc amount < unit price
                        If e.ColumnIndex = Me.ColUnitPrice.Index Then
                            If Me.dgvItem.Item(Me.ColDiscFromPrect.Index, e.RowIndex).Value = "0" Then
                                If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > _
                                CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If

                        IndexTujuan = Me.GetIndexSama(e.RowIndex)
                        If IndexTujuan <> -1 Then
                            'gabungkan baris
                            Me.GabungkanBaris(e.RowIndex, IndexTujuan)
                        Else
                            'jika lewat semua maka hitung amount
                            Me.Hitung_Amount(Me.dgvItem.Rows(e.RowIndex))
                        End If
                        'hitung ulang total
                        'Me.Hitung_Total()
                    ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Or e.ColumnIndex = Me.ColAliasDesc.Index Or _
                    e.ColumnIndex = Me.ColAliasCode.Index Then
                        'jika kolom remarks maka cek, length maksimum = 150
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Length > 150 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Substring(0, 150)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarksItem.Index Or _
                    e.ColumnIndex = Me.ColAliasDesc.Index Or _
                    e.ColumnIndex = Me.ColAliasCode.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.TabStop = True
            'If Me.TxtRemarks.Enabled = True Then
            '    Me.btnUP.Enabled = True
            '    Me.btnDown.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                'If Me.TxtRemarks.Enabled = True Then
                '    Me.btnUP.Enabled = False
                '    Me.btnDown.Enabled = False
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            'If Me.TxtRemarks.Enabled = True Then
            '    If Me.dgvItem.SelectedRows.Count > 0 Then
            '        Me.btnUP.Enabled = True
            '        Me.btnDown.Enabled = True
            '    Else
            '        Me.btnUP.Enabled = False
            '        Me.btnDown.Enabled = False
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)
            'hitung ulang nilai total
            'If Me.TxtRemarks.Enabled = True Then
            '    Me.Hitung_Total()
            'End If
            'If Me.chkAlias.Checked = True Then
            '    If Me.ApakahAdaAlias = True Then
            '        Me.chkAlias.Checked = True
            '    Else
            '        Me.chkAlias.Checked = False
            '    End If
            'End If
            'If Me.dgvItem.RowCount <= 0 Then
            '    Me.txtCustomerCode.ReadOnly = False
            '    Me.btnFindCustomer.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Me.buffer = e.Row.Cells(Me.ColPrimaryDO.Index).Value
    End Sub

    Private Function ApakahEditSah(ByVal rowIndex As Integer, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim row, rowEdit As DataGridViewRow
        Dim JmlDGV, JmlGudang, JmlDatabase, JmlPending As Decimal
        Dim PrimaryCodeNo, PrimaryLocation, PrimaryCustom As Double
        Dim BatchNumber As String = ""
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim AngkaDecimal As Integer = 0

        Try
            rowEdit = Me.dgvItem.Rows(rowIndex)
            'cek qty
            If Me.ObjProduct.ApakahGoods(rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                BatchNumber = rowEdit.Cells(Me.ColLotNumber.Index).Value
                PrimaryCodeNo = rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value
                PrimaryLocation = rowEdit.Cells(Me.ColPrimaryLocation.Index).Value
                AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(PrimaryCodeNo)
                PrimaryCustom = rowEdit.Cells(Me.ColPrimaryCustomProduct.Index).Value

                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColPrimaryLocation.Index).Value = PrimaryLocation And _
                    row.Cells(Me.ColPrimaryCodeNo.Index).Value = PrimaryCodeNo And _
                    row.Cells(Me.ColLotNumber.Index).Value = BatchNumber And _
                    row.Cells(Me.ColPrimaryCustomProduct.Index).Value = PrimaryCustom Then
                        JmlDGV += CDec(row.Cells(Me.ColTotalQty.Index).Value)
                    End If
                Next
                JmlGudang = Me.ObjProduct.GetJml2(PrimaryCodeNo, DTA1, -1, PrimaryLocation, BatchNumber)
                JmlDatabase = Me.obj.GetJml_Default(Me.PrimaryInvoice, PrimaryCodeNo, BatchNumber, PrimaryLocation)
                JmlPending = Me.obj.GetJmlPending_Default(PrimaryCodeNo)


                If JmlGudang + JmlDatabase - JmlDGV - JmlPending < 0 Then
                    If Me.ObjCD.DigitGrouping((CDec(rowEdit.Cells(Me.ColTotalQty.Index).Value) + _
                    (JmlGudang + JmlDatabase - JmlDGV - JmlPending)), AngkaDecimal) < 0 Then
                        If vShowMsg = True Then
                            MsgBox("Maximum qty you can enter is: " + Me.buffer + " " + rowEdit.Cells(Me.ColUnitsMulti.Index).Value.ToString)
                        End If
                    Else
                        Dim MaxDefault As Decimal = CDec(rowEdit.Cells(Me.ColTotalQty.Index).Value) + (JmlGudang + JmlDatabase - JmlDGV - JmlPending)
                        'Dim MaxMulti As String = Me.ObjCD.DigitGrouping(MaxDefault * CDec(rowEdit.Cells(Me.ColRatioReference.Index).Value) / CDec(rowEdit.Cells(Me.ColRatioDefault.Index).Value), ObjProduct.GetAngkaDecimal(rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value, rowEdit.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                        Dim MaxMulti As String = MaxDefault * CDec(rowEdit.Cells(Me.ColRatioReference.Index).Value) / CDec(rowEdit.Cells(Me.ColRatioDefault.Index).Value)
                        MaxMulti = ObjCD.RoundDown1(MaxMulti, ObjProduct.GetAngkaDecimal(rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value, rowEdit.Cells(Me.ColPrimaryUnitsMulti.Index).Value))

                        If vShowMsg = True Then
                            MsgBox("Maximum qty you can enter is: " + MaxMulti + " " + rowEdit.Cells(Me.ColUnitsMulti.Index).Value.ToString)
                        End If
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
        End Try
    End Function

    Private Function GetIndexSama(ByVal rowIndexAcuan As Integer) As Integer
        Dim rowDGV, rowAcuan As DataGridViewRow

        Try
            rowAcuan = Me.dgvItem.Rows(rowIndexAcuan)
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> rowIndexAcuan Then
                    With rowAcuan
                        If .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        .Cells(Me.ColUnitPrice.Index).Value = rowDGV.Cells(Me.ColUnitPrice.Index).Value And _
                        .Cells(Me.ColDiscount.Index).Value = rowDGV.Cells(Me.ColDiscount.Index).Value And _
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value And _
                        .Cells(Me.ColLotNumber.Index).Value = rowDGV.Cells(Me.ColLotNumber.Index).Value And _
                        .Cells(Me.ColPrimaryUnit.Index).Value = rowDGV.Cells(Me.ColPrimaryUnit.Index).Value And _
                        .Cells(Me.ColPrimaryCurrency.Index).Value = rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value And _
                        .Cells(Me.ColPrimaryCustomProduct.Index).Value = rowDGV.Cells(Me.ColPrimaryCustomProduct.Index).Value Then
                            Return rowDGV.Index
                        End If
                    End With
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GabungkanBaris(ByVal vIndexAsal As Integer, ByVal vIndexTujuan As Integer) As Integer
        Dim rowAsal, rowTujuan As DataGridViewRow

        Try
            rowAsal = Me.dgvItem.Rows(vIndexAsal)
            rowTujuan = Me.dgvItem.Rows(vIndexTujuan)
            'tambahakan qty pada rowtujuan
            rowTujuan.Cells(Me.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowAsal.Cells(Me.ColTotalQtyMulti.Index).Value) + _
            CDec(rowTujuan.Cells(Me.ColTotalQtyMulti.Index).Value), Me.ObjProduct.GetAngkaDecimal(rowTujuan.Cells(Me.ColPrimaryCodeNo.Index).Value, rowTujuan.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
            SetQtyDefaultOnDgv1(rowTujuan, Me.ObjProduct.GetAngkaDecimal(rowTujuan.Cells(Me.ColPrimaryCodeNo.Index).Value))
            'hitung ulang amount pada row tujuan
            Me.Hitung_Amount(rowTujuan)
            'buang baris row asal
            Me.dgvItem.Rows.Remove(rowAsal)
            'set ulang nomor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'kembalikan index rowtujuan
            Return rowTujuan.Index
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtDiscount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If Me.TxtRemarkItem.Enabled = True Then
                If IsNumeric(Me.TxtDiscount.Text) = True And IsNumeric(Me.TxtUnitPrice.Text) = True And IsNumeric(Me.TxtQtyDefault) = True Then
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)
                If CDec(Me.TxtDiscount.Text) <> CDec(Me.Buffer2) Then
                    Me.TxtDiscFromPerct.Text = "1"
                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            Me.Buffer2 = Me.TxtDiscount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.Enter
        Try
            Me.Buffer2 = Me.TxtDiscAmount1.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount1.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscAmount1, sender, e, Me.TxtUnitPrice.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount1.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscAmount1, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscAmount1.Text, Me.DefaultPrecision)
                If CDec(Me.TxtDiscAmount1.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtDiscAmount1.Text = Me.Buffer2
                    Exit Sub
                End If
                If CDec(Me.TxtDiscAmount1.Text) <> CDec(Me.Buffer2) Then
                    Me.TxtDiscFromPerct.Text = "0"
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNo.SelectedIndexChanged
        Try
            If Me.chkLocationControlled.Checked = True Then
                If Me.cboBatchNo.Text <> "" Then
                    'isikan data lokasi
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNo.Text, PrimaryInvoice, Me.cboBatchNo.Text)
                End If
            Else
                'kosongkan data lokasi
                Me.cboLocation.DataSource = Nothing
                Me.cboLocation.Items.Clear()
            End If
            'isikan data expiry date 
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.DtpExpired.Value = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            Else
                If Me.DtpExpired.Visible = True Then
                    Me.DtpExpired.Value = Today
                Else
                    Me.DtpExpired.Value = #1/1/1950#
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsikanTBLBatchNo(ByVal TBL As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item2DataTable, _
   ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New DseTransaksiSalesInvoiceQuick_AddByCustomProductTableAdapters.tbltrans_salesinvoice_item2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLLocation(ByVal TBL As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_salesinvoice_item3DataTable, _
ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Dim DTA As New DseTransaksiSalesInvoiceQuick_AddByCustomProductTableAdapters.tbltrans_salesinvoice_item3TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

End Class