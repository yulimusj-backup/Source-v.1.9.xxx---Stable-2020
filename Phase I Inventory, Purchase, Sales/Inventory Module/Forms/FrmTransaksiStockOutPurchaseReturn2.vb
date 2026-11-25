Option Strict Off
Option Explicit On

Imports Core_Module

Public Class FrmTransaksiStockOutPurchaseReturn2
    Public UserCode As String = ""
    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
    Public FormNo As Integer = 0
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objStockOut As New clsTransaksiStockInOut
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim objStockOutOutItem As New clsTransaksiStockInOutItem
    Dim ObjPR As New ClsTransaksiPurchaseReturn
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim ObjLicense As New ClsLicense
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim ObjGRN As New Inventory_Module.ClsTransaksiGRN
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjUser As New ClsUser

    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL3 As New dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item1DataTable
    Dim TBL5 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item2DataTable
    Dim TBL6 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item3DataTable
    Dim objPO As New ClsTransaksiPO
    Dim ObjSupplier As New clsMasterSupplier
    Dim buffer As String = ""


    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)

            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)

            ''@rosadinnn
            ''@MUaT
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)

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
            Me.DtpExpired.Enabled = False
            Me.SetGboItem()
            Me.FormatDGV_Multi()
            Me.Atur1()

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
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)

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

            Me.TxtPrimaryStockIn.Text = "-1"
            Me.TxtPrimaryPurchaseReturn.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.DGVPurchaseReturn.Rows.Clear()

            
            Me.TblItemCustom.Clear()
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

            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"

            ''@rosadinnn
            ''@MUaT
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.TxtRatioReference.Text = "1"
            Me.TxtRatioDefault.Text = "1"

            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimaryPO.Text = "-1"
            Me.TxtQtyperBatchMulti.Text = "0"

            ''@rosadinnn
            ''@MUaT
            Me.TxtQtyperBatchDefault.Text = "0"

            Me.txtSecondaryQty.Text = "-1"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.lblSecondaryUnits.Visible = False

            ''@rosadinnn
            ''@MUaT
            Me.lblUnitsMulti.Left = Me.BtnInput.Left
            Me.LblQty.Text = "Qty"

            Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
            Me.lblSecQty.Visible = False
            Me.lblExpiryDate.Visible = False
            Me.DtpExpired.Visible = True
            Me.DtpExpired.Value = Today
            Me.lblLocation.Visible = False
            Me.lblWarehouse.Visible = False
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Visible = True
            Me.cboLocation.Items.Clear()
            Me.TxtWarehouse.Visible = True
            Me.lblBatchNo.Visible = False
            Me.cboBatchNo.Visible = True
            Me.cboBatchNo.TabStop = True
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Items.Clear()
            Me.cboSecondaryUnits.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            'data item salah bila:
            '0. jika no po tidak si spesifikasikan
            If Me.txtPurchaseOrderNo.Text = "" Then
                MsgBox("Please specify purchase order number", MsgBoxStyle.Information, "Information")
                Me.txtPurchaseOrderNo.Focus()
                Return False
            End If

            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.TxtCodeNo.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If


            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlled.Checked = True Then
                If Me.cboBatchNo.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.cboBatchNo.Focus()
                    Return False
                End If
            End If
            '4.a Untuk transaksi proses (in/back/out proses) no Pdo harus diisikan
            '4.b Untuk transaksi yang berkaitan dengan purchase (stock out purchase return dan stout in purchase)
            '    harus menyertakan no PO bila terhubungkan dengan software purchase
            '4.c Untuk transaksi yang berkaitan dengan sales (stock out sales dan stock in sales return) harrus
            '    menyertakan no DO bila terhubungkan dengan software sales
            '5. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlled.Checked = True Then
                If Me.txtPrimaryLocation.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, _
                    "Information")
                    Me.cboLocation.Focus()
                    Return False
                End If
            End If
            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlled.Checked = True Then
                If Me.DtpExpired.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
            End If
            '7. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnit.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.txtSecondaryQty.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.txtSecondaryQty.Focus()
                    Return False
                End If
            End If
            '8a. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlled.Checked = True And Me.chkExpiryDateControlled.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.cboBatchNo.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text And Me.cboBatchNo.Text <> "" And row.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColExpiredDate.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(Me.ColExpiredDate.Index).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.DtpExpired.Focus()
                            Return False
                        End If
                    End If
                Next
                '8.c Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam transaksi lainnya: stock in adjust, stock in purchase, 
                '    stock in sales return, stock tag, transaksi bundling (item before) 
                '    (untuk stock out tidak dicek karena semua harus masuk dulu sebelum keluar, jadi dimasukka ke dalam pengecekan
                '    adalah semua transaksi yang menjadi pintu masuk inventory)
                'msgbox("belum")
            End If

            '9. Bila setelah diinput jumlah product menjadi minus
            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNo.Text) = True Then
                Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal

                ''@rosadinnn
                ''@MUaT
                JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
                        Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value)
                JmlDGV = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
                        Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchDefault.Text)

                If JmlSetelahTambah < 0 Then
                    If Me.chkUseSecondaryUnit.Checked = True Then

                        ''@rosadinnn
                        ''@MUaT
                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + _
                                Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                                Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + _
                                CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, "Information")

                    End If
                    Me.TxtQtyperBatchMulti.Focus()
                    Return False
                End If
                '9.b. Cek stock pada tanggal yang bersangkutan
                JmlGudang = ObjSC.GetStock(Me.TxtPrimaryCodeNo.Text, Me.dtpDate.Value)
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
                            Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value)

                ''@rosadinnn
                ''@MUaT
                JmlDGV = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
                    Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchDefault.Text)
                If JmlSetelahTambah < 0 Then
                    If JmlGudang + JmlObjLama - JmlDGV > 0 Then
                        If Me.chkUseSecondaryUnit.Checked = True Then
                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + CStr(JmlGudang + JmlObjLama - JmlDGV) + " " + _
                            Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                            Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, _
                            "Information")
                        Else
                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + CStr(JmlGudang + JmlObjLama - JmlDGV) + " " + _
                            CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                            "Information")
                        End If
                    Else
                        MsgBox("Stock level on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + _
                        " is 0. Entering this data will cause negative value in your stock.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.TxtQtyperBatchMulti.Focus()
                    Return False
                End If
            End If

            If Me.ApakahAdadiDGVPR() = False Then
                Return False
            End If

            '10. Bila transaksi berkaitan dengan sales dan terinstall software sales maka ada batasan 
            '    berkaitan dengan jumlah item DO
            '11. Bila transaksi berkaitan dengan purchase dan terinstall software purchase maka ada batasan
            '    berkaitan dengan jumlah item DO
            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function

    Private Function ApakahAdadiDGVPR() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowItem, rowPR As DataGridViewRow
        Dim Cocok As Boolean
        Try
            For Each rowItem In Me.dgvItem.Rows
                Cocok = False
                For Each rowPR In Me.DGVPurchaseReturn.Rows
                    If Me.txtPurchaseOrderNo.Text = rowPR.Cells(Me.ColPONODGVPR.Index).Value And _
                    Me.TxtCodeNo.Text = rowPR.Cells(Me.colCodeNoDGVPR.Index).Value Then
                        Cocok = True
                        Exit For
                    End If
                Next
                If Cocok = False Then
                    If Me.TxtPurchaseReturnNumber.Text <> "" Then
                        MsgBox("Can not find Product: " + rowItem.Cells(Me.ColCodeNo.Index).Value + " in Purchase Return Item.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDGVItemNPRCocok() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowItem, rowGIT, rowItem2, rowGIT2 As DataGridViewRow
        Dim Ketemu As Boolean
        Try
            Me.DGVItem2.Rows.Clear()
            Me.DGVPurchaseReturn2.Rows.Clear()
            Me.DGVItemLebihKecilDariPR.Rows.Clear()
            Me.DGVItemLebihBesarDariPR.Rows.Clear()

            For Each rowGIT In Me.DGVPurchaseReturn.Rows
                Ketemu = False
                For Each rowItem In Me.dgvItem.Rows
                    If rowItem.Cells(Me.ColCodeNo.Index).Value = rowGIT.Cells(Me.colCodeNoDGVPR.Index).Value And _
                    rowItem.Cells(Me.ColPONoDONo.Index).Value = rowGIT.Cells(Me.ColPONODGVPR.Index).Value Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    MsgBox("There is one or more item you didn't input, please input every item on purchase return.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next
            Ketemu = False
            For Each rowItem In Me.dgvItem.Rows
                If Me.DGVItem2.Rows.Count = 0 Then

                    'Me.DGVItem2.Rows.Add("", rowItem.Cells(Me.ColPONoDONo.Index).Value, "", rowItem.Cells(Me.ColCodeNo.Index).Value, _
                    'rowItem.Cells(Me.ColDescription.Index).Value, rowItem.Cells(Me.ColQty.Index).Value, _
                    'rowItem.Cells(Me.ColUnit.Index).Value, rowItem.Cells(Me.ColSecondaryQty.Index).Value, _
                    'rowItem.Cells(Me.ColSecondaryUnit.Index).Value, rowItem.Cells(Me.ColBatch.Index).Value, _
                    'rowItem.Cells(Me.ColExpiryDate.Index).Value, rowItem.Cells(Me.ColLocation.Index).Value, _
                    'rowItem.Cells(Me.ColWareHouse.Index).Value, rowItem.Cells(Me.ColRemark.Index).Value, _
                    'rowItem.Cells(Me.ColPrimaryPOnoDoNo.Index).Value, rowItem.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'rowItem.Cells(Me.ColPrimaryLocation.Index).Value, rowItem.Cells(Me.ColPrimaryUnits.Index).Value, _
                    'rowItem.Cells(Me.ColPrimarySecondaryUnits.Index).Value, rowItem.Cells(Me.ColPrimaryPdO.Index).Value)
                    Me.DGVItem2.Rows.Add()
                    rowItem2 = Me.DGVItem2.Rows(Me.DGVItem2.Rows.Count - 1)
                    For i2 As Integer = 0 To Me.dgvItem.Columns.Count - 1
                        rowItem2.Cells(i2).Value = rowItem.Cells(i2).Value
                    Next

                Else
                    Ketemu = False
                    For Each rowItem2 In Me.DGVItem2.Rows
                        If rowItem2.Cells(Me.ColPONoDONo.Index).Value = rowItem.Cells(Me.ColPONoDONo.Index).Value And _
                           rowItem2.Cells(Me.ColCodeNo.Index).Value = rowItem.Cells(Me.ColCodeNo.Index).Value Then

                            ''@rosadinnn
                            ''@MUaT
                            rowItem2.Cells(Me.ColQtyMulti.Index).Value = CDbl(rowItem2.Cells(Me.ColQtyMulti.Index).Value) + CDbl(rowItem.Cells(Me.ColQtyMulti.Index).Value)

                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then

                        'Me.DGVItem2.Rows.Add("", rowItem.Cells(Me.ColPONoDONo.Index).Value, "", rowItem.Cells(Me.ColCodeNo.Index).Value, _
                        'rowItem.Cells(Me.ColDescription.Index).Value, rowItem.Cells(Me.ColQty.Index).Value, _
                        'rowItem.Cells(Me.ColUnit.Index).Value, rowItem.Cells(Me.ColSecondaryQty.Index).Value, _
                        'rowItem.Cells(Me.ColSecondaryUnit.Index).Value, rowItem.Cells(Me.ColBatch.Index).Value, _
                        'rowItem.Cells(Me.ColExpiryDate.Index).Value, rowItem.Cells(Me.ColLocation.Index).Value, _
                        'rowItem.Cells(Me.ColWareHouse.Index).Value, rowItem.Cells(Me.ColRemark.Index).Value, _
                        'rowItem.Cells(Me.ColPrimaryPOnoDoNo.Index).Value, rowItem.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        'rowItem.Cells(Me.ColPrimaryLocation.Index).Value, rowItem.Cells(Me.ColPrimaryUnits.Index).Value, _
                        'rowItem.Cells(Me.ColPrimarySecondaryUnits.Index).Value, rowItem.Cells(Me.ColPrimaryPdO.Index).Value)
                        Me.DGVItem2.Rows.Add()
                        rowItem2 = Me.DGVItem2.Rows(Me.DGVItem2.Rows.Count - 1)
                        For i2 As Integer = 0 To Me.dgvItem.Columns.Count - 1
                            rowItem2.Cells(i2).Value = rowItem.Cells(i2).Value
                        Next

                    End If
                End If
            Next
            For Each rowGIT In Me.DGVPurchaseReturn.Rows
                If Me.DGVPurchaseReturn2.Rows.Count = 0 Then

                    'Me.DGVPurchaseReturn2.Rows.Add("0", rowGIT.Cells(Me.colCodeNoDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.colDescDGVPR.Index).Value, rowGIT.Cells(Me.ColPONODGVPR.Index).Value, _
                    'rowGIT.Cells(Me.ColBatchNoDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.ColExpiryDateDGVPR.Index).Value, rowGIT.Cells(Me.ColQtyMasterDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.ColQtyPerMasterDGVPR.Index).Value, rowGIT.Cells(Me.ColMasterUnitsDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.colQTYDGVPR.Index).Value, rowGIT.Cells(Me.ColTotalQtyDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.colUnitDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.ColSecQtyDGVPR.Index).Value, rowGIT.Cells(Me.ColSecUnitDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.colRemarksDGVPR.Index).Value, rowGIT.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.ColPrimaryPODGVPR.Index).Value, rowGIT.Cells(Me.ColPrimaryMasterUnitsDGVPR.Index).Value, _
                    'rowGIT.Cells(Me.ColPrimaryUnitDGVPR.Index).Value, rowGIT.Cells(Me.ColPrimarySecUnitDGVPR.Index).Value)
                    Me.DGVPurchaseReturn2.Rows.Add()
                    rowGIT2 = Me.DGVPurchaseReturn2.Rows(Me.DGVPurchaseReturn2.Rows.Count - 1)
                    For i2 As Integer = 0 To Me.DGVPurchaseReturn2.Columns.Count - 1
                        rowGIT2.Cells(i2).Value = rowGIT.Cells(i2).Value
                    Next

                Else
                    Ketemu = False
                    For Each rowGIT2 In Me.DGVPurchaseReturn2.Rows
                        If rowGIT2.Cells(Me.ColPONODGVPR.Index).Value = rowGIT.Cells(Me.ColPONODGVPR.Index).Value And _
                           rowGIT2.Cells(Me.colCodeNoDGVPR.Index).Value = rowGIT.Cells(Me.colCodeNoDGVPR.Index).Value Then

                            ''@rosadinnn
                            ''@MUaT
                            rowGIT2.Cells(Me.ColQtyMulti_PR.Index).Value = CDbl(rowGIT2.Cells(Me.ColQtyMulti_PR.Index).Value) + CDbl(rowGIT.Cells(Me.ColQtyMulti_PR.Index).Value)

                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then

                        'Me.DGVPurchaseReturn2.Rows.Add("0", rowGIT.Cells(Me.colCodeNoDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.colDescDGVPR.Index).Value, rowGIT.Cells(Me.ColPONODGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.ColBatchNoDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.ColExpiryDateDGVPR.Index).Value, rowGIT.Cells(Me.ColQtyMasterDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.ColQtyPerMasterDGVPR.Index).Value, rowGIT.Cells(Me.ColMasterUnitsDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.colQTYDGVPR.Index).Value, rowGIT.Cells(Me.ColTotalQtyDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.colUnitDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.ColSecQtyDGVPR.Index).Value, rowGIT.Cells(Me.ColSecUnitDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.colRemarksDGVPR.Index).Value, rowGIT.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.ColPrimaryPODGVPR.Index).Value, rowGIT.Cells(Me.ColPrimaryMasterUnitsDGVPR.Index).Value, _
                        '   rowGIT.Cells(Me.ColPrimaryUnitDGVPR.Index).Value, rowGIT.Cells(Me.ColPrimarySecUnitDGVPR.Index).Value)
                        Me.DGVPurchaseReturn2.Rows.Add()
                        rowGIT2 = Me.DGVPurchaseReturn2.Rows(Me.DGVPurchaseReturn2.Rows.Count - 1)
                        For i2 As Integer = 0 To Me.DGVPurchaseReturn2.Columns.Count - 1
                            rowGIT2.Cells(i2).Value = rowGIT.Cells(i2).Value
                        Next

                    End If
                End If
            Next

            ''@rosadinnn
            ''@MUaT
            For Each rowItem2 In Me.DGVItem2.Rows
                For Each rowGIT2 In Me.DGVPurchaseReturn2.Rows
                    If rowItem2.Cells(Me.ColPONoDONo.Index).Value = rowGIT2.Cells(Me.ColPONODGVPR.Index).Value And _
                       rowItem2.Cells(Me.ColQtyMulti.Index).Value < rowGIT2.Cells(Me.ColQtyMulti_PR.Index).Value And _
                       rowItem2.Cells(Me.ColCodeNo.Index).Value = rowGIT2.Cells(Me.colCodeNoDGVPR.Index).Value Then

                        Me.DGVItemLebihKecilDariPR.Rows.Add(rowItem2.Cells(Me.ColPONoDONo.Index).Value, rowItem2.Cells(Me.ColCodeNo.Index).Value, _
                        rowItem2.Cells(Me.ColBatch.Index).Value, rowItem2.Cells(Me.ColQtyMulti.Index).Value)

                    ElseIf rowItem2.Cells(Me.ColPONoDONo.Index).Value = rowGIT2.Cells(Me.ColPONODGVPR.Index).Value And _
                       rowItem2.Cells(Me.ColQtyMulti.Index).Value > rowGIT2.Cells(Me.ColQtyMulti_PR.Index).Value And _
                       rowItem2.Cells(Me.ColCodeNo.Index).Value = rowGIT2.Cells(Me.colCodeNoDGVPR.Index).Value Then

                        Me.DGVItemLebihBesarDariPR.Rows.Add(rowItem2.Cells(Me.ColPONoDONo.Index).Value, rowItem2.Cells(Me.ColCodeNo.Index).Value, _
                        rowItem2.Cells(Me.ColBatch.Index).Value, rowItem2.Cells(Me.ColQtyMulti.Index).Value)

                    End If
                Next
            Next
            If Me.DGVItemLebihBesarDariPR.Rows.Count > 0 Then
                Dim i As String = ""
                For Each rowItem In Me.DGVItemLebihBesarDariPR.Rows
                    If i = "" Then
                        i = "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is > purchase return qty"
                    Else
                        i = i + ", " + "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is > purchase return qty"
                    End If
                Next
                i = i + "."
                MsgBox("Can not save because " + i, MsgBoxStyle.OkOnly, "Confirm")
                Return False
            End If
            If Me.DGVItemLebihKecilDariPR.Rows.Count > 0 Then
                Dim i As String = ""
                For Each rowItem In Me.DGVItemLebihKecilDariPR.Rows
                    If i = "" Then
                        i = "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is <than purchase return qty"
                    Else
                        i = i + ", " + "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is < purchase return qty"
                    End If
                Next
                i = i + "."
                MsgBox("Can not save because " + i, MsgBoxStyle.OkOnly, "Confirm")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub UpdateDataItemPurchaseReturn()
    '    Try
    '        Me.DGVPurchaseReturn2.Rows.Clear()
    '        Me.DGVItem2.Rows.Clear()
    '        Dim rowItem, rowItem2 As DataGridViewRow
    '        Dim Ketemu As Boolean = False

    '        For Each rowItem In Me.dgvItem.Rows
    '            If Me.DGVItem2.Rows.Count = 0 Then
    '                Me.DGVItem2.Rows.Add("", rowItem.Cells(Me.ColPONoDONo.Index).Value, "", rowItem.Cells(Me.ColCodeNo.Index).Value, _
    '                rowItem.Cells(Me.ColDescription.Index).Value, rowItem.Cells(Me.ColQty.Index).Value, _
    '                rowItem.Cells(Me.ColUnit.Index).Value, rowItem.Cells(Me.ColSecondaryQty.Index).Value, _
    '                rowItem.Cells(Me.ColSecondaryUnit.Index).Value, rowItem.Cells(Me.ColBatch.Index).Value, _
    '                rowItem.Cells(Me.ColExpiryDate.Index).Value, rowItem.Cells(Me.ColLocation.Index).Value, _
    '                rowItem.Cells(Me.ColWareHouse.Index).Value, rowItem.Cells(Me.ColRemark.Index).Value, _
    '                rowItem.Cells(Me.ColPrimaryPOnoDoNo.Index).Value, rowItem.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '                rowItem.Cells(Me.ColPrimaryLocation.Index).Value, rowItem.Cells(Me.ColPrimaryUnits.Index).Value, _
    '                rowItem.Cells(Me.ColPrimarySecondaryUnits.Index).Value, rowItem.Cells(Me.ColPrimaryPdO.Index).Value)
    '            Else
    '                Ketemu = False
    '                For Each rowItem2 In Me.DGVItem2.Rows
    '                    If rowItem2.Cells(Me.ColPONoDONo.Index).Value = rowItem.Cells(Me.ColPONoDONo.Index).Value And _
    '                       rowItem2.Cells(Me.ColCodeNo.Index).Value = rowItem.Cells(Me.ColCodeNo.Index).Value And _
    '                       rowItem2.Cells(Me.ColBatch.Index).Value = rowItem.Cells(Me.ColBatch.Index).Value Then
    '                        rowItem2.Cells(Me.ColQty.Index).Value = CDbl(rowItem2.Cells(Me.ColQty.Index).Value) + CDbl(rowItem.Cells(Me.ColQty.Index).Value)
    '                        Ketemu = True
    '                        Exit For
    '                    End If
    '                Next
    '                If Ketemu = False Then
    '                    Me.DGVItem2.Rows.Add("", rowItem.Cells(Me.ColPONoDONo.Index).Value, "", rowItem.Cells(Me.ColCodeNo.Index).Value, _
    '                    rowItem.Cells(Me.ColDescription.Index).Value, rowItem.Cells(Me.ColQty.Index).Value, _
    '                    rowItem.Cells(Me.ColUnit.Index).Value, rowItem.Cells(Me.ColSecondaryQty.Index).Value, _
    '                    rowItem.Cells(Me.ColSecondaryUnit.Index).Value, rowItem.Cells(Me.ColBatch.Index).Value, _
    '                    rowItem.Cells(Me.ColExpiryDate.Index).Value, rowItem.Cells(Me.ColLocation.Index).Value, _
    '                    rowItem.Cells(Me.ColWareHouse.Index).Value, rowItem.Cells(Me.ColRemark.Index).Value, _
    '                    rowItem.Cells(Me.ColPrimaryPOnoDoNo.Index).Value, rowItem.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '                    rowItem.Cells(Me.ColPrimaryLocation.Index).Value, rowItem.Cells(Me.ColPrimaryUnits.Index).Value, _
    '                    rowItem.Cells(Me.ColPrimarySecondaryUnits.Index).Value, rowItem.Cells(Me.ColPrimaryPdO.Index).Value)
    '                End If
    '            End If
    '        Next

    '        For Each rowItem2 In Me.DGVItem2.Rows
    '            Me.DGVPurchaseReturn2.Rows.Add("0", rowItem2.Cells(Me.ColCodeNo.Index).Value, _
    '            rowItem2.Cells(Me.ColDescription.Index).Value, rowItem2.Cells(Me.ColPONoDONo.Index).Value, _
    '            rowItem2.Cells(Me.ColBatch.Index).Value, _
    '            rowItem2.Cells(Me.ColExpiryDate.Index).Value, rowItem2.Cells(Me.ColQty.Index).Value, _
    '            rowItem2.Cells(Me.ColQty.Index).Value, rowItem2.Cells(Me.ColUnit.Index).Value, _
    '            rowItem2.Cells(Me.ColQty.Index).Value, rowItem2.Cells(Me.ColQty.Index).Value, _
    '            rowItem2.Cells(Me.ColUnit.Index).Value, _
    '            rowItem2.Cells(Me.ColSecondaryQty.Index).Value, rowItem2.Cells(Me.ColSecondaryUnit.Index).Value, _
    '            rowItem2.Cells(Me.ColRemark.Index).Value, rowItem2.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '            rowItem2.Cells(Me.ColPrimaryPOnoDoNo.Index).Value, rowItem2.Cells(Me.ColPrimaryUnits.Index).Value, _
    '            rowItem2.Cells(Me.ColPrimaryUnits.Index).Value, rowItem2.Cells(Me.ColPrimarySecondaryUnits.Index).Value)
    '        Next

    '        Dim objPRItem As New ClsTransaksiPurchaseReturnItem
    '        objPRItem.Simpan(Me.DGVPurchaseReturn2, Me.TxtPrimaryPurchaseReturn.Text)
    '        objPRItem = Nothing
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal BatchNoControlled As Boolean, ByVal LocationControlled As Boolean, ByVal UseSecondaryUnit As Boolean) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    OK = True
                    'cek secondary qty
                    If UseSecondaryUnit = True And OK = True Then
                        If row.Cells(Me.ColSecondaryQty.Index).Value <> Secondary_Qty Then
                            OK = False
                        End If
                    End If
                    'cek location
                    If LocationControlled = True And OK = True Then
                        If vPrimaryLocation <> row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If BatchNoControlled = True Then
                        If vBatchNo <> row.Cells(Me.ColBatch.Index).Value Then
                            OK = False
                        End If
                    End If
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    If OK = True Then
                        Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal BatchNoControlled As Boolean, ByVal LocationControlled As Boolean, ByVal UseSecondaryUnit As Boolean) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    OK = True
                    'cek secondary qty
                    If UseSecondaryUnit = True And OK = True Then
                        If row.Cells(Me.ColSecondaryQty.Index).Value <> Secondary_Qty Then
                            OK = False
                        End If
                    End If
                    'cek location
                    If LocationControlled = True And OK = True Then
                        If vPrimaryLocation <> row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If BatchNoControlled = True Then
                        If vBatchNo <> row.Cells(Me.ColBatch.Index).Value Then
                            OK = False
                        End If
                    End If
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    If OK = True Then
                        Jml = Jml + CDec(row.Cells(Me.ColQtyMulti.Index).Value)
                    End If
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
        Dim TanggalClosingMax As Date = #1/1/1950#

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("stock out number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'harus jelas sales ke siapa
            If Me.TxtPrimaryPihakAcuan.Text = "-1" Then
                MsgBox("Please specify supplier's code", MsgBoxStyle.Information, "Information")
                Me.TxtKodePihakAcuan.Focus()
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'berkaitan denga tanggal closing max
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.objStockOut.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objStockOut.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtStockInNo.Text = Me.objStockOut.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objStockOut.TanggalTransaksi
                        Me.TxtStockInNo.Text = Me.objStockOut.NoTransaksi
                        Return False
                    End If
                End If
            End If
            If Me.ApakahDGVItemNPRCocok() = False Then
                Return False
            End If

            'cek tanggal transaksi, harus berada dalam periode yang sama dengan tanggal purchase return-nya
            If Format(Me.ObjPR.GetTanggalTransaksi(Me.TxtPrimaryPurchaseReturn.Text), "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                MsgBox("The purchase return must be date in the same period as the stock out purchase return.", _
                MsgBoxStyle.Information, "Information")
                Me.dtpDate.Focus()
                Return False
            End If

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_OUT_PURCHASE_RETURN, Me.TBLDetailLocation) = False Then
                    Return False
                End If
            End If

            Dim ObjCust As New CustomSMESS.ClsTransaksiStockInOut
            Dim rd As DataGridViewRow
            Dim TblDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
            TblDummy = Me.TblItemCustom.Copy
            'If TblDummy.Rows.Count > 0 Then
            For i As Integer = 0 To Me.dgvItem.RowCount - 1
                rd = Me.dgvItem.Rows(i)
                With rd
                    Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    Dim PrimaryPOSO As Double = .Cells(Me.ColPrimaryPOnoDoNo.Index).Value
                    Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                    Dim Batch As String = .Cells(Me.ColBatch.Index).Value
                    Dim QtyDef As Decimal = .Cells(Me.ColQty.Index).Value
                    Dim QtyMulti As Decimal = .Cells(Me.ColQtyMulti.Index).Value
                    Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                    If ObjCust.CompareQtySMESS_Custom(PrimaryProduct, PrimaryPOSO, PrimaryLocation, Batch, QtyDef, QtyMulti, TblDummy) = False Then
                        If String.IsNullOrEmpty(Batch) = False Then
                            ProdCode &= " batch: " & Batch
                        End If
                        MsgBox("Inconsisten data, please cek product: " & ProdCode & "between main form and custom form.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        rd.Selected = True
                        Return False
                    End If
                End With
            Next
            'End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub LoadLocation()
        Dim RowDetailLocation As DseUserLocationPrivilegeSetting.tblmaster_storage_areaRow
        Dim rowDGV As DataGridViewRow
        Dim ListPrimary As New List(Of Double)

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryLocation.Index).Value <> -1 Then
                    If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryLocation.Index).Value) = False Then
                        RowDetailLocation = Me.TBLDetailLocation.Newtblmaster_storage_areaRow
                        With RowDetailLocation
                            .Primary = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value
                            .Storage_Area_Code = rowDGV.Cells(Me.ColLocation.Index).Value
                            .Primary_Warehouse = Me.ObjLocation.GetPrimaryWarehouse(.Primary)
                        End With
                        Me.TBLDetailLocation.Rows.Add(RowDetailLocation)
                        ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryLocation.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objStockOut
                .NoTransaksi = Me.TxtStockInNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Remarks = Me.TxtRemark.Text
                .PrimarySupplier = Me.TxtPrimaryPihakAcuan.Text
                .PrimaryPOPurchase = Me.TxtPrimaryPurchaseReturn.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim PrimaryGRN As Double = -1
        Dim PrimaryStockInPurchase As Double = -1

        Try
            With Me.objStockOut
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.TxtPrimaryPurchaseReturn.Text = .PrimaryPOPurchase
                Me.TxtPurchaseReturnNumber.Text = Me.ObjPR.GetNoTransaksi(Me.TxtPrimaryPurchaseReturn.Text)
                Me.TxtPrimaryPihakAcuan.Text = .PrimarySupplier
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.TxtPrimaryPihakAcuan.Text = .PrimarySupplier
                Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtKodePihakAcuan.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimaryPihakAcuan.Text)

                PrimaryGRN = Me.ObjPR.GetPrimaryGRN(Me.TxtPrimaryPurchaseReturn.Text)
                PrimaryStockInPurchase = Me.ObjGRN.GetPrimaryStockInPurchase(PrimaryGRN)
                Me.TxtSupplierDONumber.Text = Me.objStockOut.GetSupplierDONumber(PrimaryStockInPurchase)
            End With

            Me.FormatDGV_Multi()
            'digit grouping
            Dim row As DataGridViewRow

            ''@rosadinnn
            ''@MUaT
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyMulti.Index).Value, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitMulti.Index).Value))
                row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    row.Cells(Me.ColSecondaryQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecondaryQty.Index).Value)
                End If
            Next

            Me.IsikanDGVpurchasereturn(Me.DGVPurchaseReturn, Me.TxtPrimaryPurchaseReturn.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtSupplierDONumber.CharacterCasing = CharacterCasing.Normal

            Me.baru = True
            Me.NonAktifkanDetail()
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
            Me.KosongkanField()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVPurchaseReturn.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()

                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockOut.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                            'Me.UpdateDataItemPurchaseReturn()
                            'update data primary stock out purchase return pada transaksi purchase return 
                            Me.SalinKeField()
                            'Me.ObjPR.SetPrimaryStockOutPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text, Me.TxtPrimaryStockIn.Text)
                            'simpan data jurnal purchase return -> sudah ada di kelas stock in out
                            'Me.objStockOut.TransferJurnal(Me.objStockOut.PrimarySupplier, Me.objStockOut.Primary, -1)
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockOut.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                            Me.SalinKeField()
                            'update data purchase return yang ditunjuk -> dipindah ke kelas
                            'Me.ObjPR.SetPrimaryStockOutPurchaseReturn_Minus_Satu(Me.TxtPrimaryStockIn.Text)
                            'Me.ObjPR.SetPrimaryStockOutPurchaseReturn(Me.TxtPrimaryPurchaseReturn.Text, Me.TxtPrimaryStockIn.Text)
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            'Me.UpdateDataItemPurchaseReturn()
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
            Me.TxtCodeNo.Text = "" And Me.cboBatchNo.Text = "" And Me.cboLocation.Text = "" And _
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objStockOut.CariDetail(Me.TxtPrimaryStockIn.Text, Me.dgvItem, Me.TblItemCustom)
                        Me.SalinKeField()
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
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.objStockOut.ResetObject(Me.dgvItem)
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
            Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.TxtPurchaseReturnNumber.ReadOnly = False
            Me.TxtPurchaseReturnNumber.BackColor = Me.txtPrimaryCategory.BackColor
            Me.btnFindPurchaseReturn.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiStockInOut_Purchase
        Try
            hj.JenisTransaksi = Me.TipeTransaksi
            hj.ShowDialog()
            If Not hj.PrimaryTransaksiStockInOut = -1 Then
                If Me.objStockOut.CariDetail(hj.PrimaryTransaksiStockInOut, Me.dgvItem, Me.TblItemCustom) = True Then
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
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                'jika stock out berasal dari quick pr, maka hanya bisa diedit menggunakan transaksi quick pr
                If Me.ObjPR.ApakahDariQuickPR(Me.TxtPrimaryStockIn.Text) = True Then
                    MsgBox("Please use quick purchase form to edit this data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objStockOut.ApakahBolehEdit(Me.UserCode, True, True) Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                    Me.TxtPurchaseReturnNumber.ReadOnly = False
                    Me.TxtPurchaseReturnNumber.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.btnFindPurchaseReturn.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                'jika stock out berasal dari quick pr, maka hanya bisa diedit menggunakan transaksi quick pr
                If Me.ObjPR.ApakahDariQuickPR(Me.TxtPrimaryStockIn.Text) = True Then
                    MsgBox("Please use quick purchase form to delete this data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                'tidak cek stock minus ketika cek apakah boleh hapus, karena akan di cek tepat sebelum di hapus,
                'pengecekan ada dalam fungsi hapus di kelas, sehingga memperhitungkan antrian
                If Me.objStockOut.ApakahBolehHapus(Me.UserCode, True, True, False, False) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockOut.Hapus(Me.TxtPrimaryStockIn.Text, Me.UserCode, Me.TblItemCustom) = True Then
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
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
                            If Me.objStockOut.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom) = True Then
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
                            If Me.objStockOut.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) = True Then
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

    Private Sub IsiCboItem(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Try
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNo, Me.TBL)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As dseTransaksiStockInOut.tbltransaksi_stockinout_item1Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockOutOutItem.IsikanTBLBatchNo(Me.TBL4, vPrimaryProduct, vPrimaryTransaksi)
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
            If Me.chkExpiryDateControlled.Checked = False Then
                Me.TBL.DefaultView.Sort = "[Lot No] ASC"
            Else
                Me.TBL.DefaultView.Sort = "[Expiry Date] ASC"
            End If
            Me.cboBatchNo.DataSource = Me.TBL.DefaultView
            Me.cboBatchNo.ValueMember = "Lot No"
            Me.cboBatchNo.DisplayMember = "Lot No"
            Me.cboBatchNo.Refresh()
            'pilih salah satu
            'If Me.cboBatchNo.Items.Count > 0 Then
            'Me.cboBatchNo.SelectedIndex = -1
            'Me.cboBatchNo.SelectedIndex = 0
            'End If
            If Me.chkLocationControlled.Checked = False And _
            Me.chkUseSecondaryUnit.Checked = True And Me.cboBatchNo.Text <> "" Then
                Me.IsiCboSecondary(vPrimaryProduct, vPrimaryTransaksi, Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Double)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockInOut.tbltransaksi_stockinout_item1Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColBatch.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = True Then
                            ED = row.Cells(Me.ColExpiredDate.Index).Value
                        End If
                        Me.TBL4.Rows.Add(CStr(row.Cells(Me.ColBatch.Index).Value), CDbl(row.Cells(Me.ColSecondaryQty.Index).Value), ED)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboLokasi(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboLocation(vPrimaryProduct, Me.cboLocation, _
            Me.TBL2, vBatchNo)
            Dim row1 As dseTransaksiStockInOut.tbltransaksi_stockinout_item2Row
            Dim row2 As dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockOutOutItem.IsikanTBLLocation(Me.TBL5, vPrimaryProduct, vPrimaryTransaksi, vBatchNo)
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
            Me.cboLocation.DataSource = Me.TBL2.DefaultView
            Me.cboLocation.ValueMember = "Location"
            Me.cboLocation.DisplayMember = "Location"
            Me.cboLocation.Refresh()
            'pilih salah satu
            'If Me.cboLocation.Items.Count > 0 Then
            'Me.cboLocation.SelectedIndex = -1
            'Me.cboLocation.SelectedIndex = 0
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLLocation(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockInOut.tbltransaksi_stockinout_item2Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And CStr(row.Cells(Me.ColBatch.Index).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL5.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColBatch.Index).Value) And row1.Location = row.Cells(Me.ColLocation.Index).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL5.Rows.Add(CStr(row.Cells(Me.ColBatch.Index).Value), row.Cells(Me.ColLocation.Index).Value, CDbl(row.Cells(Me.ColSecondaryQty.Index).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboSecondary(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, _
    ByVal vBatchNo As String, ByVal vPrimaryLokasi As Double)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboSecondaryUnit(vPrimaryProduct, Me.cboSecondaryUnits, _
            Me.TBL3, vBatchNo, vPrimaryLokasi)
            Dim row1 As dseTransaksiStockInOut.tbltransaksi_stockinout_item3Row
            Dim row2 As dseMasterProduct.tblmaster_productqty2Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockOutOutItem.IsikanTBLSecondaryUnit(Me.TBL6, vPrimaryProduct, _
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
            Me.cboSecondaryUnits.DataSource = Me.TBL3.DefaultView
            Me.cboSecondaryUnits.ValueMember = "Qty Secondary"
            Me.cboSecondaryUnits.DisplayMember = "Qty Secondary"
            Me.cboSecondaryUnits.Refresh()
            'pilih salah satu
            If Me.cboSecondaryUnits.Items.Count > 0 Then
                Me.cboSecondaryUnits.SelectedIndex = -1
                Me.cboSecondaryUnits.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLSecondary(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockInOut.tbltransaksi_stockinout_item3Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And CStr(row.Cells(Me.ColBatch.Index).Value) = vBatchNo And _
                row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL6.Rows
                        If row1.Secondary_Qty = CDbl(row.Cells(Me.ColSecondaryQty.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = True Then
                            ED = row.Cells(Me.ColExpiredDate.Index).Value
                        End If
                        Me.TBL6.Rows.Add(CDbl(row.Cells(Me.ColSecondaryQty.Index).Value), CStr(row.Cells(Me.ColQty.Index).Value), _
                        CStr(row.Cells(Me.ColBatch.Index).Value), CStr(row.Cells(Me.ColLocation.Index).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
            Me.chkPdoControlled, Me.txtPrimaryUnitsDefault, Me.txtPrimarySecondaryUnits)

            Me.cboBatchNo.DataSource = Nothing
            Me.cboLocation.DataSource = Nothing
            Me.cboSecondaryUnits.DataSource = Nothing

            Me.txtPrimaryPDO.Text = "-1"
            If Me.chkBatchNoControlled.Checked = True Then
                Me.cboBatchNo.TabStop = True
                Me.lblBatchNo.Visible = False
                Me.cboBatchNo.Visible = True
                'isikan data batch no dari basis data
                Me.IsiCboItem(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text)
            Else
                Me.cboBatchNo.TabStop = False
                Me.lblBatchNo.Visible = True
                Me.cboBatchNo.Visible = False
                Me.cboBatchNo.Items.Clear()
            End If
            If Me.chkLocationControlled.Checked = False Then
                Me.txtPrimaryLocation.Text = "-1"
                Me.cboLocation.Items.Clear()
                Me.TxtWarehouse.Clear()
                Me.lblLocation.Visible = True
                Me.lblWarehouse.Visible = True
                Me.cboLocation.Visible = False
                Me.TxtWarehouse.Visible = False
            Else
                Me.lblLocation.Visible = False
                Me.lblWarehouse.Visible = False
                Me.cboLocation.Visible = True
                Me.TxtWarehouse.Visible = True
                If Me.chkLocationControlled.Checked = True Then
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text, Me.cboBatchNo.Text)
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
            If Me.chkUseSecondaryUnit.Checked = False Then
                Me.lblSecondaryUnits.Visible = False

                ''@rosadinnn
                ''@MUaT
                Me.lblUnitsMulti.Left = Me.BtnInput.Left

                Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
                Me.cboSecondaryUnits.Visible = False
                Me.cboSecondaryUnits.DataSource = Nothing
                Me.cboSecondaryUnits.Items.Clear()
            Else
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True

                ''@rosadinnn
                ''@MUaT
                'Me.lblUnitsMulti.Left = Me.lblSecQty.Left - 37

                Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
                Me.cboSecondaryUnits.Visible = True
                If Me.chkLocationControlled.Checked = False Then
                    'isikan data secondary jika perlu
                    Me.IsiCboSecondary(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text, _
                    Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                End If
            End If
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then

                ''@rosadinnn
                ''@MUaT
                'Me.lblUnitsMulti.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)

                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            Else
                Me.lblUnitsMulti.Text = "Units"

                ''@rosadinnn
                ''@MUaT
                Me.lblUnitsDefault.Text = "Units"
                Me.LblQty.Text = "Qty"

                Me.lblSecondaryUnits.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.cboBatchNo.Visible = True Then
                        Me.cboBatchNo.Focus()
                    ElseIf Me.cboLocation.Visible = True Then
                        Me.cboLocation.Focus()
                    Else
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                    Me.SetPONumber()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

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
                            Me.TxtQtyperBatchMulti.Focus()
                        End If
                        Me.SetPONumber()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.cboBatchNo.Visible = True Then
                        Me.cboBatchNo.Focus()
                    ElseIf Me.cboLocation.Visible = True Then
                        Me.cboLocation.Focus()
                    Else
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                    Me.SetPONumber()
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

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
                            Me.TxtQtyperBatchMulti.Focus()
                        End If
                        Me.SetPONumber()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
            End If
            Me.SetGboItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetPONumber()
        Dim row As DataGridViewRow
        Dim Jml As Integer = 0
        Dim PO_Number As String = ""
        Dim PrimaryPO As Double = -1
        Try
            For Each row In Me.DGVPurchaseReturn.Rows
                If row.Cells(Me.colCodeNoDGVPR.Index).Value.ToString = Me.TxtCodeNo.Text Then
                    PO_Number = row.Cells(Me.ColPONODGVPR.Index).Value
                    PrimaryPO = row.Cells(Me.ColPrimaryPODGVPR.Index).Value
                    Jml += 1
                End If
            Next

            If Jml > 0 Then
                Me.txtPrimaryPO.Text = PrimaryPO
                Me.txtPurchaseOrderNo.Text = PO_Number
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

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
                    Me.TxtQtyperBatchMulti.Focus()
                End If
                Me.SetPONumber()
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan index baris dalam dgv yang sesuai dengan parameter yang diberikan
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="vPrimaryLocation"></param>
    ''' <param name="BatchNo"></param>
    ''' <param name="SecondaryQty"></param>
    ''' <param name="PrimaryPdO"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex_Default(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) Then
                            Return row.Index
                        End If
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetIndex_Multi(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
            ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                                row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                                row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = _
                                PrimaryPdO And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) Then
                            Return row.Index
                        End If
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim RowDgv As DataGridViewRow
        Try 
            If Me.ApakahDataItemBenar_Multi = False Then
                Exit Sub
            End If

            If Me.ApakahDataItemBenar = True Then
                Dim Index As Integer = Me.GetIndex_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, _
                        Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(Me.TxtQtyperBatchMulti.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.SetQtyDefaultOnDgv(Index)

                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If

                    ''@rosadinnn
                    ''@MUaT
                    'Me.dgvItem.Rows.Add("", Me.txtPurchaseOrderNo.Text, "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatchMulti.Text, _
                    'Me.lblUnitsMulti.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.cboBatchNo.Text, _
                    'ExpiredDateStr, Me.cboLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnitsDefault.Text, _
                    'Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)
                    Me.dgvItem.Rows.Add()
                    RowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPONoDONo.Index).Value = Me.txtPurchaseOrderNo.Text
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                        .Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                        .Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text
                        .Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColLocation.Index).Value = Me.cboLocation.Text
                        .Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryPOnoDoNo.Index).Value = Me.txtPrimaryPO.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                        .Cells(Me.ColPrimaryPdO.Index).Value = Me.txtPrimaryPDO.Text

                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                        .Cells(Me.ColUnitCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                        If Me.txtPrimaryUnitsMulti.Text <> Me.txtPrimaryUnitsDefault.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                                    ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) _
                                    & " " & Me.lblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                                    Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) & " " & Me.lblUnitsDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        .Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                        .Cells(Me.ColPrimaryUnitMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                    End With

                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    'setelah menambahkan baris baru, data purchase return tidak boleh diubah
                    Me.TxtPurchaseReturnNumber.ReadOnly = True
                    Me.TxtPurchaseReturnNumber.BackColor = Me.TxtReadOnly.BackColor
                    Me.btnFindPurchaseReturn.Enabled = False
                End If
                'kosongkan gbo item
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
                'format dgv
                Me.FormatDGV_Multi()
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

    Private Sub FormatDGV_Default()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    row.Cells(Me.ColExpiredDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiredDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiredDate.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiredDate.Index).Value = Format(CDate(row.Cells(Me.ColExpiredDate.Index).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecondaryQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.ColSecondaryQty.Index).Value = "-1" Then
                    row.Cells(Me.ColSecondaryQty.Index).Value = "-"
                    row.Cells(Me.ColSecondaryUnit.Index).Value = "-"
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(Me.ColPrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.ColLocation.Index).Value = "-"
                    row.Cells(Me.ColWareHouse.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatch.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(Me.ColSecondaryQty.Index).Visible = False
                Me.dgvItem.Columns(Me.ColSecondaryUnit.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColSecondaryQty.Index).Visible = True
                Me.dgvItem.Columns(Me.ColSecondaryUnit.Index).Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = False
                Me.dgvItem.Columns(Me.ColWareHouse.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = True
                Me.dgvItem.Columns(Me.ColWareHouse.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColBatch.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColBatch.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiredDate.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiredDate.Index).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV_Multi()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Dim IsMultiUnit As Boolean = False

        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    row.Cells(Me.ColExpiredDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiredDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiredDate.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiredDate.Index).Value = Format(CDate(row.Cells(Me.ColExpiredDate.Index).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecondaryQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.ColSecondaryQty.Index).Value = "-1" Then
                    'row.Cells(Me.ColSecondaryQty.Index).Value = "-"
                    row.Cells(Me.ColSecondaryUnit.Index).Value = "-"
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(Me.ColPrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.ColLocation.Index).Value = "-"
                    row.Cells(Me.ColWareHouse.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatch.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(Me.ColSecondaryQty.Index).Visible = False
                Me.dgvItem.Columns(Me.ColSecondaryUnit.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColSecondaryQty.Index).Visible = True
                Me.dgvItem.Columns(Me.ColSecondaryUnit.Index).Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = False
                Me.dgvItem.Columns(Me.ColWareHouse.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = True
                Me.dgvItem.Columns(Me.ColWareHouse.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColBatch.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColBatch.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiredDate.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiredDate.Index).Visible = True
            End If

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
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV_PR()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            'yang diformat dgv-nya:
            For Each row In Me.DGVPurchaseReturn.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiryDateDGVPR.Index).Value) = False Then
                    row.Cells(Me.ColExpiryDateDGVPR.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiryDateDGVPR.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiryDateDGVPR.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiryDateDGVPR.Index).Value = Format(CDate(row.Cells(Me.ColExpiryDateDGVPR.Index).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(Me.ColSecQtyDGVPR.Index).Value) = True Then
                    If row.Cells(Me.ColSecQtyDGVPR.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.ColSecQtyDGVPR.Index).Value = "-1" Then
                    row.Cells(Me.ColSecQtyDGVPR.Index).Value = "-"
                    row.Cells(Me.ColSecQtyDGVPR.Index).Value = "-"
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatchNoDGVPR.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem PR tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.DGVPurchaseReturn.Columns(Me.ColSecQtyDGVPR.Index).Visible = False
                Me.DGVPurchaseReturn.Columns(Me.ColSecUnitDGVPR.Index).Visible = False
            Else
                Me.DGVPurchaseReturn.Columns(Me.ColSecQtyDGVPR.Index).Visible = True
                Me.DGVPurchaseReturn.Columns(Me.ColSecUnitDGVPR.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.DGVPurchaseReturn.Columns(Me.ColBatchNoDGVPR.Index).Visible = False
            Else
                Me.DGVPurchaseReturn.Columns(Me.ColBatchNoDGVPR.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.DGVPurchaseReturn.Columns(Me.ColExpiryDateDGVPR.Index).Visible = False
            Else
                Me.DGVPurchaseReturn.Columns(Me.ColExpiryDateDGVPR.Index).Visible = True
            End If
            Me.DGVPurchaseReturn.Columns(Me.colQTYDGVPR.Index).Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
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

    Private Sub txtSecondaryQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQty, sender, e)
    End Sub

    Private Sub txtSecondaryQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQty, sender, e)
    End Sub

    'Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
    '    Try
    '        'yang bisa diedit adalah:
    '        '1. Qty
    '        '2. Secondary Qty
    '        '3. Batch No
    '        '4. Location Code
    '        '5. Remarks
    '        'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
    '        If Me.TxtRemark.Enabled = True Then
    '            If e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 Or e.ColumnIndex = 11 _
    '            Or e.ColumnIndex = 13 Then
    '                'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
    '                'masing-masing parameter yang akan diedit dikontrol
    '                If e.ColumnIndex = 7 Then
    '                    'cek secondary qty
    '                    If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
    '                        e.Cancel = True
    '                        Exit Sub
    '                    End If
    '                ElseIf e.ColumnIndex = 9 Then
    '                    'cek batch no
    '                    If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItem.Item(15, e.RowIndex).Value) = False Then
    '                        e.Cancel = True
    '                        Exit Sub
    '                    End If
    '                ElseIf e.ColumnIndex = 10 Then
    '                    'cek expiry date
    '                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
    '                        e.Cancel = True
    '                        Exit Sub
    '                    End If
    '                ElseIf e.ColumnIndex = 11 Then
    '                    'cek location controlled
    '                    If Me.ObjProduct.ApakahLocationControlled(Me.dgvItem.Item(15, e.RowIndex).Value) = False Then
    '                        e.Cancel = True
    '                        Exit Sub
    '                    End If
    '                End If
    '                'kalo sampe sini maka boleh edit
    '                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
    '            Else
    '                e.Cancel = True
    '            End If
    '        Else
    '            'hanya bisa edit pada saat detail aktif
    '            e.Cancel = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub dgvItem_CellBeginEdit_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Location Code
            '5. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColSecondaryQty.Index Or e.ColumnIndex = Me.ColBatch.Index Or e.ColumnIndex = Me.ColLocation.Index _
                Or e.ColumnIndex = Me.ColRemark.Index Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = Me.ColSecondaryQty.Index Then
                        'cek secondary qty
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColBatch.Index Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColExpiredDate.Index Then
                        'cek expiry date
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColLocation.Index Then
                        'cek location controlled
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
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

    'Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
    '    Try
    '        Dim SecondaryQty, JmlDGV As Double
    '        Dim BatchNo As String = ""
    '        Dim ExpiryDate As Date = #1/1/1950#
    '        Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

    '        'pastikan ada isinya
    '        With row
    '            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
    '                'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
    '                If e.ColumnIndex <> 10 Then
    '                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
    '                        Exit Sub
    '                    End If
    '                Else
    '                    If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
    '                        If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
    '                            Exit Sub
    '                        End If
    '                    End If
    '                End If

    '                If e.ColumnIndex <> 13 Then
    '                    'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
    '                    If IsDate(.Cells(10).Value) = True Then
    '                        ExpiryDate = CDate(.Cells(10).Value)
    '                    Else
    '                        ExpiryDate = #1/1/1950#
    '                    End If
    '                    If IsNumeric(.Cells(7).Value) = True Then
    '                        SecondaryQty = .Cells(7).Value
    '                    Else
    '                        SecondaryQty = -1
    '                    End If
    '                    JmlDGV = 0

    '                    If e.ColumnIndex = 5 Then
    '                        Me.EndEditQty(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
    '                    ElseIf e.ColumnIndex = 7 Then
    '                        Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
    '                    ElseIf e.ColumnIndex = 9 Then
    '                        Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
    '                    ElseIf e.ColumnIndex = 11 Then
    '                        Me.EndEditLocation(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
    '                    End If
    '                End If
    '            Else
    '                If e.ColumnIndex = 13 Then
    '                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
    '                Else
    '                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
    '                End If
    '            End If
    '        End With
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim SecondaryQty, JmlDGV As Double
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> Me.ColExpiredDate.Index Then
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColRemark.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(ColExpiredDate.Index).Value) = True Then
                            ExpiryDate = CDate(.Cells(ColExpiredDate.Index).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(ColSecondaryQty.Index).Value) = True Then
                            SecondaryQty = .Cells(ColSecondaryQty.Index).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = ColQtyMulti.Index Then
                            Me.EndEditQty_Multi(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = ColSecondaryQty.Index Then
                            Me.EndEditSecondaryUnit_Multi(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = ColBatch.Index Then
                            Me.EndEditBatchNo_Multi(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = ColLocation.Index Then
                            Me.EndEditLocation_Multi(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = ColRemark.Index Then
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

    Private Sub EndEditQty_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value)
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.ObjProduct.ApakahGoods(row.Cells(15).Value) = True Then
                    JmlDGV = .Cells(5).Value
                    If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                    .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditQty_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)
        Try
            With row
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value))

                Me.SetQtyDefaultOnDgv(e.RowIndex)
                If Me.ObjProduct.ApakahGoods(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If

                Me.SetQtyDefaultOnDgv(e.RowIndex)

                If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) = 0 Then
                    MsgBox("Qty restored to previous value, because qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(e.RowIndex)
                    Exit Sub
                End If


            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnit_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                .Cells(9).Value, ExpiryDate, .Cells(7).Value, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama_Default
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris_Default()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditSecondaryUnit_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColBatch.Index).Value, ExpiryDate, .Cells(Me.ColSecondaryQty.Index).Value, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama_Multi
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris_Multi()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNo_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'batch no
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else

                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItem.Item(10, e.RowIndex).Value) = True Then
                        'basis data
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(15, e.RowIndex).Value, _
                        Me.dgvItem.Item(9, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItem(10, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(9, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(10, e.RowIndex).Value)
                                    Me.dgvItem.Item(10, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        'data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItem.Rows
                            If row1.Cells(9).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells(15).Value = Me.dgvItem.Item(15, e.RowIndex).Value Then
                                If CDate(row1.Cells(10).Value) <> CDate(Me.dgvItem.Item(10, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(9, e.RowIndex).Value) + " is " + _
                                        Format(CDate(row1.Cells(10).Value), "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItem.Item(10, e.RowIndex).Value)
                                        Me.dgvItem.Item(10, e.RowIndex).Value = row1.Cells(10).Value
                                        Exit For
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '3. cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama_Default
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem.Item(10, e.RowIndex).Value = BufferExpiryDate
                            Exit Sub
                        End If
                        Me.GabungkanBaris_Default()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditBatchNo_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'batch no
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else

                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value) = True Then
                        'basis data
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItem(Me.ColExpiredDate.Index, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) + " is " + Format(ExpiryDateLama, "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        'data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItem.Rows
                            If row1.Cells(Me.ColBatch.Index).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value Then
                                If CDate(row1.Cells(Me.ColExpiredDate.Index).Value) <> CDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) + " is " + _
                                        Format(CDate(row1.Cells(Me.ColExpiredDate.Index).Value), "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value)
                                        Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value = row1.Cells(Me.ColBatch.Index).Value
                                        Exit For
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '3. cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama_Multi
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = BufferExpiryDate
                            Exit Sub
                        End If
                        Me.GabungkanBaris_Multi()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocation_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)
        Dim Obj As New clsMasterStorageArea
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                If Obj.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterStorageArea
                    If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <> "" Then
                        i.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea <> -1 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetStorageAreaCode(i.PrimaryStorageArea)).ToUpper
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        i = Nothing
                        Exit Sub
                    End If
                    i = Nothing
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                End If
                'update data primary location
                .Cells(16).Value = Obj.GetPrimary(.Cells(11).Value)
                'update data gudang
                .Cells(12).Value = Obj.GetWarehouseCode(.Cells(16).Value)

                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    'kembalikan ke data semula
                    .Cells(16).Value = Obj.GetPrimary(.Cells(11).Value)
                    'update data gudang
                    .Cells(12).Value = Obj.GetWarehouseCode(.Cells(16).Value)
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama_Default
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'update data primary location
                            .Cells(16).Value = Obj.GetPrimary(.Cells(11).Value)
                            'update data gudang
                            .Cells(12).Value = Obj.GetWarehouseCode(.Cells(16).Value)
                            Exit Sub
                        End If
                        Me.GabungkanBaris_Default()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditLocation_Multi(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)
        Dim Obj As New clsMasterStorageArea
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                If Obj.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterStorageArea
                    If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <> "" Then
                        i.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea <> -1 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetStorageAreaCode(i.PrimaryStorageArea)).ToUpper
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        i = Nothing
                        Exit Sub
                    End If
                    i = Nothing
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                End If
                'update data primary location
                .Cells(Me.ColPrimaryLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                'update data gudang
                .Cells(Me.ColWareHouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)

                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    'kembalikan ke data semula
                    .Cells(Me.ColPrimaryLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                    'update data gudang
                    .Cells(Me.ColWareHouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama_Multi
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'update data primary location
                            .Cells(Me.ColPrimaryLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                            'update data gudang
                            .Cells(Me.ColWareHouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                            Exit Sub
                        End If
                        Me.GabungkanBaris_Multi()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Function ApakahAdaBarisYangSama_Default() As Boolean
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

                    If row.Cells(1).Value = row1.Cells(1).Value And row.Cells(2).Value = row1.Cells(2).Value And _
                    row.Cells(3).Value = row1.Cells(3).Value And row.Cells(4).Value = row1.Cells(4).Value And _
                    CStr(row.Cells(9).Value) = CStr(row1.Cells(9).Value) And row.Cells(11).Value = _
                    row1.Cells(11).Value Then
                        If IsNumeric(row.Cells(7).Value) = True And IsNumeric(row1.Cells(7).Value) = True Then
                            If row.Cells(7).Value = row1.Cells(7).Value Then
                                Return True
                            Else
                                Return False
                            End If
                        ElseIf IsNumeric(row.Cells(7).Value) = False And IsNumeric(row1.Cells(7).Value) = False Then
                            Return True
                        End If
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

    ''@rosadinnn
    ''@MUaT
    Private Function ApakahAdaBarisYangSama_Multi() As Boolean
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

                    If row.Cells(ColPrimaryPOnoDoNo.Index).Value = row1.Cells(ColPrimaryPOnoDoNo.Index).Value And row.Cells(Me.ColSuppDONo.Index).Value = row1.Cells(Me.ColSuppDONo.Index).Value And _
                    row.Cells(Me.ColCodeNo.Index).Value = row1.Cells(Me.ColCodeNo.Index).Value And row.Cells(Me.ColDescription.Index).Value = row1.Cells(Me.ColDescription.Index).Value And _
                    CStr(row.Cells(Me.ColBatch.Index).Value) = CStr(row1.Cells(Me.ColBatch.Index).Value) And row.Cells(Me.ColLocation.Index).Value = _
                    row1.Cells(Me.ColLocation.Index).Value Then
                        If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                            If row.Cells(ColSecondaryQty.Index).Value = row1.Cells(ColSecondaryQty.Index).Value Then
                                Return True
                            Else
                                Return False
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                            Return True
                        End If
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
    Private Sub GabungkanBaris_Default()
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
                    If CStr(row.Cells(1).Value) = CStr(row1.Cells(1).Value) And CStr(row.Cells(2).Value) = CStr(row1.Cells(2).Value) And _
                    CStr(row.Cells(3).Value) = CStr(row1.Cells(3).Value) And CStr(row.Cells(4).Value) = CStr(row1.Cells(4).Value) And _
                    CStr(row.Cells(9).Value) = CStr(row1.Cells(9).Value) And _
                    CStr(row.Cells(11).Value) = CStr(row1.Cells(11).Value) Then
                        If IsNumeric(row.Cells(7).Value) = True And IsNumeric(row1.Cells(7).Value) = True Then
                            If CDbl(row.Cells(7).Value) = CDbl(row1.Cells(7).Value) Then
                                Me.GabungkanBaris_Default(row, row1)
                                Me.ObjCD.SetNumber(Me.dgvItem)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(7).Value) = False And IsNumeric(row1.Cells(7).Value) = False Then
                            Me.GabungkanBaris_Default(row, row1)
                            Me.ObjCD.SetNumber(Me.dgvItem)
                            Exit Sub
                        End If
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

    ''@rosadinnn
    ''@MUaT
    Private Sub GabungkanBaris_Multi()
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
                    If CStr(row.Cells(Me.ColPONoDONo.Index).Value) = CStr(row1.Cells(Me.ColPONoDONo.Index).Value) And CStr(row.Cells(Me.ColSuppDONo.Index).Value) = CStr(row1.Cells(Me.ColSuppDONo.Index).Value) And _
                    CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
                    CStr(row.Cells(Me.ColBatch.Index).Value) = CStr(row1.Cells(Me.ColBatch.Index).Value) And _
                    CStr(row.Cells(Me.ColLocation.Index).Value) = CStr(row1.Cells(Me.ColLocation.Index).Value) Then
                        If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                            If CDbl(row.Cells(Me.ColSecondaryQty.Index).Value) = CDbl(row1.Cells(Me.ColSecondaryQty.Index).Value) Then
                                Me.GabungkanBaris_Multi(row, row1)
                                Me.ObjCD.SetNumber(Me.dgvItem)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                            Me.GabungkanBaris_Multi(row, row1)
                            Me.ObjCD.SetNumber(Me.dgvItem)
                            Exit Sub
                        End If
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

    Private Sub GabungkanBaris_Default(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow)
        Try
            'gabungkan data qty.
            row.Cells(5).Value = CDbl(row.Cells(5).Value) + CDbl(row1.Cells(5).Value)
            'gabungkan data remarks
            If CStr(row.Cells(13).Value).EndsWith(".") = False Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(13).Value = CStr(row.Cells(13).Value) + ". "
            End If
            row.Cells(13).Value = CStr(row.Cells(13).Value) + CStr(row1.Cells(13).Value)
            'hapus baris ybs
            Me.dgvItem.Rows.RemoveAt(row1.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub GabungkanBaris_Multi(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow)
        Try
            'gabungkan data qty.
            row.Cells(Me.ColQtyMulti.Index).Value = CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(row1.Cells(Me.ColQtyMulti.Index).Value)

            Me.SetQtyDefaultOnDgv(row.Index)

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

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
    ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard
        Dim DefPrec As Integer = 0

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV_Default(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    DefPrec = Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(JmlGudang + JmlObjLama - JmlDGVLain, DefPrec) + " " + _
                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            'cek stock untuk tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(PrimaryProduct, Me.dtpDate.Value)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV_Multi(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + "  minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    DefPrec = Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)
                    MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(JmlGudang + JmlObjLama - JmlDGVLain, DefPrec) + " " + _
                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'sampe sini 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function

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
            Me.ObjCD.SetNumber(Me.dgvItem)
            If Me.dgvItem.Rows.Count = 0 Then
                Me.TxtPurchaseReturnNumber.ReadOnly = False
                Me.TxtPurchaseReturnNumber.BackColor = Me.txtPrimaryCategory.BackColor
                Me.btnFindPurchaseReturn.Enabled = True
            End If

            ''@rosadinnn
            ''@MUaT
            Me.FormatDGV_Multi()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'ketika user akan menghapus data, tidak perlu cek apakah akibatnya akan mengakibatkan ada stock yang minus, karena akan mengakibatkan
            'stock menjadi positif
            'tetapi untuk transaksi yang berkaitan dengan transaksi lainnya (DO, Invoice, GRN atau PO) harus cek validitas jumlah berkaitan dengan mereka
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            Me.TxtRemark.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            Me.TxtRemark.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            Me.TxtCodeNo.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            Me.TxtCodeNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.GotFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.GotFocus
        Try
            Me.txtSecondaryQty.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.LostFocus
        Try
            Me.txtSecondaryQty.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        Try
            Me.TxtRemarkItem.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        Try
            Me.TxtRemarkItem.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNo.SelectedIndexChanged
        Try
            If Me.chkLocationControlled.Checked = True Then
                If Me.cboBatchNo.Text <> "" Then
                    'isikan data lokasi
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text, Me.cboBatchNo.Text)
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
            'isikan secondary unit
            If Me.chkLocationControlled.Checked = False And _
            Me.chkUseSecondaryUnit.Checked = True And Me.cboBatchNo.Text <> "" Then
                Me.IsiCboSecondary(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text, _
                Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocation_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocation.SelectedIndexChanged
        Dim objLoc As New clsMasterStorageArea

        Try
            If Me.cboLocation.Text = "" Or objLoc.ApakahSudahAda(Me.cboLocation.Text) = False Then
                Me.txtPrimaryLocation.Text = "-1"
                Me.TxtWarehouse.Clear()
            Else
                Me.txtPrimaryLocation.Text = objLoc.GetPrimary(Me.cboLocation.Text)
                Me.TxtWarehouse.Text = objLoc.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                Me.IsiCboSecondary(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text, Me.cboBatchNo.Text, _
                Me.txtPrimaryLocation.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objLoc = Nothing
        End Try
    End Sub

    Private Sub cboSecondaryUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSecondaryUnits.SelectedIndexChanged
        Try
            If Me.cboSecondaryUnits.SelectedIndex <> -1 Then
                Me.txtSecondaryQty.Text = Me.cboSecondaryUnits.Text
            Else
                Me.txtSecondaryQty.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodePihakAcuan_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtKodePihakAcuan.KeyDown
        Dim i As New FrmCariMasterSupplier
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjSupplier.GetJml(Me.TxtKodePihakAcuan.Text, Core_Module.EnumSearchCriteria.ByCode)
                If Jml = 1 Then
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtRemark.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Supplier dengan kode yang sama atau tidak ada Supplier yang dimaksud
                    If Me.TxtKodePihakAcuan.Text <> "" Then
                        i.txtCari.Text = Me.TxtKodePihakAcuan.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimaryPihakAcuan.Text = i.PrimarySupplier
                        Me.TxtKodePihakAcuan.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtRemark.Focus()
                    Else
                        Me.TxtKodePihakAcuan.Clear()
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtKodePihakAcuan.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtKodePihakAcuan_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodePihakAcuan.Leave
        Dim i As New FrmCariMasterSupplier

        Try
            If Me.TxtKodePihakAcuan.Text <> "" Then
                If Me.ObjSupplier.GetJml(Me.TxtKodePihakAcuan.Text) = 1 Then
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtRemark.Focus()
                Else
                    If Me.TxtKodePihakAcuan.Text <> "" Then
                        i.txtCari.Text = Me.TxtKodePihakAcuan.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimaryPihakAcuan.Text = i.PrimarySupplier
                        Me.TxtKodePihakAcuan.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtRemark.Focus()
                    Else
                        Me.TxtKodePihakAcuan.Clear()
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtKodePihakAcuan.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtPihakAcuan.Clear()
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnCariPihakAcuan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCariPihakAcuan.Click
        Dim i As New FrmCariMasterSupplier

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtKodePihakAcuan.Text <> "" Then
                i.txtCari.Text = Me.TxtKodePihakAcuan.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtPrimaryPihakAcuan.Text = i.PrimarySupplier
                Me.TxtKodePihakAcuan.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtRemark.Focus()
            Else
                Me.TxtKodePihakAcuan.Clear()
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.KosongkanItem()
                Me.TxtKodePihakAcuan.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub


    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.objStockOut.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objStockOut.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtStockInNo.Text = Me.objStockOut.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindPurchaseOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPurchaseOrder.Click
        Dim i As New FrmCariTransaksiPO
        Try
            '  i.cboCriteria.SelectedIndex = 0
            If Me.txtPurchaseOrderNo.Text <> "" Then
                i.txtCari.Text = Me.txtPurchaseOrderNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryPO > 0 Then
                Me.txtPrimaryPO.Text = i.PrimaryPO
                Me.txtPurchaseOrderNo.Text = Me.objPO.GetNoTransaksiPO(Me.txtPrimaryPO.Text)
            Else
                Me.txtPrimaryPO.Text = "-1"
                Me.txtPurchaseOrderNo.Text = ""
                Me.txtPurchaseOrderNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtPurchaseOrderNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPurchaseOrderNo.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Double = objPO.GetPrimary(Me.txtPurchaseOrderNo.Text)
                If primary > -1 Then
                    Me.txtPrimaryPO.Text = primary
                    Me.txtPurchaseOrderNo.Text = Me.objPO.GetNoTransaksiPO(primary)
                Else
                    Me.btnFindPurchaseOrder.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPurchaseOrderNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPurchaseOrderNo.Leave
        Try
            If Me.txtPurchaseOrderNo.Text <> "" Then
                Dim primary As Double = objPO.GetPrimary(Me.txtPurchaseOrderNo.Text)
                If primary > -1 Then
                    Me.txtPrimaryPO.Text = primary
                    Me.txtPurchaseOrderNo.Text = Me.objPO.GetNoTransaksiPO(primary)
                Else
                    Me.btnFindPurchaseOrder.PerformClick()
                End If
            Else
                Me.txtPrimaryPO.Text = "-1"
                Me.txtPurchaseOrderNo.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindPurchaseReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPurchaseReturn.Click
        Dim i As New FrmCariTransaksiPurchaseReturn
        Dim obj As New ClsTransaksiPurchaseReturn
        i.BackColorAll = Me.BackColorAll
        Dim SupName, SupCode, NoPR, SupDONO As String
        SupName = ""
        SupCode = ""
        NoPR = ""
        SupDONO = ""

        Try
            If Me.TxtPurchaseReturnNumber.Text <> "" Then
                i.txtCari.Text = Me.TxtPurchaseReturnNumber.Text
            End If
            i.ShowDialog()
            If i.PrimaryPurchaseReturn > 0 Then
                'cek apakah sudah digunakan oleh yang lain
                If Me.CekApakahSudahDigunakan(Me.TxtPrimaryStockIn.Text, i.PrimaryPurchaseReturn) = False Then
                    Exit Sub
                End If

                If Me.TxtPihakAcuan.Text <> "" Then
                    obj.GetInfoSupplierNameNCodeNTransNo(i.PrimaryPurchaseReturn, SupName, SupCode, NoPR, SupDONO)
                    If SupCode <> Me.TxtKodePihakAcuan.Text Then
                        MsgBox("Supplier for this Purchase Return is " + SupName + ".", MsgBoxStyle.Information, "Information")
                        If IsNumeric(Me.TxtPrimaryPurchaseReturn.Text) = False Then
                            Me.TxtPrimaryPurchaseReturn.Text = "-1"
                        End If
                        Me.TxtPurchaseReturnNumber.Text = obj.GetNoTransaksi(Me.TxtPrimaryPurchaseReturn.Text)
                        Me.TxtPurchaseReturnNumber.Focus()
                    Else
                        Me.TxtPrimaryPurchaseReturn.Text = i.PrimaryPurchaseReturn
                        Me.TxtPurchaseReturnNumber.Text = NoPR
                        Me.TxtKodePihakAcuan.Text = SupCode
                        Me.TxtPihakAcuan.Text = SupName
                        Me.IsikanDGVpurchasereturn(Me.DGVPurchaseReturn, Me.TxtPrimaryPurchaseReturn.Text)
                        Me.TransferDataPR_Multi()
                        Me.TxtKodePihakAcuan.Enabled = False
                        Me.BtnCariPihakAcuan.Enabled = False
                        Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                    End If
                Else
                    Me.TxtPrimaryPurchaseReturn.Text = i.PrimaryPurchaseReturn
                    obj.GetInfoSupplierNameNCodeNTransNo(Me.TxtPrimaryPurchaseReturn.Text, Me.TxtPihakAcuan.Text, _
                            Me.TxtKodePihakAcuan.Text, Me.TxtPurchaseReturnNumber.Text, Me.TxtSupplierDONumber.Text)
                    Me.IsikanDGVpurchasereturn(Me.DGVPurchaseReturn, Me.TxtPrimaryPurchaseReturn.Text)
                    TransferDataPR_Multi()
                    Me.TxtKodePihakAcuan.Enabled = False
                    Me.BtnCariPihakAcuan.Enabled = False
                    Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                    Me.TxtRemark.Focus()
                End If
            Else
                Me.TxtPrimaryPurchaseReturn.Text = "-1"
                Me.TxtPurchaseReturnNumber.Text = ""
                Me.TxtKodePihakAcuan.Clear()
                Me.TxtPihakAcuan.Clear()
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtSupplierDONumber.Clear()
                Me.TxtPurchaseReturnNumber.Focus()
                Me.DGVPurchaseReturn.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub IsikanDGVpurchasereturn(ByVal vDGV As DataGridView, ByVal vPrimaryPR As Double)
        Dim objItem As New ClsTransaksiPurchaseReturnItem
        Try
            objItem.IsiDGV_Multi(vDGV, vPrimaryPR)

            ''@rosadinnn
            ''@MUaT
            Me.SetMultiUnitProperty()

            Me.FormatDGV_PR()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objItem = Nothing
        End Try
    End Sub

    Private Sub TxtPurchaseReturnNumber_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPurchaseReturnNumber.KeyDown
        Dim obj As New ClsTransaksiPurchaseReturn
        Dim SupName, SupCode, NoTrans, SupDONO As String
        SupName = ""
        SupCode = ""
        NoTrans = ""
        SupDONO = ""

        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Double = obj.GetPrimary(Me.TxtPurchaseReturnNumber.Text)
                If primary > -1 Then
                    'cek apakah sudah digunakan oleh yang lain
                    If Me.CekApakahSudahDigunakan(Me.TxtPrimaryStockIn.Text, primary) = False Then
                        Exit Sub
                    End If

                    If Me.TxtPihakAcuan.Text <> "" Then
                        obj.GetInfoSupplierNameNCodeNTransNo(primary, SupName, SupCode, NoTrans, SupDONO)
                        If SupCode <> Me.TxtKodePihakAcuan.Text Then
                            MsgBox("Supplier for this Packing List is " + SupName + ".", MsgBoxStyle.Information, "Information")
                            If IsNumeric(Me.TxtPrimaryPurchaseReturn.Text) = False Then
                                Me.TxtPrimaryPurchaseReturn.Text = "-1"
                            End If
                            Me.TxtPurchaseReturnNumber.Text = obj.GetNoTransaksi(Me.TxtPrimaryPurchaseReturn.Text)
                            Me.TxtPurchaseReturnNumber.Focus()
                        Else
                            Me.TxtPrimaryPurchaseReturn.Text = primary
                            Me.TxtPurchaseReturnNumber.Text = NoTrans
                            Me.TxtKodePihakAcuan.Text = SupCode
                            Me.TxtPihakAcuan.Text = SupName
                            Me.IsikanDGVpurchasereturn(Me.DGVPurchaseReturn, Me.TxtPrimaryPurchaseReturn.Text)
                            TransferDataPR_Multi()
                            Me.TxtKodePihakAcuan.Enabled = False
                            Me.BtnCariPihakAcuan.Enabled = False
                            Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                        End If
                    Else
                        Me.TxtPrimaryPurchaseReturn.Text = primary
                        obj.GetInfoSupplierNameNCodeNTransNo(Me.TxtPrimaryPurchaseReturn.Text, Me.TxtPihakAcuan.Text, _
                        Me.TxtKodePihakAcuan.Text, Me.TxtPurchaseReturnNumber.Text, Me.TxtSupplierDONumber.Text)
                        Me.IsikanDGVpurchasereturn(Me.DGVPurchaseReturn, Me.TxtPrimaryPurchaseReturn.Text)
                        TransferDataPR_Multi()
                        Me.TxtKodePihakAcuan.Enabled = False
                        Me.BtnCariPihakAcuan.Enabled = False
                        Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                        Me.TxtRemark.Focus()
                    End If
                Else
                    Me.btnFindPurchaseReturn.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini mengecek, purchase return yang dituliskan, tidak boleh pernah digunakan oleh stock out purchase return lainnya
    ''' kalo edit, tentu saja ini pengecualian untuk dirinya sendiri :)
    ''' </summary>
    ''' <remarks></remarks>
    Private Function CekApakahSudahDigunakan(ByVal vPrimaryStockOut As Double, ByVal vPrimaryPR As Double) As Boolean
        Dim RetVal As String = ""

        Try
            RetVal = Me.ObjPR.ApakahBolehDiRujukStockOutPurchaseReturn(vPrimaryStockOut, vPrimaryPR)
            If RetVal <> "" Then
                MsgBox("This purchase return is already refered by stock out purchase return: " + RetVal, MsgBoxStyle.Information, "Information")
                'kosongkan data purchase return
                Me.TxtPrimaryPurchaseReturn.Text = "-1"
                Me.TxtPurchaseReturnNumber.Text = ""
                Me.TxtKodePihakAcuan.Clear()
                Me.TxtPihakAcuan.Clear()
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtSupplierDONumber.Clear()
                Me.TxtPurchaseReturnNumber.Focus()
                Me.DGVPurchaseReturn.Rows.Clear()

                'nilai balikan
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtPurchaseReturnNumber_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPurchaseReturnNumber.Leave
        Dim obj As New ClsTransaksiPurchaseReturn
        Try
            If Me.TxtRemark.Enabled = True And Me.TxtPurchaseReturnNumber.Enabled = True Then
                If Me.TxtPurchaseReturnNumber.Text <> "" Then
                    Dim primary As Double = obj.GetPrimary(Me.TxtPurchaseReturnNumber.Text)
                    If primary > -1 Then
                        'cek apakah sudah digunakan oleh yang lain
                        If Me.CekApakahSudahDigunakan(Me.TxtPrimaryStockIn.Text, primary) = False Then
                            Exit Sub
                        End If

                        If Me.TxtPihakAcuan.Text <> "" Then
                            Dim SupName, SupCode, NoTrans, SupDONO As String
                            SupName = ""
                            SupCode = ""
                            NoTrans = ""
                            SupDONO = ""

                            obj.GetInfoSupplierNameNCodeNTransNo(primary, SupName, SupCode, NoTrans, SupDONO)
                            If SupCode <> Me.TxtKodePihakAcuan.Text Then
                                MsgBox("Supplier for this Packing List is " + SupName + ".", MsgBoxStyle.Information, "Information")
                                If IsNumeric(Me.TxtPrimaryPurchaseReturn.Text) = False Then
                                    Me.TxtPrimaryPurchaseReturn.Text = "-1"
                                End If
                                Me.TxtPurchaseReturnNumber.Text = obj.GetNoTransaksi(Me.TxtPrimaryPurchaseReturn.Text)
                                Me.TxtPurchaseReturnNumber.Focus()
                            Else
                                Me.TxtPrimaryPurchaseReturn.Text = primary
                                Me.TxtPurchaseReturnNumber.Text = NoTrans
                                Me.TxtKodePihakAcuan.Text = SupCode
                                Me.TxtPihakAcuan.Text = SupName
                                Me.IsikanDGVpurchasereturn(Me.DGVPurchaseReturn, Me.TxtPrimaryPurchaseReturn.Text)
                                TransferDataPR_Multi()
                                Me.TxtKodePihakAcuan.Enabled = False
                                Me.BtnCariPihakAcuan.Enabled = False
                                Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                            End If
                        Else
                            Me.TxtPrimaryPurchaseReturn.Text = primary
                            obj.GetInfoSupplierNameNCodeNTransNo(Me.TxtPrimaryPurchaseReturn.Text, Me.TxtPihakAcuan.Text, _
                            Me.TxtKodePihakAcuan.Text, Me.TxtPurchaseReturnNumber.Text, Me.TxtSupplierDONumber.Text)
                            Me.IsikanDGVpurchasereturn(Me.DGVPurchaseReturn, Me.TxtPrimaryPurchaseReturn.Text)
                            TransferDataPR_Multi()
                            Me.TxtKodePihakAcuan.Enabled = False
                            Me.BtnCariPihakAcuan.Enabled = False
                            Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                            Me.TxtRemark.Focus()
                        End If
                    Else
                        Me.btnFindPurchaseReturn.PerformClick()
                    End If
                Else
                    Me.TxtPrimaryPurchaseReturn.Text = "-1"
                    Me.TxtPurchaseReturnNumber.Text = ""
                    Me.TxtKodePihakAcuan.Clear()
                    Me.TxtPihakAcuan.Clear()
                    Me.TxtPrimaryPihakAcuan.Text = "-1"
                    Me.TxtSupplierDONumber.Clear()
                    Me.DGVPurchaseReturn.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Function TransferDataPR_Default() As Boolean
        Dim row_PR, rowDGV As DataGridViewRow
        Dim Index As Double = -1
        Dim BatchNumberControlled As Boolean = False
        Dim TambahBarisBaru As Boolean = False
        Dim Jml As Double = 0
        Dim ED As Date = #1/1/1950#

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'yang bisa ditransfer hanya yang tidak kontrol lokasi dan tidak kontrol sec qty
            For i As Integer = 0 To Me.DGVPurchaseReturn.Rows.Count - 1
                row_PR = Me.DGVPurchaseReturn.Rows(i)
                If Me.ObjProduct.ApakahLocationControlled(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value) = False Then
                    BatchNumberControlled = Me.ObjProduct.ApakahBatchNoControlled(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value)
                    If IsDate(row_PR.Cells(Me.ColExpiryDateDGVPR.Index).Value) = True Then
                        ED = row_PR.Cells(Me.ColExpiryDateDGVPR.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    If Me.ObjProduct.ApakahUseSecondary(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value) = False Then
                        'cari baris yang sudah ada
                        Index = Me.GetIndex_Multi(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value, -1, row_PR.Cells(Me.ColBatchNoDGVPR.Index).Value, -1, _
                        -1, ED)

                        If Index >= 0 Then
                            'bila ada, tambahkan qty saja
                            rowDGV = Me.dgvItem.Rows(Index)
                            Jml = Me.GetJmlDGV_Multi(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value, -1, -1, row_PR.Cells(Me.ColBatchNoDGVPR.Index).Value, BatchNumberControlled, False, False)
                            If CDbl(row_PR.Cells(Me.ColTotalQtyDGVPR.Index).Value) > Jml Then
                                rowDGV.Cells(Me.ColQty.Index).Value = CDbl(rowDGV.Cells(Me.ColQty.Index).Value) + _
                                CDbl(row_PR.Cells(Me.ColTotalQtyDGVPR.Index).Value) - Jml
                            End If
                        Else
                            'bila tidak ada, maka tambahkan baris baru
                            TambahBarisBaru = True

                            Me.dgvItem.Rows.Add()
                            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                            With rowDGV
                                .Cells(Me.ColCodeNo.Index).Value = row_PR.Cells(Me.colCodeNoDGVPR.Index).Value
                                .Cells(Me.ColPONoDONo.Index).Value = row_PR.Cells(Me.ColPONODGVPR.Index).Value
                                .Cells(2).Value = ""
                                .Cells(Me.ColDescription.Index).Value = row_PR.Cells(Me.colDescDGVPR.Index).Value
                                .Cells(Me.ColQty.Index).Value = row_PR.Cells(Me.ColTotalQtyDGVPR.Index).Value
                                .Cells(Me.ColUnit.Index).Value = row_PR.Cells(Me.colUnitDGVPR.Index).Value
                                .Cells(Me.ColSecondaryQty.Index).Value = -1
                                .Cells(Me.ColSecondaryUnit.Index).Value = ""
                                .Cells(Me.ColBatch.Index).Value = row_PR.Cells(Me.ColBatchNoDGVPR.Index).Value
                                .Cells(Me.ColExpiredDate.Index).Value = row_PR.Cells(Me.ColExpiryDateDGVPR.Index).Value
                                .Cells(Me.ColLocation.Index).Value = ""
                                .Cells(Me.ColWareHouse.Index).Value = ""
                                .Cells(Me.ColRemark.Index).Value = ""
                                .Cells(Me.ColPrimaryPOnoDoNo.Index).Value = row_PR.Cells(Me.ColPrimaryPODGVPR.Index).Value
                                .Cells(Me.ColPrimaryCodeNo.Index).Value = row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value
                                .Cells(Me.ColPrimaryLocation.Index).Value = -1
                                .Cells(Me.ColPrimaryUnits.Index).Value = row_PR.Cells(Me.ColPrimaryUnitDGVPR.Index).Value
                                .Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                                .Cells(Me.ColPrimaryPdO.Index).Value = -1
                            End With
                        End If
                    End If
                End If
            Next

            If TambahBarisBaru = True Then
                Me.ObjCD.SetNumber(Me.dgvItem, False)
                Me.TxtPurchaseReturnNumber.ReadOnly = True
                Me.btnFindPurchaseReturn.Enabled = False
                Me.TxtPurchaseReturnNumber.BackColor = Me.TxtReadOnly.BackColor
                Me.FormatDGV_Multi()
            End If

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function TransferDataPR_Multi() As Boolean
        Dim row_PR, rowDGV As DataGridViewRow
        Dim Index As Double = -1
        Dim BatchNumberControlled As Boolean = False
        Dim TambahBarisBaru As Boolean = False
        Dim Jml As Double = 0
        Dim ED As Date = #1/1/1950#
        Try
            If Me.TxtCodeNo.Enabled = False Then Return False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For i As Integer = 0 To Me.DGVPurchaseReturn.Rows.Count - 1
                row_PR = Me.DGVPurchaseReturn.Rows(i)
                If Me.ObjProduct.ApakahLocationControlled(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value) = False Then
                    BatchNumberControlled = Me.ObjProduct.ApakahBatchNoControlled(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value)
                    If IsDate(row_PR.Cells(Me.ColExpiryDateDGVPR.Index).Value) = True Then
                        ED = row_PR.Cells(Me.ColExpiryDateDGVPR.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    If Me.ObjProduct.ApakahUseSecondary(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value) = False Then
                        Index = Me.GetIndex_Multi(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value, -1, row_PR.Cells(Me.ColBatchNoDGVPR.Index).Value, -1, -1, ED)
                        If Index >= 0 Then
                            rowDGV = Me.dgvItem.Rows(Index)
                            Jml = Me.GetJmlDGV_Multi(row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value, -1, -1, row_PR.Cells(Me.ColBatchNoDGVPR.Index).Value, BatchNumberControlled, False, False)
                            If CDbl(row_PR.Cells(Me.ColTotalQtyDGVPR.Index).Value) > Jml Then
                                rowDGV.Cells(Me.ColQtyMulti.Index).Value = CDbl(rowDGV.Cells(Me.ColQtyMulti.Index).Value) + _
                                CDbl(row_PR.Cells(Me.ColQtyMulti_PR.Index).Value) - Jml

                                ''@rosadinnn
                                ''@MUaT
                                SetQtyDefaultOnDgv(Index)

                            End If
                        Else
                            TambahBarisBaru = True

                            Me.dgvItem.Rows.Add()
                            rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                            With rowDGV
                                .Cells(Me.ColNo.Index).Value = ""
                                .Cells(Me.ColPONoDONo.Index).Value = row_PR.Cells(Me.ColPONODGVPR.Index).Value
                                .Cells(Me.ColSuppDONo.Index).Value = ""
                                .Cells(Me.ColCodeNo.Index).Value = row_PR.Cells(Me.colCodeNoDGVPR.Index).Value
                                .Cells(Me.ColDescription.Index).Value = row_PR.Cells(Me.colDescDGVPR.Index).Value
                                .Cells(Me.ColQty.Index).Value = row_PR.Cells(Me.ColTotalQtyDGVPR.Index).Value
                                .Cells(Me.ColUnit.Index).Value = row_PR.Cells(Me.colUnitDGVPR.Index).Value
                                .Cells(Me.ColSecondaryQty.Index).Value = -1
                                .Cells(Me.ColSecondaryUnit.Index).Value = ""
                                .Cells(Me.ColBatch.Index).Value = row_PR.Cells(Me.ColBatchNoDGVPR.Index).Value
                                .Cells(Me.ColExpiredDate.Index).Value = row_PR.Cells(Me.ColExpiryDateDGVPR.Index).Value
                                .Cells(Me.ColLocation.Index).Value = ""
                                .Cells(Me.ColWareHouse.Index).Value = ""
                                .Cells(Me.ColRemark.Index).Value = ""
                                .Cells(Me.ColPrimaryPOnoDoNo.Index).Value = row_PR.Cells(Me.ColPrimaryPODGVPR.Index).Value
                                .Cells(Me.ColPrimaryCodeNo.Index).Value = row_PR.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value
                                .Cells(Me.ColPrimaryLocation.Index).Value = -1
                                .Cells(Me.ColPrimaryUnits.Index).Value = row_PR.Cells(Me.ColPrimaryUnitDGVPR.Index).Value
                                .Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                                .Cells(Me.ColPrimaryPdO.Index).Value = -1

                                ''@rosadinnn
                                ''@MUaT
                                .Cells(Me.ColQtyMulti.Index).Value = row_PR.Cells(Me.ColQtyMulti_PR.Index).Value
                                .Cells(Me.ColUnitCodeMulti.Index).Value = row_PR.Cells(Me.ColUnitCodeMulti_PR.Index).Value
                                .Cells(Me.ColRatioDesc.Index).Value = row_PR.Cells(Me.ColRatioDesc_PR.Index).Value
                                .Cells(Me.ColRatioReference.Index).Value = row_PR.Cells(Me.ColRaioReference_PR.Index).Value
                                .Cells(Me.ColRatioDefault.Index).Value = row_PR.Cells(Me.ColRatioDefault_Pr.Index).Value
                                .Cells(Me.ColPrimaryUnitMulti.Index).Value = row_PR.Cells(Me.ColPrimaryUnitsMulti_Pr.Index).Value

                            End With
                        End If
                    End If
                End If
            Next

            If TambahBarisBaru = True Then
                Me.ObjCD.SetNumber(Me.dgvItem, False)
                Me.TxtPurchaseReturnNumber.ReadOnly = True
                Me.btnFindPurchaseReturn.Enabled = False
                Me.TxtPurchaseReturnNumber.BackColor = Me.TxtReadOnly.BackColor
                Me.FormatDGV_Multi()
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtPurchaseReturnNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseReturnNumber.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPurchaseReturnNumber.ReadOnly = True Then
                    Me.TxtPurchaseReturnNumber.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPurchaseReturnNumber.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPurchaseReturnNumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPurchaseReturnNumber.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPurchaseReturnNumber.ReadOnly = True Then
                    Me.TxtPurchaseReturnNumber.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtPurchaseReturnNumber.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then

                    ''@rosadinnn
                    ''@MUaT
                    If Me.txtPrimaryUnitsMulti.Text = "" Then
                        Me.txtPrimaryUnitsMulti.Text = "-1"
                    End If
                    Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.SetQtyDefault()

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub SetQtyDefaultOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, vRowIdx).Value) * _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value)), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar_Multi() As Boolean
        Try
            If Me.TxtPrimaryPurchaseReturn.Text = "" Or Me.TxtPrimaryPurchaseReturn.Text = "-1" Then
                MsgBox("Specify purchase return first.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.TxtRatioDefault.Text = "" Or Me.TxtRatioDefault.Text = "0" Then
                MsgBox("Invalid ratio default.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.TxtRatioDefault.Text = "" Or Me.TxtRatioDefault.Text = "0" Then
                MsgBox("Invalid ratio multi.", MsgBoxStyle.Information, "Information")
                Return False
            End If


            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If



            If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True Then
                If CDec(Me.TxtQtyperBatchMulti.Text) > 0 And (Me.TxtQtyperBatchDefault.Text = "" Or Me.TxtQtyperBatchDefault.Text = "0") Then
                    MsgBox("Invalid qty default.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            Me.SetMultiUnitProperty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsContainMatchData(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Boolean
        Try
            For Each RowDgv As DataGridViewRow In Me.DGVPurchaseReturn.Rows
                With RowDgv
                    If CDbl(.Cells(Me.ColPrimaryPODGVPR.Index).Value) = vPrimaryPO And CDbl(.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value = vPrimaryProduct) Then
                        Return True
                    End If
                End With
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Enum EnumGetOption As Integer
        PRIMARY_UNIT_DEFAULT
        PRIMARY_UNIT_MULTI
        UNIT_CODE_DEFAULT
        UNIT_CODE_MULTI
        RATIO_REFERENCE
        RATIO_DEFAULT
        RATIO_DESC
    End Enum

    Private Function GetMultiUnitProperty(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double, ByVal vEnumGetOption As EnumGetOption) As String
        Try
            For Each RowDgv As DataGridViewRow In Me.DGVPurchaseReturn.Rows
                With RowDgv
                    If CDbl(.Cells(Me.ColPrimaryPODGVPR.Index).Value) = vPrimaryPO And CDbl(.Cells(Me.colPrimaryCodeNoDGVPR.Index).Value = vPrimaryProduct) Then
                        Select Case vEnumGetOption
                            Case EnumGetOption.PRIMARY_UNIT_DEFAULT
                                Return .Cells(Me.ColPrimaryUnitDGVPR.Index).Value.ToString
                            Case EnumGetOption.PRIMARY_UNIT_MULTI
                                Return .Cells(Me.ColPrimaryUnitsMulti_Pr.Index).Value.ToString
                            Case EnumGetOption.RATIO_DEFAULT
                                Return ObjCD.DigitGrouping(.Cells(Me.ColRatioDefault_Pr.Index).Value.ToString, ObjProduct.GetAngkaDecimal(vPrimaryProduct))
                            Case EnumGetOption.RATIO_DESC
                                Return .Cells(Me.ColRatioDesc_PR.Index).Value
                            Case EnumGetOption.RATIO_REFERENCE
                                Return ObjCD.DigitGrouping(.Cells(Me.ColRaioReference_PR.Index).Value, ObjProduct.GetAngkaDecimal(vPrimaryProduct, CDbl(.Cells(Me.ColPrimaryUnitsMulti_Pr.Index).Value)))
                            Case EnumGetOption.UNIT_CODE_DEFAULT
                                Return .Cells(Me.colUnitDGVPR.Index).Value
                            Case EnumGetOption.UNIT_CODE_MULTI
                                Return .Cells(Me.ColUnitCodeMulti_PR.Index).Value
                        End Select
                    End If
                End With
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub SetMultiUnitProperty()
        Try
            If Me.TxtPrimaryPurchaseReturn.Text <> "-1" And Me.TxtPrimaryPurchaseReturn.Text <> "" And _
                    Me.txtPrimaryPO.Text <> "-1" And Me.txtPrimaryPO.Text <> "" And _
                    Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                If IsContainMatchData(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) = True Then
                    Me.txtPrimaryUnitsDefault.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, EnumGetOption.PRIMARY_UNIT_DEFAULT)
                    Me.txtPrimaryUnitsMulti.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, EnumGetOption.PRIMARY_UNIT_MULTI)
                    Me.TxtRatioReference.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, EnumGetOption.RATIO_REFERENCE)
                    Me.TxtRatioDefault.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, EnumGetOption.RATIO_DEFAULT)
                    Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.lblUnitsMulti.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, EnumGetOption.UNIT_CODE_MULTI)
                    Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                    Me.lblUnitsDefault.Text = Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, EnumGetOption.UNIT_CODE_DEFAULT)
                    Me.LblQty.Text = "Qty (" & Me.GetMultiUnitProperty(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, EnumGetOption.RATIO_DESC) & ")"
                    SetQtyDefault()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryPO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryPO.TextChanged
        Try
            Me.SetMultiUnitProperty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            If Me.TxtRatioDefault.Text <> "0" And Me.TxtRatioReference.Text <> "0" And Me.TxtRatioDefault.Text <> "" And Me.TxtRatioReference.Text <> "" Then
                Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVPurchaseReturn_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVPurchaseReturn.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                Dim i As New FrmLaporanFormatStockInOut_Bundling

                i.primary = Me.TxtPrimaryStockIn.Text
                i.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutPurchaseReturn
                i.ShowDialog()

                i = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If

            If Me.TxtRemark.Enabled = False Then
                If Me.objStockOut.CariDetail(vPrimary, Me.dgvItem, Me.TblItemCustom) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub EditTanggal(ByVal vTglTujuan As Date)
        'Dim hasil As MsgBoxResult

        Try
            If Me.TxtStockInNo.Text <> "" Then
                'If Me.objStockOut.ApakahBolehEdit(Me.UserCode) = True Then

                Me.baru = False
                Me.AktifkanDetail()
                Me.dgvItem.Focus()

                'ganti tanggal
                Me.dtpDate.Value = vTglTujuan

                'If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    'hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    'If hasil = MsgBoxResult.Yes Then
                    '    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    '    If Me.objStockOut.SimpanBaru(Me.dgvItem, Me.UserCode, False, Me.vPrimaryPickUpOrder) Then
                    '        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    '        Me.SalinKeField()
                    '        Me.NonAktifkanDetail()
                    '        Me.btnFind.Focus()
                    '        If Me.vFromPickUpOrder = True Then
                    '            Me.Close()
                    '        End If
                    '    End If
                    'End If
                Else
                    'konfirmasi disimpan atau tidak
                    'hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    'If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.objStockOut.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                        'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.SalinKeField()
                        Me.NonAktifkanDetail()
                        Me.btnEdit.Focus()
                    End If
                    'End If
                End If
                'End If
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_SOPR_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal vTblPrimaryPR As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryPO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core)
        Dim TblTempPO As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.objStockOut.ResetObject(Me.dgvItem)
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn

            Me.objStockOut.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.TblItemCustom)
            Me.objStockOut.PrimaryPOPurchase = ObjCutOff_Core.GetPrimaryUpdate(Me.objStockOut.PrimaryPOPurchase, vTblPrimaryPR)

            'Update Primary PO
            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value) = False And rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value <> "-1" Then
                    TblTempPO.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value)
                    If TblTempPO.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value), vTblPrimaryPO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value), PrimaryUpdate, TblTempPO)
                        rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value), TblTempPO)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_SOPR_Save()
        Try
            Me.baru = True
            Me.objStockOut.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, -1, False, False, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblAdvance.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(Me.lblAdvance, 0, 10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiStockOutPurchaseReturn2_AddByCustomProduct

        Try

            If Me.TxtPrimaryPurchaseReturn.Text = "-1" And Me.TxtRemark.Enabled = True Then
                MsgBox("Specify purchase return first", MsgBoxStyle.Information, "Information")
                Me.TxtPurchaseReturnNumber.Focus()
                Exit Sub
            End If
            i.PrimaryPurchaseReturn = Me.TxtPrimaryPurchaseReturn.Text
            i.baru = Me.baru

            If Me.TxtRemark.Enabled = True Then
                i.edit = True
            Else
                i.edit = False
            End If
            i.BackColorAll = Me.BackColorAll
            i.TanggalTransaksi = Me.dtpDate.Value.Date
            i.PurchaseReturnNumber = Me.TxtPurchaseReturnNumber.Text
            i.JenisTransAsalCustom = Me.TipeTransaksi
            i.PrimaryStockIn = Me.TxtPrimaryStockIn.Text
            i.PrimaryPO = Me.txtPrimaryPO.Text

            If Me.baru = False Then
                Me.TblDBItemCustom = Me.TblItemCustom
                i.TBLDBCustomItem = Me.TblDBItemCustom.Copy
            Else
                i.TBLDBCustomItem = Me.TblItemCustom.Copy
            End If
            i.TBLCustomItem = Me.TblItemCustom.Copy

            If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TblItemCustom = i.TBLCustomItem.Copy

                Me.IsiDgvItemFromTblCustom()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    'TR
    Public TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
    Public TblDBItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
    Sub Atur1()
        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutPurchaseReturn") = True Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutPurchaseReturn2_AddByCustomProduct") = False Then
                    'Me.gboItem.Enabled = True
                    Me.SplitContainer2.Enabled = True
                    Me.AddItemByCustom.Visible = False
                    Me.dgvItem.ReadOnly = False
                    Me.dgvItem.AllowUserToDeleteRows = True
                Else
                    Me.dgvItem.ReadOnly = True
                    'Me.gboItem.Enabled = False
                    Me.SplitContainer2.Enabled = False
                    Me.AddItemByCustom.Visible = True
                    Me.dgvItem.AllowUserToDeleteRows = False
                    'Me.dgvItem.Rows.Clear()
                End If
            Else
                'Me.gboItem.Enabled = True
                Me.SplitContainer2.Enabled = True
                Me.AddItemByCustom.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
            End If

            Me.lblAdvance.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiDgvItemFromTblCustom() As Boolean
        Dim RowCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemRow
        Dim RowDgv As DataGridViewRow
        Try

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'kosongkan dgv
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AllowUserToAddRows = False

            For Each RowCustom In Me.TblItemCustom.Rows
                Dim Index As Integer = Me.GetIndex_Multi(RowCustom.Primary_Product, RowCustom.Primary_Storage_Area, RowCustom.Batch_No, _
                                    -1, RowCustom.Primary_PO_Purchase, RowCustom.Expiry_Date)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + RowCustom.Qty_Multi_Unit, ObjProduct.GetAngkaDecimal(RowCustom.Primary_Product, RowCustom.Primary_Multi_Units))
                    row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + RowCustom.Qty, ObjProduct.GetAngkaDecimal(RowCustom.Primary_Product))

                    If RowCustom.Is_Kain = True Then
                        row.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value), RowCustom.dec_multi) _
                        + " " + row.Cells(Me.ColUnitCodeMulti.Index).Value + " = " + ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value), _
                        RowCustom.dec) + " " + row.Cells(Me.ColUnit.Index).Value
                    End If
                    If String.IsNullOrEmpty(row.Cells(Me.ColRemark.Index).Value) = False Then
                        row.Cells(Me.ColRemark.Index).Value = row.Cells(Me.ColRemark.Index).Value & ", " & RowCustom.Remarks
                    Else
                        row.Cells(Me.ColRemark.Index).Value = RowCustom.Remarks
                    End If
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If RowCustom.Expiry_Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(RowCustom.Expiry_Date, "dd MMM yyyy")
                    End If

                    ''@rosadinnn
                    ''@MUaT
                    'Me.dgvItem.Rows.Add("", Me.txtPurchaseOrderNo.Text, "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatchMulti.Text, _
                    'Me.lblUnitsMulti.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.cboBatchNo.Text, _
                    'ExpiredDateStr, Me.cboLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnitsDefault.Text, _
                    'Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)
                    Me.dgvItem.Rows.Add()
                    RowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPONoDONo.Index).Value = RowCustom.po_so_number
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = RowCustom.Product_Code
                        .Cells(Me.ColDescription.Index).Value = RowCustom.Description
                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(RowCustom.Qty, RowCustom.dec)
                        .Cells(Me.ColUnit.Index).Value = RowCustom.unit
                        .Cells(Me.ColSecondaryQty.Index).Value = -1
                        .Cells(Me.ColSecondaryUnit.Index).Value = ""
                        .Cells(Me.ColBatch.Index).Value = RowCustom.Batch_No
                        .Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColLocation.Index).Value = RowCustom.Storage_Area_Code
                        .Cells(Me.ColWareHouse.Index).Value = RowCustom.Warehouse_Code
                        .Cells(Me.ColRemark.Index).Value = RowCustom.Remarks
                        .Cells(Me.ColPrimaryPOnoDoNo.Index).Value = RowCustom.Primary_PO_Purchase
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = RowCustom.Primary_Product
                        .Cells(Me.ColPrimaryLocation.Index).Value = RowCustom.Primary_Storage_Area
                        .Cells(Me.ColPrimaryUnits.Index).Value = RowCustom.Primary_Units
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryPdO.Index).Value = RowCustom.Primary_PO_Purchase
                        .Cells(Me.ColPrimaryUnitMulti.Index).Value = RowCustom.Primary_Multi_Units
                        .Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(RowCustom.Qty_Multi_Unit, RowCustom.dec_multi)
                        .Cells(Me.ColUnitCodeMulti.Index).Value = RowCustom.unit_multi

                        If RowCustom.Is_Kain = False Then
                            If RowCustom.Primary_Units <> RowCustom.Primary_Multi_Units Then
                                .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(RowCustom.Rasio, _
                                       RowCustom.dec_multi) _
                                        & " " & RowCustom.unit_multi & " = " & Me.ObjCD.DigitGrouping(RowCustom.Ratio_Default, _
                                        RowCustom.dec) & " " & RowCustom.unit
                            Else
                                .Cells(Me.ColRatioDesc.Index).Value = "-"
                            End If
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(RowCustom.Qty_Roll, RowCustom.dec_multi) _
                            + " " + RowDgv.Cells(Me.ColUnitCodeMulti.Index).Value + " = " + ObjCD.DigitGrouping(RowCustom.Qty, _
                            RowCustom.dec) + " " + RowDgv.Cells(Me.ColUnit.Index).Value
                        End If
                        .Cells(Me.ColRatioReference.Index).Value = RowCustom.Rasio
                        .Cells(Me.ColRatioDefault.Index).Value = RowCustom.Ratio_Default

                    End With

                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    'setelah menambahkan baris baru, data purchase return tidak boleh diubah
                    Me.TxtPurchaseReturnNumber.ReadOnly = True
                    Me.TxtPurchaseReturnNumber.BackColor = Me.TxtReadOnly.BackColor
                    Me.btnFindPurchaseReturn.Enabled = False
                End If
            Next


            'kosongkan gbo item
            Me.KosongkanItem()
            Me.TxtCodeNo.Focus()
            'format dgv
            Me.FormatDGV_Multi()
            'fokuskan ke baris terakhir
            While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
            Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
            End While

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

End Class