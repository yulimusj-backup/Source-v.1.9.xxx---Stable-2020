Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockOutSales
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public UserCode As String = ""
    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public EnableDOOverride As Boolean = False
    Public FormNo As Integer = 0

    Dim baru As Boolean
    Dim objStockOut As New clsTransaksiStockInOut
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim objStockOutOutItem As New clsTransaksiStockInOutItem
    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL3 As New dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item1DataTable
    Dim TBL5 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item2DataTable
    Dim TBL6 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item3DataTable
    Dim ObjCustomer As New clsMasterCustomer
    Dim buffer As String = ""
    Dim ObjDO As New ClsTransaksiDO
    Dim ObjDOItem As New ClsTransaksiDOItem
    'Dim FrmDODetail As New FrmTransaksiDOView
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim CloseManual As Boolean = False
    Dim ObjLS As New clsMasterListedSoftware
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim ObjLicense As New ClsLicense
    Dim ObjAcc As New Core_Module.ClsTranskasiAccClosing
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjUser As New ClsUser

    Dim TBLItemDO As New dseTransaksiStockInOut.tbltrans_do_item_view3DataTable
    Dim DTAItemDO As New dseTransaksiStockInOutTableAdapters.tbltrans_do_item_view3TableAdapter
    Dim DTAProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
    Dim rowItemDO As dseTransaksiStockInOut.tbltrans_do_item_view3Row

    Dim TBLLocationPriority As New Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable
    Dim ObjLocationPriority As New Core_Module.ClsMasterLocationPriority
    Dim LewatiFocusBatch As Boolean = False
    Dim IsPriorityLocation As Boolean = True
    Dim IsIncludeOtherLocation As Boolean = False
    Dim ApakahBolehDONonStock As Boolean = False
    Dim ApakahBolehOverrideQty As Boolean = False

    Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
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
            Me.FormatDGV()

            If Me.baru = False Then
                Me.txtDONo.Enabled = False
            Else
                Me.txtDONo.Enabled = True
            End If

            Me.AturHakAksesFormCustom()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
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
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.txtPrimaryDO.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.DgvItemDO2.Rows.Clear()
            Me.TblItemCustom.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimarySO.Text = "-1"
            Me.TxtQtyperBatchDefault.Text = "0"
            Me.TxtQtyperBatchMulti.Text = "0"
            Me.txtSecondaryQty.Text = "-1"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnitsDefault.Left = Me.lblSecondaryUnits.Left
            'Me.TxtQtyperBatchDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyperBatchDefault.Left - 6
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
            Me.TxtRatioDefault.Text = "1"
            Me.TxtRatioReference.Text = "1"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim ObjSC As New ClsTransaksiStockCard
        Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal

        Try
            'data item salah bila:
            '0. sales order no yang dimasukkan salah
            If Me.txtSalesOrderNo.Text = "" Then
                If Me.txtDONo.Text <> "" Then
                    MsgBox("Please specify sales order no.", MsgBoxStyle.Information, "Information")
                    Me.txtSalesOrderNo.Text = ""
                    Me.txtSalesOrderNo.Focus()
                Else
                    MsgBox("Please specify delivery order no.", MsgBoxStyle.Information, "Information")
                    Me.txtDONo.Text = ""
                    Me.txtDONo.Focus()
                End If
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
                If Me.LewatiFocusBatch = False Then
                    Me.TxtQtyperBatchMulti.Focus()
                End If
                Return False
            End If
            '2.b. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, _
                "Information")
                If Me.LewatiFocusBatch = False Then
                    Me.TxtQtyperBatchMulti.Focus()
                End If
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
                JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.txtPrimaryPDO.Text, _
                Me.DtpExpired.Value)
                JmlDGV = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
                Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchDefault.Text)
                If JmlSetelahTambah < 0 Then
                    If Me.chkUseSecondaryUnit.Checked = True Then
                        MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGV) + " " + _
                        Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, "Information")
                    Else
                        Dim Max As String = ObjCD.DigitGrouping(CStr((JmlGudang + JmlObjLama - JmlDGV) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text)), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                        MsgBox("Maximum value you can enter is " & Max & " " & Me.lblUnitsMulti.Text & ".", MsgBoxStyle.Information, "Information")
                    End If
                    If Me.LewatiFocusBatch = False Then
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                    Return False
                End If

                '9.b. Cek stock pada tanggal yang bersangkutan
                JmlGudang = ObjSC.GetStock(Me.TxtPrimaryCodeNo.Text, Me.dtpDate.Value)
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
                            Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.txtPrimaryPDO.Text, _
                            Me.DtpExpired.Value)
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
                            Dim Max As String = ObjCD.DigitGrouping(CStr((JmlGudang + JmlObjLama - JmlDGV) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text)), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Max + " " + _
                            Me.lblUnitsMulti.Text + ".", MsgBoxStyle.Information, _
                            "Information")
                        End If
                    Else
                        MsgBox("Stock level on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + _
                        " is 0. Entering this data will cause negative value in your stock.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    If Me.LewatiFocusBatch = False Then
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                    Return False
                End If
            End If

            '10. Bila transaksi berkaitan dengan sales dan terinstall software sales maka ada batasan 
            '    berkaitan dengan jumlah item DO. Berkaitan dengan jumlah item DO secara global, 
            '    hanya memperdulikan primary SO dan primary product
            Dim JmlDo As Decimal = Me.ObjDO.GetJml2_Multi(Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)
            JmlDGV = Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
            If JmlDo - JmlDGV - CDec(Me.TxtQtyperBatchMulti.Text) < 0 Then
                MsgBox("Maximum value you can enter for:" + Me.TxtCodeNo.Text + " and SO number: " + Me.txtSalesOrderNo.Text + " is " + ObjCD.DigitGrouping(CStr(JmlDo - JmlDGV), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) + " " + _
                lblUnitsMulti.Text + ".", MsgBoxStyle.Information, "Information")
                If Me.LewatiFocusBatch = False Then
                    Me.TxtQtyperBatchMulti.Focus()
                End If
                Return False
            End If

            JmlDo = Me.ObjDO.GetJml2_Default(Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)
            JmlDGV = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
            If JmlDo - JmlDGV - CDbl(Me.TxtQtyperBatchDefault.Text) < 0 Then
                MsgBox("Maximum value you can enter for:" + Me.TxtCodeNo.Text + " and SO number: " + Me.txtSalesOrderNo.Text + " is " + ObjCD.DigitGrouping(CStr(JmlDo - JmlDGV), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + _
                lblUnitsDefault.Text + ".", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchDefault.Focus()
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function



    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And vPrimarySO = row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value Then
                    Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And vPrimarySO = row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value Then
                    Jml = Jml + CDec(row.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double, ByVal vBatchNumber As String) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And vPrimarySO = row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value _
                And row.Cells(Me.ColBatch.Index).Value = vBatchNumber Then
                    Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim TglConvSMESS As Date = #1/1/1950# 
        Dim rowDGV As DataGridViewRow

        Try
            'If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
            '    MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            'tidak boleh bikin stock in out, pada tanggal sebelum convertion date
            TglConvSMESS = Me.ObjAcc.GetSMESSConvertionDate
            If Me.dtpDate.Value.Date < TglConvSMESS.Date Then
                MsgBox("Minimum transaction date is on SMESS convertion date: " + Format(TglConvSMESS, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("Stock out number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            KoleksiNamaData.Add("Delivery order no")
            KoleksiTextBox.Add(Me.txtDONo)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'harus jelas sales ke siapa
            If Me.TxtPrimaryPihakAcuan.Text = "-1" Then
                MsgBox("Please specify customer's code", MsgBoxStyle.Information, "Information")
                Me.TxtKodePihakAcuan.Focus()
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ApakahBolehDONonStock = True Then
                'cek sales order balance

                'jika boleh do non stock, maka tidak cek kecocokan dengan do per batch
                'karena boleh ditimpa batchnya pada do
                'yang perlu dicek adalah kesesuaian dengan jumlah secara keseluruhan mengabaikan batch dan lokasi
                If Me.ApakahBolehOverrideQty = True Then
                    'hanya cek sales order balance saja
                    'ga cek apa2 lagi disini
                Else
                    'cek qty berdasarkan primary so dan primary product
                    'hanya boleh edit bacth dan location saja

                End If
            Else
                'cek data secara keseluruhan, kecocokannya dengan DO, tidak hanya berkaitan dengan jumlah
                'tetapi juga berkaitan dengan batch no dan secondary qty.
                Me.ObjDO.CariDetail(Me.txtPrimaryDO.Text, Me.dgvItemDO)
                Me.GabungkanBarisDO2()
                Dim rowDO, rowStockOut As DataGridViewRow
                Dim Jml As Decimal = 0
                Dim JmlDefault As Decimal = 0

                Dim tblDoKiri As New dseTransaksiStockInOut.DataTable22DataTable
                Dim rowDoKiri As dseTransaksiStockInOut.DataTable22Row
                Dim tblDoKanan As New dseTransaksiStockInOut.DataTable22DataTable
                Dim rowDoKanan As dseTransaksiStockInOut.DataTable22Row
                For Each rowDgvDoKiri As DataGridViewRow In Me.dgvItem.Rows
                    With rowDgvDoKiri
                        tblDoKiri.DefaultView.RowFilter = "[primary product] = " & .Cells(Me.ColPrimaryCodeNo.Index).Value.ToString & " and [batch no] = '" & .Cells(Me.ColBatch.Index).Value & "'"
                        If tblDoKiri.DefaultView.Count = 0 Then
                            rowDoKiri = tblDoKiri.NewDataTable22Row
                            rowDoKiri.Primary_Product = .Cells(Me.ColPrimaryCodeNo.Index).Value
                            rowDoKiri.Batch_No = .Cells(Me.ColBatch.Index).Value
                            rowDoKiri.Qty_Multi = .Cells(Me.ColQtyMulti.Index).Value
                            rowDoKiri.Qty_Default = .Cells(Me.ColQty.Index).Value
                            tblDoKiri.Rows.Add(rowDoKiri)
                        Else
                            rowDoKiri = tblDoKiri.DefaultView(0).Row
                            rowDoKiri.Qty_Multi += .Cells(Me.ColQtyMulti.Index).Value
                            rowDoKiri.Qty_Default += .Cells(Me.ColQty.Index).Value
                        End If
                    End With
                Next
                For Each rowDgvDokanan As DataGridViewRow In Me.dgvItemDO.Rows
                    With rowDgvDokanan
                        tblDoKanan.DefaultView.RowFilter = "[primary product] = " & .Cells(Me.ColDO2_PrimaryCodeNo.Index).Value.ToString & " and [batch no] = '" & .Cells(Me.ColDO2_BatchNo.Index).Value & "'"
                        If tblDoKanan.DefaultView.Count = 0 Then
                            rowDoKanan = tblDoKanan.NewDataTable22Row
                            rowDoKanan.Primary_Product = .Cells(Me.ColDO2_PrimaryCodeNo.Index).Value
                            rowDoKanan.Batch_No = .Cells(Me.ColDO2_BatchNo.Index).Value
                            rowDoKanan.Qty_Multi = .Cells(Me.ColDO2_TotalQtyMulti.Index).Value
                            rowDoKanan.Qty_Default = .Cells(Me.ColDO2_TotalQty.Index).Value
                            tblDoKanan.Rows.Add(rowDoKanan)
                        Else
                            rowDoKanan = tblDoKanan.DefaultView(0).Row
                            rowDoKanan.Qty_Multi += .Cells(Me.ColDO2_TotalQtyMulti.Index).Value
                            rowDoKanan.Qty_Default += .Cells(Me.ColDO2_TotalQty.Index).Value
                        End If
                    End With
                Next
                For Each rowDoKanan In tblDoKanan
                    With rowDoKanan
                        tblDoKiri.DefaultView.RowFilter = "[primary product] = " & .Primary_Product.ToString & " and [batch no] = '" & .Batch_No & "'"
                        If tblDoKiri.DefaultView.Count > 0 Then
                            rowDoKiri = tblDoKiri.DefaultView(0).Row
                            If rowDoKanan.Qty_Multi <> rowDoKiri.Qty_Multi Then 'Or rowDoKanan.Qty_Default <> rowDoKiri.Qty_Default Then -> 2019-06-17: batasan dilepas sementara karena di ABT, ada hanya kontrol qty rollnya saja, tidak qty m/y
                                If String.IsNullOrEmpty(rowDoKanan.Batch_No) = False Then
                                    MsgBox("Please check item DO Product: " & ObjProduct.GetProductCode(rowDoKanan.Primary_Product) & " Batch: " & rowDoKanan.Batch_No & " has not been fullfilled.", MsgBoxStyle.Information, "Information")
                                Else
                                    MsgBox("Please check item DO Product: " & ObjProduct.GetProductCode(rowDoKanan.Primary_Product) & ".", MsgBoxStyle.Information, "Information")
                                End If
                                Return False
                            End If
                        Else
                            If String.IsNullOrEmpty(rowDoKanan.Batch_No) = False Then
                                MsgBox("Please check item DO Product: " & ObjProduct.GetProductCode(rowDoKanan.Primary_Product) & " Batch: " & rowDoKanan.Batch_No & " has not been fullfilled.", MsgBoxStyle.Information, "Information")
                            Else
                                MsgBox("Please check item DO Product: " & ObjProduct.GetProductCode(rowDoKanan.Primary_Product) & ".", MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End With
                Next
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

            ''cek bisi ada qty default=0
            Dim ListError As String = ""
            For Each rowDGV In Me.dgvItem.Rows
                With rowDGV
                    If CDec(.Cells(Me.ColQty.Index).Value) <= 0 Then
                        ListError += .Cells(Me.ColCodeNo.Index).Value & ", "
                    End If
                End With
            Next
            If ListError <> "" Then
                ListError = ListError.Substring(0, ListError.Length - 3) & "."
                MsgBox("The following product has zero default qty: " & vbNewLine & ListError & vbNewLine & "Please change qty.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_OUT_SALES, Me.TBLDetailLocation) = False Then
                    Return False
                End If
            End If
 
            If Me.dgvItem.RowCount > 0 Then
                For Each rowDGV In Me.dgvItem.Rows  
                    If CDec(rowDGV.Cells(Me.ColRatioDefault.Index).Value) = 1 And CDec(rowDGV.Cells(Me.ColRatioReference.Index).Value) = 1 Then
                        If CDec(rowDGV.Cells(Me.ColQty.Index).Value) <> CDec(rowDGV.Cells(Me.ColQtyMulti.Index).Value) Then
                            MsgBox("Inconsistent multi unit qty, ratio and default qty  for row: " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItem.ClearSelection()
                            rowDGV.Selected = True
                            Me.dgvItem.FirstDisplayedScrollingRowIndex = rowDGV.Index

                            Return False
                        End If
                    End If
                Next
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

    Private Sub IsikanDefPrec(ByVal vTBL As dseTransaksiStockInOut_Sales.TblDefPrecDataTable)
        Dim DTAProduct As New dseTransaksiStockInOut_SalesTableAdapters.tblmaster_productTableAdapter
        Dim TBLProduct As New dseTransaksiStockInOut_Sales.tblmaster_productDataTable
        Dim rowProduct As dseTransaksiStockInOut_Sales.tblmaster_productRow
        Dim PrimaryMin, PrimaryMax As Double
        Dim rowDGV As DataGridViewRow
        Dim ListPrimaryProduct As New List(Of Double)
        Dim TblDefPrec As New dseTransaksiStockInOut_Sales.TblDefPrecDataTable
        Dim rowDefPrec1, rowDefPrec2 As dseTransaksiStockInOut_Sales.TblDefPrecRow

        Try
            PrimaryMin = Double.MaxValue
            PrimaryMax = Double.MinValue

            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimaryProduct.Contains(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) = False Then
                    ListPrimaryProduct.Add(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value)

                    rowDefPrec1 = TblDefPrec.NewTblDefPrecRow
                    rowDefPrec1.Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                    rowDefPrec1.DefPrec = 0

                    TblDefPrec.Rows.Add(rowDefPrec1)

                    If PrimaryMax < rowDefPrec1.Primary_Product Then
                        PrimaryMax = rowDefPrec1.Primary_Product
                    End If
                    If PrimaryMin > rowDefPrec1.Primary_Product Then
                        PrimaryMin = rowDefPrec1.Primary_Product
                    End If
                End If
            Next

            DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
            DTAProduct.Fill(TBLProduct, PrimaryMax, PrimaryMin)

            TblDefPrec.DefaultView.Sort = "[Primary Product] ASC"
            vTBL.Rows.Clear()
            For Each DV As Data.DataRowView In TblDefPrec.DefaultView
                rowDefPrec1 = DV.Row

                rowDefPrec2 = vTBL.NewTblDefPrecRow
                rowDefPrec2.Primary_Product = rowDefPrec1.Primary_Product
                rowDefPrec2.DefPrec = rowDefPrec1.DefPrec
                For Each rowProduct In TBLProduct.Rows
                    If rowProduct.Primary = rowDefPrec2.Primary_Product Then
                        rowDefPrec2.DefPrec = rowProduct._Decimal
                        Exit For
                    ElseIf rowProduct.Primary > rowDefPrec2.Primary_Product Then
                        Exit For
                    End If
                Next

                vTBL.Rows.Add(rowDefPrec2)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAProduct.Dispose()
            TBLProduct.Dispose()
            TblDefPrec.Dispose()
        End Try
    End Sub

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
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objStockOut
                .NoTransaksi = Me.TxtStockInNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Remarks = Me.TxtRemark.Text
                .PrimaryCustomer = Me.TxtPrimaryPihakAcuan.Text
                .PrimaryDOSales = Me.txtPrimaryDO.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objStockOut
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.TxtPrimaryPihakAcuan.Text = .PrimaryCustomer
                Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                Me.txtPrimaryDO.Text = .PrimaryDOSales
                Me.txtDONo.Text = Me.ObjDO.GetDoNo(Me.txtPrimaryDO.Text)
                If Me.txtPrimaryDO.Text <> "-1" Then
                    Me.ObjDOItem.IsiDGV_Multi(Me.DgvItemDO2, Me.txtPrimaryDO.Text)
                    Me.GabungkanBarisDO()
                    Me.FormatDGVDO()
                Else
                    Me.DgvItemDO2.Rows.Clear()
                End If
            End With
            Me.FormatDGV()
            'digit grouping
            'Dim row As DataGridViewRow
            'For Each row In Me.dgvItem.Rows
            '    row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
            '    If IsNumeric(row.Cells(7).Value) = True Then
            '        row.Cells(7).Value = Me.ObjCD.DigitGrouping(row.Cells(7).Value)
            '    End If
            'Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub FrmTransaksiStockOutSales_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    Me.FrmDODetail.Dispose()
    'End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
            Me.DgvItemDO2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.KosongkanField()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.TxtQtyperBatchDefault.MaxLength = 15
            Me.TxtQtyperBatchMulti.MaxLength = 15
            Me.ShowIcon = False

            Me.LinkLabel1.Parent = Me.TabPage1
            Me.LinkLabel1.BringToFront()

            Me.AturHakAksesFormCustom()
            Me.dgvItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvItem.ColumnHeadersHeight = 25
            Me.DgvItemDO2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.DgvItemDO2.ColumnHeadersHeight = 25
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
                        If Me.objStockOut.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
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
            'If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
            '    MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
            '    MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.objStockOut.ResetObject(Me.dgvItem)
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
            Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            Me.txtDONo.Enabled = True
            Me.txtDONo.ReadOnly = False
            Me.txtDONo.BackColor = Me.txtPrimaryDO.BackColor

            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, ClsTransaksiSetting.EnumNamaParameter.DeliveryOrder_AllowNoStock) = "1" Then
                Me.ApakahBolehDONonStock = True
            Else
                Me.ApakahBolehDONonStock = False
            End If
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Sales, _
            ClsTransaksiSetting.EnumNamaParameter.Stock_Out_Sales_AllowOverrideDO) = "1" Then
                Me.ApakahBolehOverrideQty = True
            Else
                Me.ApakahBolehOverrideQty = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiStockInOut_Sales
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
                If Me.objStockOut.ApakahBolehEdit(Me.UserCode, True) = True Then
                    Dim ObjLS As New clsMasterListedSoftware
                    If ObjLS.IsSalesInstalled = True And Me.dgvItem.Item(Me.ColPONoDONo.Index, 0).Value = "" Then
                        MsgBox("You cannot edit this data. This data entered before sales program is installed", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    ObjLS = Nothing
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                End If
                'nilai min juga harus diatas tanggal pengkuncian data terakhir
                'msgbox("belum")
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
                If Me.objStockOut.ApakahBolehHapus(Me.UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    Dim ObjLS As New clsMasterListedSoftware
                    If ObjLS.IsSalesInstalled = True And Me.dgvItem.Item(Me.ColPONoDONo.Index, 0).Value = "" Then
                        MsgBox("You cannot edit this data. This data entered before sales program is installed", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    ObjLS = Nothing
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
                                Me.CloseManual = True
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
                                Me.CloseManual = True
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
                    ketemu = False
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
                ketemu = False
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
                        Me.TBL4.Rows.Add(CStr(row.Cells(Me.ColBatch.Index).Value), CDbl(row.Cells(Me.ColQty.Index).Value), ED)
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
                    ketemu = False
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
                ketemu = False
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
                        Me.TBL5.Rows.Add(CStr(row.Cells(Me.ColBatch.Index).Value), row.Cells(Me.ColLocation.Index).Value, CDbl(row.Cells(Me.ColQty.Index).Value))
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
                    ketemu = False
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
                ketemu = False
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
            If IsNumeric(Me.TxtPrimaryCodeNo.Text) = False Then
                Me.TxtPrimaryCodeNo.Text = "-1"
            End If
            If IsNumeric(Me.txtPrimaryUnitsDefault.Text) = False Then
                Me.txtPrimaryUnitsDefault.Text = "-1"
            End If
            If IsNumeric(Me.txtPrimarySecondaryUnits.Text) = False Then
                Me.txtPrimarySecondaryUnits.Text = "-1"
            End If
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
                'Me.lblUnitsDefault.Left = Me.lblSecondaryUnits.Left
                'Me.TxtQtyperBatchDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyperBatchDefault.Left - 6
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
                Me.cboSecondaryUnits.Visible = False
                Me.cboSecondaryUnits.DataSource = Nothing
                Me.cboSecondaryUnits.Items.Clear()
            Else
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                'Me.lblUnitsDefault.Left = Me.lblSecQty.Left - 37
                'Me.TxtQtyperBatchDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyperBatchDefault.Left - 6
                Me.cboSecondaryUnits.Visible = True
                If Me.chkLocationControlled.Checked = False Then
                    'isikan data secondary jika perlu
                    Me.IsiCboSecondary(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text, _
                    Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                End If
            End If
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.lblUnitsDefault.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            Else

                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.txtPrimaryUnitsMulti.Text = "-1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtRatioReference.Text = "1"
                Me.lblUnitsDefault.Text = "Units"
                Me.lblUnitsMulti.Text = "Units"

                Me.lblSecondaryUnits.Text = "Units"
            End If

            Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(0, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
            Me.SetQtyDefault()

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
                        If Me.LewatiFocusBatch = False Then
                            Me.TxtQtyperBatchMulti.Focus()
                        End If
                    End If
                    Me.SetSonoBaru()
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
                            If Me.LewatiFocusBatch = False Then
                                Me.TxtQtyperBatchMulti.Focus()
                            End If
                        End If
                        Me.SetSonoBaru()
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
                        If Me.LewatiFocusBatch = False Then
                            Me.cboBatchNo.Focus()
                        End If
                    ElseIf Me.cboLocation.Visible = True Then
                        If Me.LewatiFocusBatch = False Then
                            Me.cboLocation.Focus()
                        End If
                    Else
                        If Me.LewatiFocusBatch = False Then
                            Me.TxtQtyperBatchMulti.Focus()
                        End If
                    End If
                    Me.SetSonoBaru()
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
                            If Me.LewatiFocusBatch = False Then
                                Me.TxtQtyperBatchMulti.Focus()
                            End If
                        End If
                        Me.SetSonoBaru()
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

    Private Sub SetSonoBaru()
        Try
            Dim TBLSO As New DseTransaksiDO.tbltrans_soDataTable
            Dim rowSONO As DseTransaksiDO.tbltrans_soRow
            Dim JmlDGV As Double = 0
            Dim JmlDO As Double = 0
            If Me.txtPrimarySO.Text = "" Or Me.txtPrimarySO.Text = "-1" Then
                'isikan no so
                Me.ObjDOItem.IsiListNoSO(Me.txtPrimaryDO.Text, Me.TxtPrimaryCodeNo.Text, TBLSO)
                For Each rowSONO In TBLSO.Rows
                    JmlDGV = Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, rowSONO.Primary_SO)
                    JmlDO = Me.ObjDO.GetJml2_Multi(Me.txtPrimaryDO.Text, rowSONO.Primary_SO, Me.TxtPrimaryCodeNo.Text)
                    If JmlDO - JmlDGV <> 0 Then
                        'ganti no sales order
                        Me.txtPrimarySO.Text = rowSONO.Primary_SO
                        Me.txtSalesOrderNo.Text = rowSONO.So_No
                        Exit For
                    End If
                Next
            Else
                'hitung ulang jumlah so, jangan sampe minus
                JmlDGV = Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
                JmlDO = Me.ObjDO.GetJml2_Multi(Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)
                If JmlDO - JmlDGV = 0 Then
                    'ganti no sales order
                    Me.ObjDOItem.IsiListNoSO(Me.txtPrimaryDO.Text, Me.TxtPrimaryCodeNo.Text, TBLSO)
                    For Each rowSONO In TBLSO.Rows
                        If rowSONO.Primary_SO <> Me.txtPrimarySO.Text Then
                            JmlDGV = Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, rowSONO.Primary_SO)
                            JmlDO = Me.ObjDO.GetJml2_Multi(Me.txtPrimaryDO.Text, rowSONO.Primary_SO, Me.TxtPrimaryCodeNo.Text)
                            If JmlDO - JmlDGV <> 0 Then
                                'ganti no sales order
                                Me.txtPrimarySO.Text = rowSONO.Primary_SO
                                Me.txtSalesOrderNo.Text = rowSONO.So_No
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            TBLSO = Nothing
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
                    If Me.LewatiFocusBatch = False Then
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                End If
                Me.SetSonoBaru()
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
    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date, ByVal vPrimarySO As _
    Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And SecondaryQty = -1 _
                        And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
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
        Try
            If Me.ApakahDataItemBenar = True Then
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value, Me.txtPrimarySO.Text)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatchMulti.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.SetQtyDefaultOnDgv(row.Index)
                    'row.Cells(5).Value = CDbl(row.Cells(5).Value) + CDbl(Me.TxtQtyperBatchDefault.Text)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If

                    Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                    Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                    Me.dgvItem.Rows.Add()
                    Dim row As DataGridViewRow
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    row.Cells(Me.ColNo.Index).Value = ""
                    row.Cells(Me.ColPONoDONo.Index).Value = Me.txtSalesOrderNo.Text
                    row.Cells(Me.ColSuppDONo.Index).Value = ""
                    row.Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                    row.Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                    row.Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                    row.Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                    row.Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                    row.Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                    row.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text
                    row.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                    row.Cells(Me.ColLocation.Index).Value = Me.cboLocation.Text
                    row.Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                    row.Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                    row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = Me.txtPrimarySO.Text
                    row.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                    row.Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                    row.Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                    row.Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                    row.Cells(Me.ColPrimaryPdO.Index).Value = Me.txtPrimaryPDO.Text
                    row.Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                    row.Cells(Me.ColUnitsCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                    row.Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                    row.Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                    row.Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                    If Me.txtPrimaryUnitsDefault.Text <> Me.txtPrimaryUnitsMulti.Text Then
                        row.Cells(Me.ColRatioDesc.Index).Value = Me.TxtRatioReference.Text _
                        & " " & Me.lblUnitsMulti.Text & " = " & Me.TxtRatioDefault.Text & " " & Me.lblUnitsDefault.Text
                    Else
                        row.Cells(Me.ColRatioDesc.Index).Value = "-"
                    End If

                    'Me.dgvItem.Rows.Add("", Me.txtSalesOrderNo.Text, "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatchDefault.Text, _
                    'Me.lblUnitsDefault.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.cboBatchNo.Text, _
                    'ExpiredDateStr, Me.cboLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnitsDefault.Text, _
                    'Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If
                'kosongkan gbo item
                Me.KosongkanItem()
                If LewatiFocusBatch = False Then
                    Me.TxtCodeNo.Focus()

                End If
                'format dgv
                Me.FormatDGV()
                'fokuskan ke baris terakhir
                If LewatiFocusBatch = False Then
                    While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                                    Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If

                'DO tidak bisa diubah lagi
                Me.txtDONo.Enabled = False
                Me.btnFindDO.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
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
                        row.Cells(ColExpiredDate.Index).Value = Format(CDate(row.Cells(ColExpiredDate.Index).Value), "dd MMM yyyy")
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

            For Each RowDgv As DataGridViewRow In Me.DgvItemDO2.Rows
                With RowDgv
                    If .Cells(Me.ColDO_PrimaryUnit.Index).Value <> .Cells(Me.ColDO_PrimaryUnitsMulti.Index).Value Then
                        IsMultiUnit = True
                        Exit For
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

    Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub txtSecondaryQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQty, sender, e)
    End Sub

    Private Sub txtSecondaryQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQty.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQty, sender, e)
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Location Code
            '5. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColSecondaryQty.Index Or e.ColumnIndex = Me.ColBatch.Index Or e.ColumnIndex = Me.ColLocation.Index _
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
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItem.Item(ColPrimaryCodeNo.Index, e.RowIndex).Value) = False Then
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
                        If IsDate(.Cells(Me.ColExpiredDate.Index).Value) = True Then
                            ExpiryDate = CDate(.Cells(ColExpiredDate.Index).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                            SecondaryQty = .Cells(Me.ColSecondaryQty.Index).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = Me.ColQtyMulti.Index Then
                            Me.EndEditQty(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColQty.Index Then
                            Me.EndEditQty_Default(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColSecondaryQty.Index Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColBatch.Index Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColLocation.Index Then
                            Me.EndEditLocation(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemark.Index Then
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

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value))


                Me.SetQtyDefaultOnDgv(row.Index)

                ''jika lebih atau sama dengan batas qty di DO, udah default/multi dengan sisanya
                Dim jmlDgvDoMulti As Decimal = 0
                Dim jmlDgvDoDefault As Decimal = 0
                Dim jmlDgvItemMulti As Decimal = 0
                Dim jmlDgvItemDefault As Decimal = 0
                For Each row2 As DataGridViewRow In DgvItemDO2.Rows
                    If row2.Cells(Me.ColDO_PrimarySO.Index).Value = row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value And row2.Cells(Me.ColDO_PrimaryCodeNo.Index).Value = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        row2.Cells(Me.ColDO_BatchNo.Index).Value = row.Cells(Me.ColBatch.Index).Value Then
                        jmlDgvDoMulti += CDec(row2.Cells(Me.ColDO_TotalQtyMulti.Index).Value)
                        jmlDgvDoDefault += CDec(row2.Cells(Me.ColDO_TotalQty.Index).Value)
                    End If
                Next
                For Each row2 As DataGridViewRow In dgvItem.Rows
                    If row2.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = .Cells(Me.ColPrimaryPOnoDoNo.Index).Value And row2.Cells(Me.ColPrimaryCodeNo.Index).Value = .Cells(Me.ColPrimaryCodeNo.Index).Value And row2.Index <> row.Index And _
                        row2.Cells(Me.ColBatch.Index).Value = .Cells(Me.ColBatch.Index).Value Then
                        jmlDgvItemMulti += CDec(row2.Cells(Me.ColQtyMulti.Index).Value)
                        jmlDgvItemDefault += CDec(row2.Cells(Me.ColQty.Index).Value)
                    End If
                Next
                If (CDec(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value) + jmlDgvItemMulti) >= jmlDgvDoMulti Then
                    Me.dgvItem(Me.ColQty.Index, e.RowIndex).Value = ObjCD.DigitGrouping(jmlDgvDoDefault - jmlDgvItemDefault, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                    Me.dgvItem(Me.ColQtyMulti.Index, e.RowIndex).Value = ObjCD.DigitGrouping(jmlDgvDoMulti - jmlDgvItemMulti, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                End If

                If CDec(row.Cells(Me.ColQty.Index).Value) <= 0 Then
                    MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(row.Index)
                    Exit Sub
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.ObjProduct.ApakahGoods(.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyDefaultOnDgv(row.Index)
                        Exit Sub
                    End If
                End If

                'cek data sales order balance
                Dim JmlDo As Double = Me.ObjDO.GetJml2_Multi(Me.txtPrimaryDO.Text, .Cells(Me.ColPrimaryPOnoDoNo.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlDGV = Me.GetJmlDGV_Multi(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPOnoDoNo.Index).Value)
                Dim JmlDGVSendiri As Double = .Cells(Me.ColQtyMulti.Index).Value
                If JmlDo - JmlDGV < 0 Then
                    MsgBox("Maximum value you can enter for:" + CStr(.Cells(Me.ColCodeNo.Index).Value) + _
                    " and SO number: " + CStr(.Cells(Me.ColPONoDONo.Index).Value) + " is " + Me.ObjCD.DigitGrouping(CStr(JmlDo - JmlDGV + JmlDGVSendiri), Me.ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryUnitsMulti.Index).Value)) + " " + _
                    row.Cells(Me.ColUnitsCodeMulti.Index).Value + ".", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(row.Index)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty_Default(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
                ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
                ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                Me.SetQtyMultiOnDgv(row.Index)

                ''jika lebih atau sama dengan batas qty di DO, udah default/multi dengan sisanya
                Dim jmlDgvDoMulti As Decimal = 0
                Dim jmlDgvDoDefault As Decimal = 0
                Dim jmlDgvItemMulti As Decimal = 0
                Dim jmlDgvItemDefault As Decimal = 0
                For Each row2 As DataGridViewRow In DgvItemDO2.Rows
                    If row2.Cells(Me.ColDO_PrimarySO.Index).Value = row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value And row2.Cells(Me.ColDO_PrimaryCodeNo.Index).Value = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                        jmlDgvDoMulti += CDec(row2.Cells(Me.ColDO_TotalQtyMulti.Index).Value)
                        jmlDgvDoDefault += CDec(row2.Cells(Me.ColDO_TotalQty.Index).Value)
                    End If
                Next
                For Each row2 As DataGridViewRow In dgvItem.Rows
                    If row2.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = .Cells(Me.ColPrimaryPOnoDoNo.Index).Value And row2.Cells(Me.ColPrimaryCodeNo.Index).Value = .Cells(Me.ColPrimaryCodeNo.Index).Value And row2.Index <> row.Index Then
                        jmlDgvItemMulti += CDec(row2.Cells(Me.ColQtyMulti.Index).Value)
                        jmlDgvItemDefault += CDec(row2.Cells(Me.ColQty.Index).Value)
                    End If
                Next
                If (CDec(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value) + jmlDgvItemDefault) >= jmlDgvDoDefault Then
                    Me.dgvItem(Me.ColQtyMulti.Index, e.RowIndex).Value = ObjCD.DigitGrouping(jmlDgvDoMulti - jmlDgvItemMulti, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                    Me.dgvItem(Me.ColQty.Index, e.RowIndex).Value = ObjCD.DigitGrouping(jmlDgvDoDefault - jmlDgvItemDefault, ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.ObjProduct.ApakahGoods(.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyMultiOnDgv(row.Index)
                        Exit Sub
                    End If
                End If

                'cek data sales order balance
                Dim JmlDo As Double = Me.ObjDO.GetJml2_Default(Me.txtPrimaryDO.Text, .Cells(Me.ColPrimaryPOnoDoNo.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlDGV = Me.GetJmlDGV_Default(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPOnoDoNo.Index).Value)
                Dim JmlDGVSendiri As Double = .Cells(Me.ColQty.Index).Value
                If JmlDo - JmlDGV < 0 Then
                    MsgBox("Maximum value you can enter for:" + CStr(.Cells(Me.ColCodeNo.Index).Value) + _
                    " and SO number: " + CStr(.Cells(Me.ColPONoDONo.Index).Value) + " is " + Me.ObjCD.DigitGrouping(CStr(JmlDo - JmlDGV + JmlDGVSendiri), Me.ObjProduct.GetAngkaDecimal(.Cells(Me.ColPrimaryCodeNo.Index).Value)) + " " + _
                    row.Cells(Me.ColUnit.Index).Value + ".", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyMultiOnDgv(row.Index)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
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
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNo(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper.Trim
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
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
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                        Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItem(Me.ColExpiredDate.Index, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(ColExpiredDate.Index, e.RowIndex).Value)
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
                                        BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value)
                                        Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = row1.Cells(Me.ColExpiredDate.Index).Value
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
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = BufferExpiryDate
                            Exit Sub
                        End If
                        Me.GabungkanBaris()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocation(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
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
                        Me.GabungkanBaris()
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Function ApakahAdaBarisYangSama() As Boolean
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

                    If row.Cells(Me.ColPONoDONo.Index).Value = row1.Cells(ColPONoDONo.Index).Value And row.Cells(Me.ColSuppDONo.Index).Value = row1.Cells(ColSuppDONo.Index).Value And _
                    row.Cells(Me.ColCodeNo.Index).Value = row1.Cells(Me.ColCodeNo.Index).Value And row.Cells(Me.ColDescription.Index).Value = row1.Cells(Me.ColDescription.Index).Value And _
                    CStr(row.Cells(Me.ColBatch.Index).Value) = CStr(row1.Cells(Me.ColBatch.Index).Value) And row.Cells(Me.ColLocation.Index).Value = _
                    row1.Cells(Me.ColLocation.Index).Value Then
                        If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                            If row.Cells(Me.ColSecondaryQty.Index).Value = row1.Cells(Me.ColSecondaryQty.Index).Value Then
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
    Private Sub GabungkanBaris()
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
                    If CStr(row.Cells(Me.ColPONoDONo.Index).Value) = CStr(row1.Cells(ColPONoDONo.Index).Value) And CStr(row.Cells(ColSuppDONo.Index).Value) = CStr(row1.Cells(Me.ColSuppDONo.Index).Value) And _
                    CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
                    CStr(row.Cells(Me.ColBatch.Index).Value) = CStr(row1.Cells(Me.ColBatch.Index).Value) And _
                    CStr(row.Cells(Me.ColLocation.Index).Value) = CStr(row1.Cells(Me.ColLocation.Index).Value) Then
                        If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                            If CDbl(row.Cells(Me.ColSecondaryQty.Index).Value) = CDbl(row1.Cells(Me.ColSecondaryQty.Index).Value) Then
                                Me.GabungkanBaris(row, row1)
                                Me.ObjCD.SetNumber(Me.dgvItem)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                            Me.GabungkanBaris(row, row1)
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

    Private Sub GabungkanBaris(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow)
        Try
            'gabungkan data qty.
            row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(row1.Cells(Me.ColQtyMulti.Index).Value), ObjProduct.GetAngkaDecimal(row1.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
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
     Optional ByVal QtyCheck As Boolean = False, Optional ByVal QtyDefault As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV_Default(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    If QtyDefault = False Then
                        Dim Max_ As String = ObjCD.DigitGrouping(CStr((JmlGudang + JmlObjLama - JmlDGVLain) * CDec(row.Cells(Me.ColRatioReference.Index).Value) / CDec(row.Cells(Me.ColRatioDefault.Index).Value)), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                        MsgBox("Maximum value you can enter is " & Max_ & " " & row.Cells(Me.ColUnitsCodeMulti.Index).Value & ".", MsgBoxStyle.Information, "Information")
                    Else
                        Dim Max_ As String = ObjCD.DigitGrouping(CStr((JmlGudang + JmlObjLama - JmlDGVLain)), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                        MsgBox("Maximum value you can enter is " & Max_ & " " & row.Cells(Me.ColUnit.Index).Value & ".", MsgBoxStyle.Information, "Information")
                    End If
                End If
                Return False
            End If
            'cek stock untuk tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(PrimaryProduct, Me.dtpDate.Value)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV_Default(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + "  minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    If QtyDefault = False Then
                        Dim Max_ As String = ObjCD.DigitGrouping(CStr((JmlGudang + JmlObjLama - JmlDGVLain) * CDec(row.Cells(Me.ColRatioReference.Index).Value) / CDec(row.Cells(Me.ColRatioDefault.Index).Value)), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                        MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + CStr(JmlGudang + JmlObjLama - JmlDGVLain) + " " + _
                        row.Cells(Me.ColUnitsCodeMulti.Index).Value + ".", MsgBoxStyle.Information, "Information")
                    Else
                        Dim Max_ As String = ObjCD.DigitGrouping(CStr((JmlGudang + JmlObjLama - JmlDGVLain)), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                        MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + CStr(JmlGudang + JmlObjLama - JmlDGVLain) + " " + _
                        row.Cells(Me.ColUnit.Index).Value + ".", MsgBoxStyle.Information, "Information")
                    End If
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
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItem.Rows.Count = 0 Then
                    Me.txtDONo.Enabled = True
                    Me.btnFindDO.Enabled = True
                    Me.txtDONo.BackColor = Me.txtPrimaryDO.BackColor
                    Me.txtDONo.ReadOnly = False
                End If
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
            If Me.TxtCodeNo.Enabled = True And Me.TxtCodeNo.ReadOnly = False And Me.TxtRemark.Enabled = True Then
                Me.TxtCodeNo.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        Try
            If Me.TxtCodeNo.Enabled = True And Me.TxtCodeNo.ReadOnly = False Then
                Me.TxtCodeNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.GotFocus
        'Try
        '    Me.TxtQtyperBatchDefault.BackColor = Core_Module.GetSMESSBackColorAll
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
        'Try
        '    Me.TxtQtyperBatchDefault.BackColor = Me.TxtPrimaryCodeNo.BackColor
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
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

    Private Sub txtDONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDONo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True And Me.txtDONo.Enabled = True Then
                If Me.txtDONo.ReadOnly = False Then
                    Me.txtDONo.BackColor = Me.BackColorAll
                Else
                    Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDONo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDONo.KeyDown
        Dim i As New FrmCariTransaksiDO
        Dim Jml As Long = 0
        i.Non_InvoiceDOOnly = False

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjDO.GetJml(Me.txtDONo.Text, Core_Module.EnumSearchCriteria.ByCode)
                If Jml = 1 Then
                    Me.txtPrimaryDO.Text = Me.ObjDO.GetPrimary(Me.txtDONo.Text)
                    'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                    Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(Me.txtPrimaryDO.Text)
                    If PrimaryStockOut <> -1 Then
                        MsgBox("This delivery order referred by stock out: " + _
                        CStr(Me.objStockOut.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
                        "Information")
                        Me.txtPrimaryDO.Text = "-1"
                        Me.txtDONo.Text = ""
                        Exit Sub
                    End If
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjDO.GetCustomerPrimary(Me.txtPrimaryDO.Text)
                    Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtRemark.Focus()
                    Me.ObjDOItem.IsiDGV_Multi(Me.DgvItemDO2, Me.txtPrimaryDO.Text)
                    Me.GabungkanBarisDO()
                    Me.FormatDGVDO()
                    'transferkan item do ke dalam datagridview
                    'Me.TransferItem(Me.txtPrimaryDO.Text)
                    If Me.dgvItem.RowCount > 0 Then
                        Me.txtDONo.ReadOnly = True
                        Me.btnFindDO.Enabled = False
                        Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                    End If

                    If Me.dgvItem.ReadOnly = False And Me.dgvItem.RowCount = 0 Then
                        If Me.ApakahBolehDONonStock = False Then
                            If Me.ApakahTransferItem3(Me.txtPrimaryDO.Text) = False Then
                                If Me.ApakahAdaItemLocationControl = True Then
                                    AddLocationControl()
                                End If
                                'transferkan item do ke dalam datagridview
                                'Me.TransferItem(Me.txtPrimaryDO.Text)
                                Me.TransferItem2(Me.txtPrimaryDO.Text, Me.IsPriorityLocation, Me.IsIncludeOtherLocation)
                            Else
                                Me.TransferItem3(Me.txtPrimaryDO.Text)
                            End If

                            If Me.dgvItem.RowCount > 0 Then
                                Me.txtDONo.ReadOnly = True
                                Me.btnFindDO.Enabled = False
                                Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                            End If
                        End If
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtDONo.Text <> "" Then
                        i.txtCari.Text = Me.txtDONo.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryTransaksi > 0 Then
                        'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                        Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(i.PrimaryTransaksi)
                        If PrimaryStockOut <> -1 Then
                            MsgBox("This delivery order referred by stock out: " + _
                            CStr(Me.objStockOut.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
                            "Information")
                            Me.txtPrimaryDO.Text = "-1"
                            Me.txtDONo.Clear()
                            Exit Sub
                        End If
                        Me.txtPrimaryDO.Text = i.PrimaryTransaksi
                        Me.txtDONo.Text = Me.ObjDO.GetDoNo(Me.txtPrimaryDO.Text)
                        Me.TxtPrimaryPihakAcuan.Text = Me.ObjDO.GetCustomerPrimary(Me.txtPrimaryDO.Text)
                        Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtRemark.Focus()
                        Me.ObjDOItem.IsiDGV_Multi(Me.DgvItemDO2, Me.txtPrimaryDO.Text)
                        Me.GabungkanBarisDO()
                        Me.FormatDGVDO()
                        'transferkan item do ke dalam datagridview
                        'Me.TransferItem(Me.txtPrimaryDO.Text)
                        If Me.dgvItem.RowCount > 0 Then
                            Me.txtDONo.ReadOnly = True
                            Me.btnFindDO.Enabled = False
                            Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                        End If

                        If Me.dgvItem.ReadOnly = False And Me.dgvItem.RowCount = 0 Then
                            If Me.ApakahBolehDONonStock = False Then
                                If Me.ApakahTransferItem3(Me.txtPrimaryDO.Text) = False Then
                                    If Me.ApakahAdaItemLocationControl = True Then
                                        AddLocationControl()
                                    End If
                                    'transferkan item do ke dalam datagridview
                                    'Me.TransferItem(Me.txtPrimaryDO.Text)
                                    Me.TransferItem2(Me.txtPrimaryDO.Text, Me.IsPriorityLocation, Me.IsIncludeOtherLocation)
                                Else
                                    Me.TransferItem3(Me.txtPrimaryDO.Text)
                                End If

                                If Me.dgvItem.RowCount > 0 Then
                                    Me.txtDONo.ReadOnly = True
                                    Me.btnFindDO.Enabled = False
                                    Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                                End If
                            End If
                        End If
                    Else
                        Me.txtDONo.Clear()
                        Me.txtPrimaryDO.Text = "-1"
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.TxtKodePihakAcuan.Text = ""
                        Me.TxtPihakAcuan.Text = ""
                        Me.txtDONo.Focus()
                        Me.DgvItemDO2.Rows.Clear()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtDONo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDONo.Leave
        Dim i As New FrmCariTransaksiDO
        i.Non_InvoiceDOOnly = False

        Try
            If Me.txtDONo.Text <> "" Then
                If Me.ObjDO.ApakahSudahAda(Me.txtDONo.Text) = True Then
                    Me.txtPrimaryDO.Text = Me.ObjDO.GetPrimary(Me.txtDONo.Text)
                    'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                    Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(Me.txtPrimaryDO.Text)
                    If PrimaryStockOut <> -1 Then
                        If Me.TxtPrimaryStockIn.Text <> PrimaryStockOut Then
                            MsgBox("This delivery order referred by stock out: " + _
                            CStr(Me.objStockOut.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
                            "Information")
                            Me.txtPrimaryDO.Text = "-1"
                            Me.txtDONo.Text = ""
                            Exit Sub
                        End If
                    End If
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjDO.GetCustomerPrimary(Me.txtPrimaryDO.Text)
                    Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtRemark.Focus()
                    Me.ObjDOItem.IsiDGV_Multi(Me.DgvItemDO2, Me.txtPrimaryDO.Text)
                    Me.GabungkanBarisDO()
                    Me.FormatDGVDO()

                    If Me.dgvItem.ReadOnly = False And Me.dgvItem.RowCount = 0 Then
                        'transferkan item do ke dalam datagridview
                        If Me.ApakahBolehDONonStock = False Then
                            If Me.ApakahTransferItem3(Me.txtPrimaryDO.Text) = False Then
                                If Me.ApakahAdaItemLocationControl = True Then
                                    AddLocationControl()
                                End If
                                'transferkan item do ke dalam datagridview
                                'Me.TransferItem(Me.txtPrimaryDO.Text)
                                Me.TransferItem2(Me.txtPrimaryDO.Text, Me.IsPriorityLocation, Me.IsIncludeOtherLocation)
                            Else
                                Me.TransferItem3(Me.txtPrimaryDO.Text)
                            End If

                            If Me.dgvItem.RowCount > 0 Then
                                Me.txtDONo.ReadOnly = True
                                Me.btnFindDO.Enabled = False
                                Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                            End If
                        End If
                    End If

                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.txtDONo.Text <> "" Then
                        i.txtCari.Text = Me.txtDONo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                        Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(i.PrimaryTransaksi)
                        If PrimaryStockOut <> -1 Then
                            MsgBox("This delivery order referred by stock out: " + _
                            CStr(Me.objStockOut.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
                            "Information")
                            Me.txtPrimaryDO.Text = "-1"
                            Me.txtDONo.Clear()
                            Exit Sub
                        End If
                        Me.txtPrimaryDO.Text = i.PrimaryTransaksi
                        Me.txtDONo.Text = Me.ObjDO.GetDoNo(Me.txtPrimaryDO.Text)
                        Me.TxtPrimaryPihakAcuan.Text = Me.ObjDO.GetCustomerPrimary(Me.txtPrimaryDO.Text)
                        Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtRemark.Focus()
                        Me.ObjDOItem.IsiDGV_Multi(Me.DgvItemDO2, Me.txtPrimaryDO.Text)
                        Me.GabungkanBarisDO()
                        Me.FormatDGVDO()

                        If Me.dgvItem.ReadOnly = False And Me.dgvItem.RowCount = 0 Then
                            'transferkan item do ke dalam datagridview
                            If Me.ApakahBolehDONonStock = False Then
                                If Me.ApakahTransferItem3(Me.txtPrimaryDO.Text) = False Then
                                    If Me.ApakahAdaItemLocationControl = True Then
                                        AddLocationControl()
                                    End If
                                    'transferkan item do ke dalam datagridview
                                    'Me.TransferItem(Me.txtPrimaryDO.Text)
                                    Me.TransferItem2(Me.txtPrimaryDO.Text, Me.IsPriorityLocation, Me.IsIncludeOtherLocation)
                                Else
                                    Me.TransferItem3(Me.txtPrimaryDO.Text)
                                End If

                                If Me.dgvItem.RowCount > 0 Then
                                    Me.txtDONo.ReadOnly = True
                                    Me.btnFindDO.Enabled = False
                                    Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                                End If
                            End If
                        End If
                    Else
                        Me.txtDONo.Clear()
                        Me.txtPrimaryDO.Text = "-1"
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.TxtKodePihakAcuan.Text = ""
                        Me.TxtPihakAcuan.Text = ""
                        Me.txtDONo.Focus()
                        Me.DgvItemDO2.Rows.Clear()
                    End If
                End If
            Else
                Me.txtPrimaryDO.Text = "-1"
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtKodePihakAcuan.Text = ""
                Me.TxtPihakAcuan.Text = ""
                Me.DgvItemDO2.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Function ApakahTransferItem3(ByVal vPrimaryDO As Double) As Boolean
        Dim DTA As New dseTransaksiStockInOut_SalesTableAdapters.DataTable2TableAdapter
        Dim TBL As New dseTransaksiStockInOut_Sales.DataTable2DataTable
        Dim row As dseTransaksiStockInOut_Sales.DataTable2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'tranfer3 bila:
            '1.semua item product yang kontrol lokasi, primary lokasinya tidak nol
            DTA.Fill(TBL, vPrimaryDO)
            For Each row In TBL.Rows
                If row.Type = 0 And row.Location_Controlled = 1 Then
                    If row.primary_location = -1 Then
                        Return False
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnFindDO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDO.Click
        Dim i As New FrmCariTransaksiDO

        Try
            i.BackColorAll = Me.BackColorAll
            'kalo non invoice do only lemot!!!
            'i.Non_InvoiceDOOnly = True

            If Me.txtDONo.Text <> "" Then
                i.txtCari.Text = Me.txtDONo.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(i.PrimaryTransaksi)
                If PrimaryStockOut <> -1 Then
                    MsgBox("This delivery order referred by stock out: " + _
                    CStr(Me.objStockOut.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
                    "Information")
                    Me.txtPrimaryDO.Text = "-1"
                    Me.txtDONo.Clear()
                    Exit Sub
                End If
                Me.txtPrimaryDO.Text = i.PrimaryTransaksi
                Me.txtDONo.Text = Me.ObjDO.GetDoNo(Me.txtPrimaryDO.Text)
                Me.TxtPrimaryPihakAcuan.Text = Me.ObjDO.GetCustomerPrimary(Me.txtPrimaryDO.Text)
                Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtRemark.Focus()
                Me.ObjDOItem.IsiDGV_Multi(Me.DgvItemDO2, Me.txtPrimaryDO.Text)
                Me.GabungkanBarisDO()
                Me.FormatDGVDO()

                If Me.dgvItem.ReadOnly = False And Me.dgvItem.RowCount = 0 Then
                    If Me.ApakahBolehDONonStock = False Then
                        If Me.ApakahTransferItem3(Me.txtPrimaryDO.Text) = False Then
                            If Me.ApakahAdaItemLocationControl = True Then
                                AddLocationControl()
                            End If
                            'transferkan item do ke dalam datagridview
                            'Me.TransferItem(Me.txtPrimaryDO.Text)
                            Me.TransferItem2(Me.txtPrimaryDO.Text, Me.IsPriorityLocation, Me.IsIncludeOtherLocation)
                        Else
                            Me.TransferItem3(Me.txtPrimaryDO.Text)
                        End If

                        If Me.dgvItem.RowCount > 0 Then
                            Me.txtDONo.ReadOnly = True
                            Me.btnFindDO.Enabled = False
                            Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                        End If
                    End If
                End If 
            Else
                Me.txtDONo.Clear()
                Me.txtPrimaryDO.Text = "-1"
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtKodePihakAcuan.Text = ""
                Me.TxtPihakAcuan.Text = ""
                Me.txtDONo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub IsiTblMultiUnit(ByVal vTBL As dseTransaksiStockInOut_Sales.DataTable3DataTable, ByVal vTBLMultiUnit_Target As dseTransaksiStockInOut_Sales.tblmaster_product_item_multiunitDataTable)
        Dim DTAMultiUnit As New dseTransaksiStockInOut_SalesTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim TBLMultiUnit As New dseTransaksiStockInOut_Sales.tblmaster_product_item_multiunitDataTable
        Dim rowMultiUnit, rowMultiUnit1 As dseTransaksiStockInOut_Sales.tblmaster_product_item_multiunitRow
        Dim row As dseTransaksiStockInOut_Sales.DataTable3Row
        Dim ListPrimarySudah As New List(Of Double)

        Try
            vTBLMultiUnit_Target.Rows.Clear()
            DTAMultiUnit.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vTBL.Rows
                If ListPrimarySudah.Contains(row.Primary_Product) = False Then
                    DTAMultiUnit.Fill(TBLMultiUnit, row.Primary_Product)
                    For Each rowMultiUnit In TBLMultiUnit.Rows
                        rowMultiUnit1 = vTBLMultiUnit_Target.Newtblmaster_product_item_multiunitRow
                        For ij As Integer = 0 To vTBLMultiUnit_Target.Columns.Count - 1
                            rowMultiUnit1.Item(ij) = rowMultiUnit.Item(ij)
                        Next
                        vTBLMultiUnit_Target.Rows.Add(rowMultiUnit1)
                    Next
                    ListPrimarySudah.Add(row.Primary_Product)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TransferItem3(ByVal vPrimaryDO As Double)
        Dim TBLMultiUnit As New dseTransaksiStockInOut_Sales.tblmaster_product_item_multiunitDataTable
        Dim rowMultiUnit As dseTransaksiStockInOut_Sales.tblmaster_product_item_multiunitRow
        Dim DTA As New dseTransaksiStockInOut_SalesTableAdapters.DataTable3TableAdapter
        Dim TBL As New dseTransaksiStockInOut_Sales.DataTable3DataTable
        Dim row As dseTransaksiStockInOut_Sales.DataTable3Row
        Dim rowDGV As DataGridViewRow
        Dim ExpiredDateStr As String
        Dim DefPrecDefault As Integer = 0
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'load item do
            DTA.Fill(TBL, vPrimaryDO)

            'isikan data multi unit
            Me.IsiTblMultiUnit(TBL, TBLMultiUnit)

            'isikan semua item do tsb ke dalam dgv
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.Rows.Clear()
            For Each row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                If row.Expiry_Date = #1/1/1950# Then
                    ExpiredDateStr = "-"
                Else
                    ExpiredDateStr = Format(row.Expiry_Date, "dd MMM yyyy")
                End If
                DefPrecDefault = row._Decimal
                If row.Primary_Units <> row.Primary_Units_Multi Then
                    TBLMultiUnit.DefaultView.RowFilter = "[Primary Product] = " + row.Primary_Product.ToString + " AND [Primary Unit] = " + row.Primary_Units_Multi.ToString
                    rowMultiUnit = TBLMultiUnit.DefaultView(0).Row
                    DefPrecMulti = rowMultiUnit.Default_Precision
                Else
                    DefPrecMulti = DefPrecDefault
                End If

                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColPONoDONo.Index).Value = row.SO_Number
                rowDGV.Cells(Me.ColSuppDONo.Index).Value = ""
                rowDGV.Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                rowDGV.Cells(Me.ColDescription.Index).Value = row.Description
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Total_Qty, DefPrecDefault)
                rowDGV.Cells(Me.ColUnit.Index).Value = row.Unit_Code
                rowDGV.Cells(Me.ColSecondaryQty.Index).Value = -1
                rowDGV.Cells(Me.ColSecondaryUnit.Index).Value = ""
                rowDGV.Cells(Me.ColBatch.Index).Value = row.Batch_No
                rowDGV.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                rowDGV.Cells(Me.ColLocation.Index).Value = row.Storage_Area_Code
                rowDGV.Cells(Me.ColWareHouse.Index).Value = row.Warehouse_Code
                rowDGV.Cells(Me.ColRemark.Index).Value = row.Remarks
                rowDGV.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = row.Primary_SO
                rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = row.primary_location
                rowDGV.Cells(Me.ColPrimaryUnits.Index).Value = row.Primary_Units
                rowDGV.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                rowDGV.Cells(Me.ColPrimaryPdO.Index).Value = -1
                rowDGV.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(row.Total_Qty_Multi, DefPrecMulti)
                rowDGV.Cells(Me.ColUnitsCodeMulti.Index).Value = row.Unit_Code_Multi
                rowDGV.Cells(Me.ColPrimaryUnitsMulti.Index).Value = row.Primary_Units_Multi
                rowDGV.Cells(Me.ColRatioDefault.Index).Value = row.Ratio_Default
                rowDGV.Cells(Me.ColRatioReference.Index).Value = row.Ratio_Reference
                If row.Primary_Units <> row.Primary_Units_Multi Then
                    rowDGV.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(row.Ratio_Reference, DefPrecMulti) _
                    & " " & row.Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(row.Ratio_Default, DefPrecDefault) & " " & row.Unit_Code
                Else
                    rowDGV.Cells(Me.ColRatioDesc.Index).Value = "-"
                End If
            Next

            Me.KosongkanItem()

            'format dgv
            Me.FormatDGV()

            'fokuskan ke baris terakhir
            'If Me.dgvItem.RowCount > 0 Then
            '    While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
            '                       Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
            '        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1

            '    End While
            'End If

            'DO tidak bisa diubah lagi
            If Me.dgvItem.Rows.Count > 0 Then
                Me.txtDONo.Enabled = False
                Me.btnFindDO.Enabled = False
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddLocationControl()
        Dim ijk As New Core_Module.FrmMasterLocationPriority
        Dim Lanjut As Boolean = False

        Try
            If Me.TxtRemark.Enabled = True Then
                If MsgBox("Add location priority control?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'isikan prioritas item before

                    ijk.BackColorAll = Me.BackColorAll
                    ijk.UserCode = Me.UserCode
                    ijk.vTipeTransaksi = ClsMasterLocationPriority.EnumTipeTransaksi.StockOutSales
                    ijk.chkIncludeOtherLocation.Visible = True
                    ijk.ShowDialog()
                    If ijk.chkIncludeOtherLocation.Checked = True Then
                        Me.IsIncludeOtherLocation = True
                    Else
                        Me.IsIncludeOtherLocation = False
                    End If
                    IsPriorityLocation = True
                Else
                    Me.IsPriorityLocation = False
                End If

                Me.ObjLocationPriority.GetLocationPrioriry(Me.TBLLocationPriority, Me.UserCode, ClsMasterLocationPriority.EnumTipeTransaksi.StockOutSales)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk = Nothing
        End Try
    End Sub

    Private Function ApakahAdaItemLocationControl() As Boolean
        Dim rowDGV As New DataGridViewRow

        Try

            For Each rowDGV In Me.DgvItemDO2.Rows
                With rowDGV
                    If Me.ObjProduct.ApakahLocationControlled(.Cells(Me.ColDO2_PrimaryCodeNo.Index).Value) = True Then
                        Return True
                    End If
                End With
            Next

            Return False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Function

    Private Sub GabungkanBarisDO()
        Dim row, rowTarget As DataGridViewRow
        Dim Sesuai As Boolean = True
        Dim i, j As Integer

        Try
            i = 0
            j = 0
            Me.DgvItemDO2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            While i <= Me.DgvItemDO2.RowCount - 2
                row = Me.DgvItemDO2.Rows(i)
                j = i + 1

                While j <= Me.DgvItemDO2.RowCount - 1
                    Sesuai = True

                    rowTarget = Me.DgvItemDO2.Rows(j)
                    'cocokan setiap kolom, kecuali kolom no dan kolom qty
                    For k As Integer = 1 To Me.DgvItemDO2.ColumnCount - 1
                        If k <> Me.ColDO_TotalQtyMulti.Index Then
                            If row.Cells(k).Value <> rowTarget.Cells(k).Value Then
                                Sesuai = False
                                Exit For
                            End If
                        End If
                    Next

                    If Sesuai = True Then
                        'gabungkan baris
                        row.Cells(Me.ColDO_TotalQtyMulti.Index).Value = CDec(row.Cells(Me.ColDO_TotalQtyMulti.Index).Value) + CDec(rowTarget.Cells(Me.ColDO_TotalQtyMulti.Index).Value)
                        Me.SetQtyDefaultOnDgvDO2(row.Index)
                        Me.DgvItemDO2.Rows.Remove(rowTarget)
                    Else
                        j += 1
                    End If
                End While

                i += 1
            End While

            Me.DgvItemDO2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GabungkanBarisDO2()
        Dim row, rowTarget As DataGridViewRow
        Dim Sesuai As Boolean = True
        Dim i, j As Integer

        Try
            i = 0
            j = 0
            Me.dgvItemDO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            While i <= Me.dgvItemDO.RowCount - 2
                row = Me.dgvItemDO.Rows(i)
                j = i + 1

                While j <= Me.dgvItemDO.RowCount - 1
                    Sesuai = True

                    rowTarget = Me.dgvItemDO.Rows(j)
                    'cocokan setiap kolom, kecuali kolom no dan kolom qty
                    For k As Integer = 1 To Me.dgvItemDO.ColumnCount - 1
                        If k <> Me.ColDO2_TotalQty.Index Then
                            If row.Cells(k).Value <> rowTarget.Cells(k).Value Then
                                Sesuai = False
                                Exit For
                            End If
                        End If
                    Next

                    If Sesuai = True Then
                        'gabungkan baris
                        row.Cells(Me.ColDO2_TotalQtyMulti.Index).Value = CDec(row.Cells(Me.ColDO2_TotalQtyMulti.Index).Value) + CDec(rowTarget.Cells(Me.ColDO2_TotalQtyMulti.Index).Value)
                        Me.SetQtyDefaultOnDgvDO(row.Index)
                        Me.dgvItemDO.Rows.Remove(rowTarget)
                    Else
                        j += 1
                    End If
                End While

                i += 1
            End While

            Me.dgvItemDO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini melakukan transfer secara langsung item do ke dalam item finish goods
    ''' </summary>
    ''' <remarks></remarks>
    'Private Sub TransferItem(ByVal vPrimaryDO As Double)
    '    Dim Lanjut As Boolean = True
    '    Dim Lanjut2 As Boolean = True
    '    Dim JmlGudang, JmlObjLama, JmlFree, JmlDGV, JmlTransfer As Double
    '    Dim ExpiredDateStr As String
    '    Dim Index As Integer
    '    Dim row As DataGridViewRow
    '    Dim objStok As New ClsTransaksiStockCard
    '    Dim StopLooping As Boolean = False

    '    Try
    '        LewatiFocusBatch = True
    '        'isikan tabel do item
    '        DTAItemDO.Fill(Me.TBLItemDO, vPrimaryDO)
    '        'Me.ObjLocationPriority.GetLocationPrioriry(Me.TBLLocationPriority, Me.UserCode, ClsMasterLocationPriority.EnumTipeTransaksi.StockOutSales)


    '        'syarat agar bisa transfer langsung adalah:
    '        For Each rowItemDO In TBLItemDO.Rows
    '            'rowItemDO = TBLItemDO.Rows(i)
    '            Lanjut = True
    '            With rowItemDO
    '                .Primary_Product = .Primary_Product
    '                ''1. barang tersebut bukan kontrol lokasi
    '                'If .Location_Controlled = 1 Then
    '                '     Lanjut = False
    '                'End If

    '                ''2. jika kontrol batch no dan atau ED maka di do juga harus ada batchnya dan expiry datenya
    '                'If Lanjut = True Then
    '                '    If .Batch_No_Controlled = 1 Then
    '                '        If .Batch_No = "" Then
    '                '            Lanjut = False
    '                '        End If
    '                '    End If
    '                'End If
    '                'If Lanjut = True Then
    '                '    If .Expiry_Date_Controlled = True Then
    '                '        If .Expiry_Date = #1/1/1950# Then
    '                '            Lanjut = False
    '                '        End If
    '                '    End If
    '                'End If

    '                If .Location_Controlled = 1 OrElse .Batch_No_Controlled = 1 OrElse .Expiry_Date_Controlled = 1 Then
    '                    '**************************************************************************************************************************************
    '                    '**************************************************************************************************************************************
    '                    'PERUBAHAN SISTEM INPUT DATA, REQUEST PAK IWAN LANGSUNG TANGGAL 31 JULI 2016 -- PERBAIKAN 03 AGUSTUS 2016
    '                    'SEBELUMNYA JIKA PRODUCT YBS LOCATION CONTROL ATAU ADA BATCH CONTROLNYA, MAKA TIDAK OTOMATIS MENGISI KE DGV, SKRG DIBUAT OTOMATIS
    '                    '**************************************************************************************************************************************
    '                    '1. barang tersebut bukan kontrol lokasi
    '                    If .Location_Controlled = 1 Then
    '                        Lanjut = False
    '                        Lanjut2 = False
    '                    End If

    '                    '2. jika kontrol batch no dan atau ED maka di do juga harus ada batchnya dan expiry datenya
    '                    If Lanjut = True Then
    '                        If .Batch_No_Controlled = 1 Then
    '                            If .Batch_No = "" Then
    '                                Lanjut = False
    '                                Lanjut2 = False
    '                            End If
    '                        End If
    '                    End If
    '                    If Lanjut = True Then
    '                        If .Expiry_Date_Controlled = True Then
    '                            If .Expiry_Date = #1/1/1950# Then
    '                                Lanjut = False
    '                            End If
    '                        End If
    '                    End If
    '                End If

    '                If Lanjut2 = False Then
    '                    StopLooping = Me.IsiItemWithLocationControl(.Primary_Product, .Total_Qty_Multi)
    '                    If StopLooping = True Then
    '                        Exit Sub
    '                    End If
    '                End If

    '                If Lanjut2 = True AndAlso Lanjut = True Then
    '                    '3. barang tersebut stock-nya ada --> ini sudah dipastikan ketika akan membuat do, 
    '                    'tapi perlu dicek kembali, bisa saja sudah ada stock out adjust yang dilakukan
    '                    JmlGudang = Me.ObjProduct.GetJml3(.Primary_Product, .Secondary_Qty, -1, .Batch_No)
    '                    'perhitungan menggunakan nilai jumlah dgv yang baru
    '                    JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, .Primary_Product, _
    '                    .Secondary_Qty, -1, .Batch_No, -1, .Expiry_Date)
    '                    JmlDGV = Me.GetJmlDGV_Default(.Primary_Product, .Primary_SO, .Batch_No)
    '                    JmlFree = JmlGudang + JmlObjLama - JmlDGV

    '                    '4. masukkan ke dalam dgv, jika jmlfree > jml item do
    '                    JmlTransfer = .Total_Qty - JmlDGV
    '                    If JmlTransfer <= JmlFree And JmlTransfer > 0 Then
    '                        Index = Me.GetIndex(.Primary_Product, -1, .Batch_No, .Secondary_Qty, -1, .Expiry_Date, .Primary_SO)

    '                        If Index = -1 Then
    '                            'tambahkan baris baru
    '                            If .Expiry_Date = #1/1/1950# Then
    '                                ExpiredDateStr = "-"
    '                            Else
    '                                ExpiredDateStr = Format(.Expiry_Date, "dd MMM yyyy")
    '                            End If

    '                            'Me.dgvItem.Rows.Add("", .SO_Number, "", .Product_Code, .Description, JmlTransfer, _
    '                            '.Unit_Code, .Secondary_Qty, .Sec_Unit_Code, .Batch_No, _
    '                            'ExpiredDateStr, "", "", "", _
    '                            '.Primary_SO, .Primary_Product, -1, .Primary_Units, _
    '                            '.Primary_Secondary_Units, -1)
    '                            Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(.Primary_Product)
    '                            Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
    '                            Me.dgvItem.Rows.Add()
    '                            row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
    '                            row.Cells(Me.ColNo.Index).Value = ""
    '                            row.Cells(Me.ColPONoDONo.Index).Value = .SO_Number
    '                            row.Cells(Me.ColSuppDONo.Index).Value = ""
    '                            row.Cells(Me.ColCodeNo.Index).Value = .Product_Code
    '                            row.Cells(Me.ColDescription.Index).Value = .Description
    '                            row.Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(JmlTransfer, DefPrecDefault)
    '                            row.Cells(Me.ColUnit.Index).Value = .Unit_Code
    '                            row.Cells(Me.ColSecondaryQty.Index).Value = .Secondary_Qty
    '                            row.Cells(Me.ColSecondaryUnit.Index).Value = .Sec_Unit_Code
    '                            row.Cells(Me.ColBatch.Index).Value = .Batch_No
    '                            row.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
    '                            row.Cells(Me.ColLocation.Index).Value = ""
    '                            row.Cells(Me.ColWareHouse.Index).Value = ""
    '                            row.Cells(Me.ColRemark.Index).Value = ""
    '                            row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = .Primary_SO
    '                            row.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
    '                            row.Cells(Me.ColPrimaryLocation.Index).Value = -1
    '                            row.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
    '                            row.Cells(Me.ColPrimarySecondaryUnits.Index).Value = .Primary_Secondary_Units
    '                            row.Cells(Me.ColPrimaryPdO.Index).Value = -1
    '                            row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
    '                            row.Cells(Me.ColUnitsCodeMulti.Index).Value = .Units_Code_Multi
    '                            row.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
    '                            row.Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault)
    '                            row.Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
    '                            If .Primary_Units <> .Primary_Units_Multi Then
    '                                row.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) _
    '                                & " " & .Units_Code_Multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault) & " " & .Unit_Code
    '                            Else
    '                                row.Cells(Me.ColRatioDesc.Index).Value = "-"
    '                            End If
    '                        Else
    '                            'tambahakn qty
    '                            row = Me.dgvItem.Rows(Index)
    '                            row.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(JmlTransfer), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
    '                            Me.SetQtyDefaultOnDgv(Index)
    '                        End If
    '                    End If
    '                End If
    '            End With
    '        Next
    '        'format dgv
    '        Me.FormatDGV()
    '        'set ulang nomor
    '        Me.ObjCD.SetNumber(Me.dgvItem)
    '        LewatiFocusBatch = False
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Function LoadQty(ByRef vTBLQty As dseTransaksiStockInOut_Sales.tblmaster_productqty1DataTable, ByVal vListPrimary As List(Of Double)) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.txtPrimaryDO.Text, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockInOut_LoadDataTBLQty, Me.UserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.LoadQty1(vTBLQty, vListPrimary)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Function LoadQty1(ByRef vTBLQty As dseTransaksiStockInOut_Sales.tblmaster_productqty1DataTable, ByVal vListPrimary As List(Of Double)) As Boolean
        Dim DtaQty As New dseTransaksiStockInOut_SalesTableAdapters.tblmaster_productqty1TableAdapter
        Dim TblQty As New dseTransaksiStockInOut_Sales.tblmaster_productqty1DataTable
        Dim RowQty, RowQty1 As dseTransaksiStockInOut_Sales.tblmaster_productqty1Row
        Dim DTAItem As New dseTransaksiStockInOut_SalesTableAdapters.DataTable1TableAdapter
        Dim TBLItem As New dseTransaksiStockInOut_Sales.DataTable1DataTable
        Dim rowItem As dseTransaksiStockInOut_Sales.DataTable1Row
        Dim DV As Data.DataRowView

        Try
            DtaQty.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DtaQty.Connection.ConnectionString

            If Me.baru = False Then
                DTAItem.Fill(TBLItem, Me.TxtPrimaryStockIn.Text)
            End If

            vTBLQty.Rows.Clear()
            For Each PrimaryProduct As Double In vListPrimary
                TblQty.Rows.Clear()
                DtaQty.Fill(TblQty, PrimaryProduct)

                If Me.baru = False Then
                    'tambahkan tblmaster product qty, dengan item db stock out sales
                    TBLItem.DefaultView.RowFilter = "[Primary Product] = " + PrimaryProduct.ToString
                    For Each DV In TBLItem.DefaultView
                        rowItem = DV.Row
                        TblQty.DefaultView.RowFilter = "[Primary Product] = " + rowItem.Primary_Product.ToString + _
                        " AND [Primary Storage Area] = " + rowItem.Primary_Storage_Area.ToString + _
                        " AND [Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowItem.Batch_No) + "'"

                        If TblQty.DefaultView.Count > 0 Then
                            RowQty = TblQty.DefaultView(0).Row
                            RowQty.Qty += rowItem.Qty
                        Else
                            RowQty = TblQty.Newtblmaster_productqty1Row
                            With RowQty
                                .Expiry_Date = rowItem.Expiry_Date
                                .Lot_No = rowItem.Batch_No
                                .Primary = -1
                                .Primary_Product = rowItem.Primary_Product
                                .Primary_Storage_Area = rowItem.Primary_Storage_Area
                                .Qty = rowItem.Qty
                                .Qty_Secondary = -1
                                .Storage_Area_Code = rowItem.Storage_Area_Code
                                .Urutan = 0
                                .Warehouse_Code = rowItem.Warehouse_Code
                                .Is_Transit = 0
                            End With
                            TblQty.Rows.Add(RowQty)
                        End If
                    Next
                End If

                For Each RowQty In TblQty.Rows
                    RowQty1 = vTBLQty.Newtblmaster_productqty1Row
                    If RowQty.Storage_Area_Code.Trim = "" Then
                        RowQty.Primary_Storage_Area = -1
                    End If
                    If RowQty.Warehouse_Code.Trim = "" Then
                        RowQty.Primary_Warehouse = -1
                    End If
                    If RowQty.Primary_Storage_Area <> -1 Then
                        If RowQty.Is_Transit = 0 Then
                            'jika location control, tidak boleh mengambil dari lokasi transit
                            For ij As Integer = 0 To vTBLQty.Columns.Count - 1
                                RowQty1.Item(ij) = RowQty.Item(ij)
                            Next
                            vTBLQty.Rows.Add(RowQty1)
                        End If
                    Else
                        For ij As Integer = 0 To vTBLQty.Columns.Count - 1
                            RowQty1.Item(ij) = RowQty.Item(ij)
                        Next
                        vTBLQty.Rows.Add(RowQty1)
                    End If
                Next
            Next


            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function TransferItem2(ByVal vPrimaryDO As Double, ByVal vLocationPriority As Boolean, ByVal vPickUpfromOtherLocation As Boolean) As Boolean
        Dim row As DataGridViewRow
        Dim TblQty As New dseTransaksiStockInOut_Sales.tblmaster_productqty1DataTable
        Dim RowQty As dseTransaksiStockInOut_Sales.tblmaster_productqty1Row
        Dim RowLocationPriority As DseMasterLocationPriority.tblmaster_location_priority1_Row
        Dim ListPrimaryProduct As New List(Of Double) 

        Try 
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DTAItemDO.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemDO.Fill(Me.TBLItemDO, vPrimaryDO)
            For Each rowItemDO In TBLItemDO.Rows
                If ListPrimaryProduct.Contains(rowItemDO.Primary_Product) = False Then
                    ListPrimaryProduct.Add(rowItemDO.Primary_Product)
                End If
            Next
            Me.LoadQty(TblQty, ListPrimaryProduct)
            For Each rowItemDO In TBLItemDO.Rows
                Dim TotalQtyMulti As Decimal = 0
                With rowItemDO
                    Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(.Primary_Product)
                    Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                    TblQty.DefaultView.RowFilter = "[Primary Product] = " & rowItemDO.Primary_Product.ToString & " AND [Lot No] = '" & rowItemDO.Batch_No & "'"
                    If rowItemDO.Location_Controlled = 1 And vLocationPriority Then
                        For Each DV As Data.DataRowView In TblQty.DefaultView
                            RowQty = DV.Row
                            Me.TBLLocationPriority.DefaultView.RowFilter = "[primary location] = " & RowQty.Primary_Storage_Area.ToString
                            If Me.TBLLocationPriority.DefaultView.Count > 0 Then
                                RowLocationPriority = Me.TBLLocationPriority.DefaultView(0).Row
                                RowQty.Urutan = RowLocationPriority.urutan
                            Else
                                RowQty.Urutan = 9999
                            End If
                        Next
                    End If
                    TblQty.DefaultView.Sort = "[Urutan] ASC, [Qty] DESC"
                    For Each DV As Data.DataRowView In TblQty.DefaultView
                        Dim Lanjut As Boolean = True
                        RowQty = DV.Row
                        If vLocationPriority And vPickUpfromOtherLocation = False And .Location_Controlled = 1 Then
                            Me.TBLLocationPriority.DefaultView.RowFilter = "[primary location] = " & RowQty.Primary_Storage_Area.ToString
                            If Me.TBLLocationPriority.DefaultView.Count = 0 Then
                                Lanjut = False
                            End If
                        End If
                        If Lanjut = True And rowItemDO.Total_Qty > 0 Then
                            Me.dgvItem.Rows.Add()
                            row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                            row.Cells(Me.ColNo.Index).Value = ""
                            row.Cells(Me.ColPONoDONo.Index).Value = .SO_Number
                            row.Cells(Me.ColSuppDONo.Index).Value = ""
                            row.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                            row.Cells(Me.ColDescription.Index).Value = .Description
                            If RowQty.Qty >= rowItemDO.Total_Qty Then
                                row.Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(rowItemDO.Total_Qty, DefPrecDefault)
                            Else
                                row.Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(RowQty.Qty, DefPrecDefault)
                            End If
                            row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value * .Ratio_Reference / .Ratio_Default, DefPrecMulti)
                            row.Cells(Me.ColUnit.Index).Value = .Unit_Code
                            row.Cells(Me.ColSecondaryQty.Index).Value = .Secondary_Qty
                            row.Cells(Me.ColSecondaryUnit.Index).Value = .Sec_Unit_Code
                            row.Cells(Me.ColBatch.Index).Value = .Batch_No
                            row.Cells(Me.ColExpiredDate.Index).Value = RowQty.Expiry_Date
                            row.Cells(Me.ColLocation.Index).Value = RowQty.Storage_Area_Code
                            row.Cells(Me.ColWareHouse.Index).Value = RowQty.Warehouse_Code
                            row.Cells(Me.ColRemark.Index).Value = .Remarks
                            row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = .Primary_SO
                            row.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                            row.Cells(Me.ColPrimaryLocation.Index).Value = RowQty.Primary_Storage_Area
                            row.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                            row.Cells(Me.ColPrimarySecondaryUnits.Index).Value = .Primary_Secondary_Units
                            row.Cells(Me.ColPrimaryPdO.Index).Value = -1
                            row.Cells(Me.ColUnitsCodeMulti.Index).Value = .Units_Code_Multi
                            row.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                            row.Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault)
                            row.Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                            If .Primary_Units <> .Primary_Units_Multi Then
                                row.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) _
                                & " " & .Units_Code_Multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault) & " " & .Unit_Code
                            Else
                                row.Cells(Me.ColRatioDesc.Index).Value = "-"
                            End If
                            rowItemDO.Total_Qty -= CDec(row.Cells(Me.ColQty.Index).Value)
                            If rowItemDO.Total_Qty = 0 Then
                                row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(rowItemDO.Total_Qty_Multi - TotalQtyMulti, DefPrecMulti)
                            End If
                            TotalQtyMulti += row.Cells(Me.ColQtyMulti.Index).Value
                        Else
                            Exit For
                        End If
                    Next
                End With
            Next
            Me.FormatDGV()
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End Try
    End Function

    Private Sub TransferItem(ByVal vPrimaryDO As Double)
        Dim Lanjut As Boolean = True
        Dim Lanjut2 As Boolean = True
        Dim JmlGudang, JmlObjLama, JmlFree, JmlDGV, JmlTransfer As Double
        Dim ExpiredDateStr As String
        Dim Index As Integer
        Dim row As DataGridViewRow
        Dim objStok As New ClsTransaksiStockCard
        Dim StopLooping As Boolean = False

        Try
            LewatiFocusBatch = True
            'isikan tabel do item
            DTAItemDO.Fill(Me.TBLItemDO, vPrimaryDO)

            'syarat agar bisa transfer langsung adalah:
            For Each rowItemDO In TBLItemDO.Rows
                Lanjut = True
                With rowItemDO
                    If .Location_Controlled = 1 OrElse .Batch_No_Controlled = 1 OrElse .Expiry_Date_Controlled = 1 Then
                        '1. barang tersebut bukan kontrol lokasi
                        If .Location_Controlled = 1 Then
                            Lanjut = False
                            Lanjut2 = False
                        End If

                        '2. jika kontrol batch no dan atau ED maka di do juga harus ada batchnya dan expiry datenya
                        If Lanjut = True Then
                            If .Batch_No_Controlled = 1 Then
                                If .Batch_No = "" Then
                                    Lanjut = False
                                    Lanjut2 = False
                                End If
                            End If
                        End If
                        If Lanjut = True Then
                            If .Expiry_Date_Controlled = True Then
                                If .Expiry_Date = #1/1/1950# Then
                                    Lanjut = False
                                End If
                            End If
                        End If
                    End If

                    If Lanjut2 = False Then
                        StopLooping = Me.IsiItemWithLocationControl(.Primary_Product, .Total_Qty_Multi, .Batch_No, .Primary_SO, .SO_Number, .Location_Controlled, .Batch_No_Controlled)
                        If StopLooping = True Then
                            Exit Sub
                        End If
                    End If

                    If Lanjut2 = True AndAlso Lanjut = True Then
                        '3. barang tersebut stock-nya ada --> ini sudah dipastikan ketika akan membuat do, 
                        'tapi perlu dicek kembali, bisa saja sudah ada stock out adjust yang dilakukan
                        JmlGudang = Me.ObjProduct.GetJml3(.Primary_Product, .Secondary_Qty, -1, .Batch_No)
                        'perhitungan menggunakan nilai jumlah dgv yang baru
                        JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, .Primary_Product, _
                        .Secondary_Qty, -1, .Batch_No, -1, .Expiry_Date)
                        JmlDGV = Me.GetJmlDGV_Default(.Primary_Product, .Primary_SO, .Batch_No)
                        JmlFree = JmlGudang + JmlObjLama - JmlDGV

                        '4. masukkan ke dalam dgv, jika jmlfree > jml item do
                        JmlTransfer = .Total_Qty - JmlDGV
                        If JmlTransfer <= JmlFree And JmlTransfer > 0 Then
                            Index = Me.GetIndex(.Primary_Product, -1, .Batch_No, .Secondary_Qty, -1, .Expiry_Date, .Primary_SO)

                            If Index = -1 Then
                                'tambahkan baris baru
                                If .Expiry_Date = #1/1/1950# Then
                                    ExpiredDateStr = "-"
                                Else
                                    ExpiredDateStr = Format(.Expiry_Date, "dd MMM yyyy")
                                End If

                                'Me.dgvItem.Rows.Add("", .SO_Number, "", .Product_Code, .Description, JmlTransfer, _
                                '.Unit_Code, .Secondary_Qty, .Sec_Unit_Code, .Batch_No, _
                                'ExpiredDateStr, "", "", "", _
                                '.Primary_SO, .Primary_Product, -1, .Primary_Units, _
                                '.Primary_Secondary_Units, -1)
                                Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(.Primary_Product)
                                Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                                Me.dgvItem.Rows.Add()
                                row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                                row.Cells(Me.ColNo.Index).Value = ""
                                row.Cells(Me.ColPONoDONo.Index).Value = .SO_Number
                                row.Cells(Me.ColSuppDONo.Index).Value = ""
                                row.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                                row.Cells(Me.ColDescription.Index).Value = .Description
                                row.Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(JmlTransfer, DefPrecDefault)
                                row.Cells(Me.ColUnit.Index).Value = .Unit_Code
                                row.Cells(Me.ColSecondaryQty.Index).Value = .Secondary_Qty
                                row.Cells(Me.ColSecondaryUnit.Index).Value = .Sec_Unit_Code
                                row.Cells(Me.ColBatch.Index).Value = .Batch_No
                                row.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                                row.Cells(Me.ColLocation.Index).Value = ""
                                row.Cells(Me.ColWareHouse.Index).Value = ""
                                row.Cells(Me.ColRemark.Index).Value = .Remarks
                                row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = .Primary_SO
                                row.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                                row.Cells(Me.ColPrimaryLocation.Index).Value = -1
                                row.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                                row.Cells(Me.ColPrimarySecondaryUnits.Index).Value = .Primary_Secondary_Units
                                row.Cells(Me.ColPrimaryPdO.Index).Value = -1
                                row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                                row.Cells(Me.ColUnitsCodeMulti.Index).Value = .Units_Code_Multi
                                row.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                                row.Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault)
                                row.Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                                If .Primary_Units <> .Primary_Units_Multi Then
                                    row.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) _
                                    & " " & .Units_Code_Multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault) & " " & .Unit_Code
                                Else
                                    row.Cells(Me.ColRatioDesc.Index).Value = "-"
                                End If
                            Else
                                'tambahakn qty
                                row = Me.dgvItem.Rows(Index)
                                row.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(JmlTransfer), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                                Me.SetQtyDefaultOnDgv(Index)
                            End If
                        End If
                    End If
                End With
            Next
            'format dgv
            Me.FormatDGV()
            'set ulang nomor
            Me.ObjCD.SetNumber(Me.dgvItem)
            LewatiFocusBatch = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Function IsiItemWithLocationControl(ByVal vPrimaryCodeNo As Double, ByVal vQty As Decimal) As Boolean

    '    Try
    '        '**************************************************************************************************************************************
    '        '**************************************************************************************************************************************
    '        'PERUBAHAN SISTEM INPUT DATA, REQUEST PAK IWAN LANGSUNG TANGGAL 31 JULI 2016 -- PERBAIKAN 03 AGUSTUS 2016
    '        'SEBELUMNYA JIKA PRODUCT YBS LOCATION CONTROL ATAU ADA BATCH CONTROLNYA, MAKA TIDAK OTOMATIS MENGISI KE DGV, SKRG DIBUAT OTOMATIS
    '        '**************************************************************************************************************************************
    '        If vPrimaryCodeNo > -1 Then
    '            Me.TxtPrimaryCodeNo.Text = vPrimaryCodeNo
    '            Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(vPrimaryCodeNo)
    '            Me.TxtCodeNo_Leave(Nothing, Nothing)
    '            If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.txtPrimarySO.Text <> "-1" Then
    '                Me.TxtQtyperBatchMulti.Text = vQty
    '                Me.TxtQtyperBatchMulti_Leave(Nothing, Nothing)
    '                IsikanDataSesuaiPrioritas()
    '                BtnInput_Click(Nothing, Nothing)
    '                Return False
    '            Else
    '                'Me.TxtPrimaryCodeNo.Text = "-1"
    '                Me.TxtCodeNo.Text = ""
    '                Me.TxtCodeNo_Leave(Nothing, Nothing)
    '                Return True
    '            End If
    '        Else
    '            Return True
    '        End If
    '        '**************************************************************************************************************************************

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Private Function IsiItemWithLocationControl(ByVal vPrimaryCodeNo As Double, ByVal vQty As Decimal) As Boolean
    '    Dim ListofCboLocation As New List(Of String)
    '    Dim ListOfCboBatch As New List(Of String)
    '    Dim vQtySisaTrans As Decimal = vQty
    '    Dim vQtyDgv As Decimal = 0
    '    Dim Lanjut As Boolean = True
    '    Dim JmlBarisDgvBefore As Integer = 0
    '    Try
    '        If vPrimaryCodeNo > -1 Then
    '            While Lanjut = True
    '                Me.TxtPrimaryCodeNo.Text = vPrimaryCodeNo
    '                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(vPrimaryCodeNo)
    '                Me.TxtCodeNo_Leave(Nothing, Nothing)
    '                If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.txtPrimarySO.Text <> "-1" Then
    '                    Me.UrutkanCboBatch(ListOfCboBatch)
    '                    Me.IsikanDataSesuaiPrioritas(ListofCboLocation)
    '                    If ListOfCboBatch.Contains(Me.cboBatchNo.SelectedValue) = False Then
    '                        ListOfCboBatch.Add(Me.cboBatchNo.SelectedValue)
    '                    End If
    '                    If ListofCboLocation.Contains(Me.cboLocation.SelectedValue) = False Then
    '                        ListofCboLocation.Add(Me.cboLocation.SelectedValue)
    '                    End If
    '                    Me.TxtQtyperBatchMulti.Text = Me.AvailableTransfer_ByLocationPriority(vQty - vQtyDgv)
    '                    Me.TxtQtyperBatchMulti_Leave(Nothing, Nothing)
    '                    JmlBarisDgvBefore = Me.dgvItem.RowCount - 1
    '                    Me.BtnInput_Click(Nothing, Nothing)
    '                    If JmlBarisDgvBefore = Me.dgvItem.RowCount - 1 Then
    '                        Lanjut = False
    '                    End If
    '                    If Me.ApakahQtyTransfer_SudahHabis(vQtySisaTrans, vQtyDgv) = True Then
    '                        Return False
    '                    End If
    '                Else
    '                    Me.TxtCodeNo.Text = ""
    '                    Me.TxtCodeNo_Leave(Nothing, Nothing)
    '                    Return True
    '                End If
    '            End While
    '        Else
    '            Return True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Function IsiItemWithLocationControl(ByVal vPrimaryCodeNo As Double, ByVal vQty As Decimal, ByVal vBatchNo As String, ByVal vPrimarySO As Double, ByVal vNoSo As String, _
    ByVal vLocationControl As Boolean, ByVal vBatchControl As Boolean) As Boolean
        Dim ListofCboLocation As New List(Of String)
        Dim vQtySisaTrans As Decimal = vQty
        Dim AvailableQty As Decimal = 0
        Dim JmlQtyTrans As Decimal = 0
        Dim Lanjut As Boolean = True
        Try
            If vPrimaryCodeNo > -1 Then
                While Lanjut = True
                    Me.TxtPrimaryCodeNo.Text = vPrimaryCodeNo
                    Me.txtPrimarySO.Text = vPrimarySO
                    Me.txtSalesOrderNo.Text = vNoSo
                    Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(vPrimaryCodeNo)
                    Me.TxtCodeNo_Leave(Nothing, Nothing)
                    If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.txtPrimarySO.Text <> "-1" Then
                        If vBatchControl = True Then
                            If String.IsNullOrEmpty(vBatchNo) Or IsNothing(vBatchNo) = False Then
                                Me.cboBatchNo.SelectedValue = vBatchNo
                            End If
                        End If

                        AvailableQty = 0
                        If vLocationControl = True Then
                            For i As Integer = 0 To cboLocation.Items.Count - 1
                                Me.IsikanDataSesuaiPrioritas(ListofCboLocation)
                                If ListofCboLocation.Contains(Me.cboLocation.SelectedValue) = False Then
                                    ListofCboLocation.Add(Me.cboLocation.SelectedValue)
                                End If

                                If ApakahLocationBenar(Me.cboLocation.SelectedValue) = True Then
                                    If ApakahQtyTersedia(AvailableQty) = True Then
                                        Exit For
                                    End If
                                End If
                            Next
                        Else
                            If vBatchControl = True Then
                                If ApakahQtyTersedia(AvailableQty) = False Then
                                    Return False
                                End If
                            End If
                        End If

                        If AvailableQty > 0 Then
                            If AvailableQty >= vQtySisaTrans Then
                                Me.TxtQtyperBatchMulti.Text = vQtySisaTrans
                            Else
                                Me.TxtQtyperBatchMulti.Text = AvailableQty
                            End If
                            vQtySisaTrans -= CDec(Me.TxtQtyperBatchMulti.Text)
                            Me.TxtQtyperBatchMulti_Leave(Nothing, Nothing)
                            Me.BtnInput_Click(Nothing, Nothing)
                        Else
                            Return False
                        End If

                        If vQtySisaTrans <= 0 Then
                            Return False
                        End If
                    Else
                        Me.TxtCodeNo.Text = ""
                        Me.TxtCodeNo_Leave(Nothing, Nothing)
                        Return True
                    End If
                End While
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahLocationBenar(ByVal vLocation As String) As Boolean
        Try
            If IsPriorityLocation = True Then
                If IsIncludeOtherLocation = False Then
                    TBLLocationPriority.DefaultView.RowFilter = "[Storage Area Code]='" & vLocation & "'"
                    If TBLLocationPriority.DefaultView.Count = 0 Then
                        Return False
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub IsikanDataSesuaiPrioritas()
    '    '**************************************************************************************************************************************
    '    '**************************************************************************************************************************************
    '    'PERUBAHAN SISTEM INPUT DATA, REQUEST PAK IWAN LANGSUNG TANGGAL 31 JULI 2016 -- PERBAIKAN 03 AGUSTUS 2016
    '    'SEBELUMNYA JIKA PRODUCT YBS LOCATION CONTROL ATAU ADA BATCH CONTROLNYA, MAKA TIDAK OTOMATIS MENGISI KE DGV, SKRG DIBUAT OTOMATIS
    '    '**************************************************************************************************************************************
    '    Dim rowLocation As dseMasterProduct.stock_location_viewRow
    '    Dim idx As Integer = 0
    '    Dim StartIdx As Integer = 0
    '    Dim Lanjut As Boolean = True
    '    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    '    Try
    '        'menentukan data yg diisikan berdasarkan location priority
    '        If Me.cboBatchNo.Visible = True Then
    '            'kombinasi batch no + expiry + location
    '            'pilih berdasarkan location priority, jika ada 2 item barang dengan lokasi yg sama tetapi expiry yang berbeda
    '            'maka pilih yang expiry paling dekat terlebih dahulu
    '            'jika sudah tidak termasuk location priority, maka tetap pilih yang expiry-nya palin dekat terlebih dahulu, alias tidak ada perubahan

    '            '1. catat index lama
    '            StartIdx = Me.cboBatchNo.SelectedIndex
    '            '2. start looping
    '            While Lanjut = True
    '                For Each rowPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1Row In Me.TBLLocationPriority.Rows
    '                    idx = 0
    '                    For Each DV As Data.DataRowView In Me.TBL2.DefaultView
    '                        rowLocation = DV.Row
    '                        If rowLocation.Location = rowPriority.Storage_Area_Code Then
    '                            Me.cboLocation.SelectedIndex = idx
    '                            Exit Sub
    '                        End If
    '                        idx += 1
    '                    Next
    '                Next
    '                If Me.cboBatchNo.SelectedIndex + 1 <= Me.cboBatchNo.Items.Count - 1 Then
    '                    'jika ada batch selanjutnya, maka cek untuk batch selanjutnya
    '                    Lanjut = True
    '                    Me.cboBatchNo.SelectedIndex += 1
    '                Else
    '                    'jika tidak ada batch selanjutnya maka kembalikan ke batch awal
    '                    Lanjut = False
    '                    Me.cboBatchNo.SelectedIndex = StartIdx
    '                End If
    '            End While
    '        Else
    '            'hanya location control
    '            'pilih lokasi sesuai dengan location priority, jika ada
    '            'jika tidak ada, maka pilih location berdasarkan abjad saja, alias tidak di ubah
    '            For Each rowPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1Row In Me.TBLLocationPriority.Rows
    '                idx = 0
    '                For Each DV As Data.DataRowView In Me.TBL2.DefaultView
    '                    rowLocation = DV.Row
    '                    If rowLocation.Location = rowPriority.Storage_Area_Code Then
    '                        Me.cboLocation.SelectedIndex = idx
    '                        Exit Sub
    '                    End If
    '                    idx += 1
    '                Next
    '            Next
    '        End If
    '        '**************************************************************************************************************************************
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub IsikanDataSesuaiPrioritas(ByVal ListLocation As List(Of String))
        Dim rowLocation As dseMasterProduct.stock_location_viewRow
        Dim idx As Integer = 0
        Dim StartIdx As Integer = 0
        Dim Lanjut As Boolean = True
        Dim ObjLocation As New Core_Module.clsMasterStorageArea
        Try
            'menentukan data yg diisikan berdasarkan location priority
            Dim TBLDummy As New dseMasterProduct.stock_location_viewDataTable
            Dim tPriority As New Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable

            TBLDummy = TBL2.Copy
            tPriority = Me.TBLLocationPriority.Copy
            StartIdx = Me.cboLocation.SelectedIndex
            If IsPriorityLocation = True Then
                For Each rowPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row In Me.TBLLocationPriority.Rows
                    idx = 0
                    Dim StrLoc As String = ""
                    For Each rPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row In tPriority
                        If ListLocation.Contains(rPriority.Storage_Area_Code) = False Then
                            StrLoc = rPriority.Storage_Area_Code
                            Exit For
                        End If
                    Next

                    TBLDummy.DefaultView.RowFilter = "[Location]='" & StrLoc & "'"
                    If TBLDummy.DefaultView.Count > 0 Then
                        For Each DV As Data.DataRowView In Me.TBL2.DefaultView
                            rowLocation = DV.Row
                            If rowLocation.Location = rowPriority.Storage_Area_Code Then
                                If ListLocation.Contains(rowLocation.Location) = False Then
                                    Me.cboLocation.SelectedIndex = idx
                                    Exit Sub
                                End If
                            End If
                            idx += 1
                        Next
                    Else
                        For Each ir As dseMasterProduct.stock_location_viewRow In Me.TBL2
                            If ListLocation.Contains(ir.Location) = False Then
                                Me.cboLocation.SelectedValue = ir.Location
                                Exit Sub
                            End If
                        Next
                    End If
                Next
            Else
                For Each ir As dseMasterProduct.stock_location_viewRow In Me.TBL2
                    If ListLocation.Contains(ir.Location) = False Then
                        Me.cboLocation.SelectedValue = ir.Location
                        Exit Sub
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGVDO()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Dim IsMultiUnit As Boolean = False

        Try
            'yang diformat dgv-nya:
            For Each row In Me.DgvItemDO2.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColDO_ExpiryDate.Index).Value) = False Then
                    row.Cells(Me.ColDO_ExpiryDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColDO_ExpiryDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColDO_ExpiryDate.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColDO_ExpiryDate.Index).Value = Format(CDate(row.Cells(Me.ColDO_ExpiryDate.Index).Value), "dd MMM yyyy")
                    End If
                End If

                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If IsNumeric(row.Cells(Me.ColDO_SecQty.Index).Value) = True Then
                    If row.Cells(ColDO_SecQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                If row.Cells(ColDO_SecQty.Index).Value = "-1" Then
                    row.Cells(ColDO_SecQty.Index).Value = "-"
                    row.Cells(ColDO_SecUnit.Index).Value = "-"
                End If

                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColDO_BatchNo.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If

                '1.e. tandai bila ada lokasi
                If row.Cells(Me.ColDO_PrimaryLocation.Index).Value <> -1 Then
                    TidakAdaLokasi = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaLokasi = True Then
                Me.ColLocation.Visible = False
                Me.ColWareHouse.Visible = False
            Else
                Me.ColLocation.Visible = True
                Me.ColWareHouse.Visible = True
            End If
            If TidakAdaSecondary = True Then
                Me.DgvItemDO2.Columns(Me.ColDO_SecQty.Index).Visible = False
                Me.DgvItemDO2.Columns(Me.ColDO_SecUnit.Index).Visible = False
            Else
                Me.DgvItemDO2.Columns(ColDO_SecQty.Index).Visible = True
                Me.DgvItemDO2.Columns(ColDO_SecUnit.Index).Visible = True
            End If
            '3. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.DgvItemDO2.Columns(Me.ColDO_BatchNo.Index).Visible = False
            Else
                Me.DgvItemDO2.Columns(Me.ColDO_BatchNo.Index).Visible = True
            End If
            '4. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.DgvItemDO2.Columns(Me.ColDO_ExpiryDate.Index).Visible = False
            Else
                Me.DgvItemDO2.Columns(Me.ColDO_ExpiryDate.Index).Visible = True
            End If

            For Each RowDgv As DataGridViewRow In Me.DgvItemDO2.Rows
                With RowDgv
                    If .Cells(Me.ColDO_PrimaryUnit.Index).Value <> .Cells(Me.ColDO_PrimaryUnitsMulti.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.ColDO_TotalQty.Visible = False
                Me.ColDO_Unit.Visible = False
                Me.ColDO_RatioDesc.Visible = False
            Else
                Me.ColDO_TotalQty.Visible = True
                Me.ColDO_Unit.Visible = True
                Me.ColDO_RatioDesc.Visible = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSalesOrderNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesOrderNo.GotFocus, TextBox2.GotFocus
        Try
            If Me.txtPrimaryDO.Text <> "" And Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryDO.Text = -1 And Me.txtPrimaryDO.Text <> "" Then
                    Me.txtDONo.Focus()
                    MsgBox("To start entering item data, you must specify delivery order data first.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.txtSalesOrderNo.Enabled = True And Me.txtSalesOrderNo.ReadOnly = False And Me.TxtRemark.Enabled = True Then
                Me.txtSalesOrderNo.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSalesOrderNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSalesOrderNo.KeyDown, TextBox2.KeyDown
        Dim i As New FrmCariTransaksiSO
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                If Me.ObjSO.ApakahSudahAda(Me.txtSalesOrderNo.Text) = True Then
                    Me.txtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.txtSalesOrderNo.Text)
                    Me.TxtCodeNo.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtSalesOrderNo.Text <> "" Then
                        i.txtCari.Text = Me.txtSalesOrderNo.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryTransaksi > 0 Then
                        Me.txtPrimarySO.Text = i.PrimaryTransaksi
                        Me.txtSalesOrderNo.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.txtSalesOrderNo.Clear()
                        Me.txtPrimarySO.Text = "-1"
                        Me.txtSalesOrderNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSalesOrderNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesOrderNo.Leave, TextBox2.Leave
        Dim i As New FrmCariTransaksiSO

        Try
            If Me.txtSalesOrderNo.Text <> "" Then
                If Me.ObjSO.ApakahSudahAda(Me.txtSalesOrderNo.Text) = True Then
                    Me.txtPrimarySO.Text = Me.ObjSO.GetPrimary(Me.txtSalesOrderNo.Text)
                    Me.TxtCodeNo.Focus()
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.txtSalesOrderNo.Text <> "" Then
                        i.txtCari.Text = Me.txtSalesOrderNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        Me.txtPrimarySO.Text = i.PrimaryTransaksi
                        Me.txtSalesOrderNo.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                        Me.TxtCodeNo.Focus()
                    Else
                        Me.txtSalesOrderNo.Clear()
                        Me.txtPrimarySO.Text = "-1"
                        Me.txtSalesOrderNo.Focus()
                    End If
                End If
            Else
                Me.txtPrimarySO.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesOrder.Click, Button1.Click
        Dim i As New FrmCariTransaksiSO

        Try
            If Me.txtPrimaryDO.Text <> "" Then
                If Me.txtPrimaryDO.Text = -1 And Me.txtPrimaryDO.Text <> "" Then
                    Me.txtDONo.Focus()
                    MsgBox("To start entering item data, you must specify delivery order data first.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            i.BackColorAll = Me.BackColorAll

            If Me.txtSalesOrderNo.Text <> "" Then
                i.txtCari.Text = Me.txtSalesOrderNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.txtPrimarySO.Text = i.PrimaryTransaksi
                Me.txtSalesOrderNo.Text = Me.ObjSO.GetNoSO(Me.txtPrimarySO.Text)
                Me.TxtCodeNo.Focus()
            Else
                Me.txtSalesOrderNo.Clear()
                Me.txtPrimarySO.Text = "-1"
                Me.txtSalesOrderNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSalesOrderNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesOrderNo.LostFocus, TextBox2.LostFocus
        Try
            If Me.txtSalesOrderNo.Enabled = True And Me.txtSalesOrderNo.ReadOnly = False Then
                Me.txtSalesOrderNo.BackColor = Me.txtPrimaryDO.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub FrmTransaksiStockOutSales_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
    '    If Me.FrmDODetail.Visible = True Then
    '        Me.FrmDODetail.Hide()
    '    End If
    'End Sub

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

    Private Sub txtDONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDONo.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Or Me.txtDONo.Enabled = True Then
                If Me.txtDONo.ReadOnly = False Then
                    Me.txtDONo.BackColor = Me.txtPrimaryDO.BackColor
                Else
                    Me.txtDONo.BackColor = Me.txtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Leave
        Try
            Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
            If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtRatioDefault.Text <> "" And Me.TxtRatioReference.Text <> "" Then
                SetQtyMulti()

                ''jika lebih atau sama dengan batas qty di DO, udah default/multi dengan sisanya
                Dim jmlDgvDoMulti As Decimal = 0
                Dim jmlDgvDoDefault As Decimal = 0
                Dim jmlDgvItemMulti As Decimal = 0
                Dim jmlDgvItemDefault As Decimal = 0
                For Each row As DataGridViewRow In DgvItemDO2.Rows
                    If row.Cells(Me.ColDO_PrimarySO.Index).Value = Me.txtPrimarySO.Text And row.Cells(Me.ColDO_PrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And _
                        row.Cells(Me.ColDO_BatchNo.Index).Value = Me.cboBatchNo.Text Then
                        jmlDgvDoMulti += CDec(row.Cells(Me.ColDO_TotalQtyMulti.Index).Value)
                        jmlDgvDoDefault += CDec(row.Cells(Me.ColDO_TotalQty.Index).Value)
                    End If
                Next
                For Each row As DataGridViewRow In dgvItem.Rows
                    If row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = Me.txtPrimarySO.Text And row.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And _
                        row.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text Then
                        jmlDgvItemMulti += CDec(row.Cells(Me.ColQtyMulti.Index).Value)
                        jmlDgvItemDefault += CDec(row.Cells(Me.ColQty.Index).Value)
                    End If
                Next
                If (CDec(Me.TxtQtyperBatchDefault.Text) + jmlDgvItemDefault) >= jmlDgvDoDefault Then
                    Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(jmlDgvDoMulti - jmlDgvItemMulti, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping(jmlDgvDoDefault - jmlDgvItemDefault, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                Dim i As New FrmLaporanFormatStockInOut_Bundling

                i.primary = Me.TxtPrimaryStockIn.Text
                i.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutSales
                i.ShowDialog()

                i = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.buffer = TxtQtyperBatchMulti.Text
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatchMulti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                    Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.SetQtyDefault()
                    If CDec(Me.TxtQtyperBatchDefault.Text) = 0 And CDec(Me.TxtQtyperBatchMulti.Text) > 0 Then
                        Me.TxtQtyperBatchMulti.Text = Me.buffer
                        MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.SetQtyDefault()
                        If Me.LewatiFocusBatch = False Then
                            Me.TxtQtyperBatchMulti.Focus()
                        End If
                    End If

                    ''jika lebih atau sama dengan batas qty di DO, udah default/multi dengan sisanya
                    Dim jmlDgvDoMulti As Decimal = 0
                    Dim jmlDgvDoDefault As Decimal = 0
                    Dim jmlDgvItemMulti As Decimal = 0
                    Dim jmlDgvItemDefault As Decimal = 0
                    For Each row As DataGridViewRow In DgvItemDO2.Rows
                        If row.Cells(Me.ColDO_PrimarySO.Index).Value = Me.txtPrimarySO.Text And row.Cells(Me.ColDO_PrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And _
                            row.Cells(Me.ColDO_BatchNo.Index).Value = Me.cboBatchNo.Text Then
                            jmlDgvDoMulti += CDec(row.Cells(Me.ColDO_TotalQtyMulti.Index).Value)
                            jmlDgvDoDefault += CDec(row.Cells(Me.ColDO_TotalQty.Index).Value)
                        End If
                    Next
                    For Each row As DataGridViewRow In dgvItem.Rows
                        If row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = Me.txtPrimarySO.Text And row.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And _
                           row.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text Then
                            jmlDgvItemMulti += CDec(row.Cells(Me.ColQtyMulti.Index).Value)
                            jmlDgvItemDefault += CDec(row.Cells(Me.ColQty.Index).Value)
                        End If
                    Next
                    If (CDec(Me.TxtQtyperBatchMulti.Text) + jmlDgvItemMulti) >= jmlDgvDoMulti Then
                        Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping(jmlDgvDoDefault - jmlDgvItemDefault, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                        Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(jmlDgvDoMulti - jmlDgvItemMulti, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    End If

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub SetQtyMultiOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColQtyMulti.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQty.Index, vRowIdx).Value) * _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value) / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgvDO2(ByVal vRowIdx As Integer)
        Try
            Me.DgvItemDO2.Item(Me.ColDO_TotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.DgvItemDO2.Item(Me.ColDO_TotalQtyMulti.Index, vRowIdx).Value) * _
                                                            CDec(Me.DgvItemDO2.Item(Me.ColDO_RatioDefault.Index, vRowIdx).Value / _
                                                            CDec(Me.DgvItemDO2.Item(Me.ColDO_RatioReference.Index, vRowIdx).Value)), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.DgvItemDO2.Item(Me.ColDO_PrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgvDO(ByVal vRowIdx As Integer)
        Try
            Me.dgvItemDO.Item(Me.ColDO2_TotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItemDO.Item(Me.ColDO2_TotalQtyMulti.Index, vRowIdx).Value) * _
                                                            CDec(Me.DgvItemDO2.Item(Me.ColDO2_RatioDefault.Index, vRowIdx).Value / _
                                                            CDec(Me.DgvItemDO2.Item(Me.ColDO2_RatioReference.Index, vRowIdx).Value)), _
                                                            Me.ObjProduct.GetAngkaDecimal(Me.dgvItemDO.Item(Me.ColDO2_PrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetMultiUnits()
        Dim Dta As New dseTransaksiStockInOutTableAdapters.DataTable11TableAdapter
        Dim Tbl As New dseTransaksiStockInOut.DataTable11DataTable
        Dim Row As dseTransaksiStockInOut.DataTable11Row
        Try

            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Me.txtPrimaryDO.Text = "" Then
                Me.txtPrimaryDO.Text = "-1"
            End If
            If Me.txtPrimarySO.Text = "" Then
                Me.txtPrimarySO.Text = "-1"
            End If
            If Me.TxtPrimaryCodeNo.Text = "" Then
                Me.TxtPrimaryCodeNo.Text = "-1"
            End If
            Dta.Fill(Tbl, Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)
            If Tbl.Count > 0 Then
                Row = Tbl.Rows(0)
                Me.txtPrimaryUnitsDefault.Text = Row.Primary_Units
                Me.txtPrimaryUnitsMulti.Text = Row.Primary_Units_Multi
                Me.TxtRatioDefault.Text = ObjCD.DigitGrouping(Row.Ratio_Default, ObjProduct.GetAngkaDecimal(Row.Primary_Product))
                Me.TxtRatioReference.Text = ObjCD.DigitGrouping(Row.Ratio_Reference, ObjProduct.GetAngkaDecimal(Row.Primary_Product, Row.Primary_Units_Multi))
                Me.lblUnitsDefault.Text = Row.Units_Code_Default
                Me.lblUnitsMulti.Text = Row.Units_Code_Multi
            Else
                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.txtPrimaryUnitsMulti.Text = "-1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtRatioReference.Text = "1"
                Me.lblUnitsDefault.Text = "Units"
                Me.lblUnitsMulti.Text = "Units"
            End If

            'Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
            Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
            Me.SetQtyDefault()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimarySO_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrimarySO.TextChanged
        Try
            Me.SetMultiUnits()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            Me.SetMultiUnits()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyMulti()
        Try
            If IsNumeric(Me.TxtQtyperBatchDefault.Text) = True And IsNumeric(Me.TxtRatioDefault.Text) = True Then
                Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchDefault.Text) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_StockOutSales_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core, _
    ByVal vTblPrimarySO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryDO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable) As Boolean
        Dim TblTemp As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.objStockOut.ResetObject(Me.dgvItem)
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales

            Me.objStockOut.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.TblItemCustom)
            Me.objStockOut.PrimaryDOSales = ObjCutOff_Core.GetPrimaryUpdate(Me.objStockOut.PrimaryDOSales, vTblPrimaryDO)
            Me.UserCode = Me.objStockOut.CreatedBy

            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value) = False And rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value <> "-1" Then
                    TblTemp.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value)
                    If TblTemp.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value), vTblPrimarySO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value), PrimaryUpdate, TblTemp)
                        rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOnoDoNo.Index).Value), TblTemp)
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function FromCutOff_StockOutSales_Save() As Double
        Try
            Me.baru = True
            Me.objStockOut.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, -1, False, False, True)
            Return Me.objStockOut.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function ApakahQtyTersedia(ByRef vQtyTersedia As Decimal) As Boolean
        Dim ObjSC As New ClsTransaksiStockCard
        Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
        Try
            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNo.Text) = True Then
                JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.txtPrimaryPDO.Text, _
                Me.DtpExpired.Value)
                JmlDGV = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
                Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchDefault.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    vQtyTersedia = JmlGudang + JmlObjLama - JmlDGV
                Else
                    vQtyTersedia = ObjCD.DigitGrouping(CStr((JmlGudang + JmlObjLama - JmlDGV) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text)), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                End If
            End If
            If vQtyTersedia > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(Me.LinkLabel1, 0, 10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiStockOutSales_AddByCustomProduct
        Try
            If Me.txtPrimaryDO.Text <> "" Then
                If Me.txtPrimaryDO.Text = -1 And Me.txtPrimaryDO.Text <> "" Then
                    MsgBox("To start entering item data, you must specify delivery order data first.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                i.UserCode = Me.UserCode
                i.BackColorAll = Me.BackColorAll
                i.TanggalTransaksi = Me.dtpDate.Value
                i.txtPrimaryDO.Text = Me.txtPrimaryDO.Text
                i.PrimarystockIn = Me.TxtPrimaryStockIn.Text
                i.TblItemCustom = Me.TblItemCustom.Copy 'sync
                If Me.TxtRemark.Enabled = True Then
                    i.AktifForm = True
                Else
                    i.AktifForm = False
                End If

                If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Me.TblItemCustom = i.TblItemCustom.Copy
                    Me.IsikanDgvFromTblDummy()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturHakAksesFormCustom()
        Dim ObjUser As New Core_Module.ClsUser
        Try
            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutSales_AddByCustomProduct") = False Then
                Me.gboItem.Enabled = True
                'Me.ContextMenuStrip1.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
                Me.LinkLabel1.Visible = False
            Else
                Me.dgvItem.ReadOnly = True
                Me.gboItem.Enabled = False
                'Me.ContextMenuStrip1.Visible = True
                Me.dgvItem.AllowUserToDeleteRows = False
                Me.LinkLabel1.Visible = True
            End If
            Me.LinkLabel1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsikanDgvFromTblDummy() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowC As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemRow
        Dim ObjP As New Core_Module.clsMasterProduct
        Try
            Me.dgvItem.Rows.Clear()
            For Each rowC In TblItemCustom.Rows
                Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(rowC.Primary_Product)
                Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(rowC.Primary_Product, rowC.Primary_Multi_Units)
                Dim Index As Integer = Me.GetIndex2(rowC.Primary_Product, rowC.Primary_Storage_Area, rowC.Batch_No, _
                -1, rowC.Primary_PdO, rowC.Expiry_Date, rowC.Primary_PO_Purchase)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(rowC.Qty_Multi_Unit), DefPrecMulti)
                    'Me.SetQtyDefaultOnDgv(row.Index)
                    'row.Cells(Me.ColQty.Index).Value = CDec(row.Cells(Me.ColQty.Index).Value) + rowC.Qty
                    'row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value, ObjProduct.GetAngkaDecimal(rowC.Primary_Product, rowC.Primary_Multi_Units))
                    row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQty.Index).Value) + rowC.Qty, DefPrecDefault)
                    If rowC.Is_Kain = True Then
                        row.Cells(Me.ColRatioDefault.Index).Value = row.Cells(Me.ColQty.Index).Value
                        row.Cells(Me.ColRatioReference.Index).Value = row.Cells(Me.ColQtyMulti.Index).Value
                        row.Cells(Me.ColRatioDesc.Index).Value = row.Cells(Me.ColQtyMulti.Index).Value & " " & rowC.unit_multi & " = " & row.Cells(Me.ColQty.Index).Value & " " & rowC.unit
                    End If

                    If String.IsNullOrEmpty(rowC.Remarks) = False Then
                        If String.IsNullOrEmpty(row.Cells(Me.ColRemark.Index).Value) = False Then
                            row.Cells(Me.ColRemark.Index).Value = row.Cells(Me.ColRemark.Index).Value & ", " & rowC.Remarks
                        Else
                            row.Cells(Me.ColRemark.Index).Value = rowC.Remarks
                        End If
                    End If
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If rowC.Expiry_Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(rowC.Expiry_Date, "dd MMM yyyy")
                    End If

                    Me.dgvItem.Rows.Add()
                    Dim row As DataGridViewRow
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    row.Cells(Me.ColNo.Index).Value = ""
                    row.Cells(Me.ColPONoDONo.Index).Value = rowC.po_so_number
                    row.Cells(Me.ColSuppDONo.Index).Value = ""
                    row.Cells(Me.ColCodeNo.Index).Value = rowC.Product_Code
                    row.Cells(Me.ColDescription.Index).Value = rowC.Description
                    row.Cells(Me.ColQty.Index).Value = rowC.Qty
                    row.Cells(Me.ColUnit.Index).Value = rowC.unit
                    row.Cells(Me.ColSecondaryQty.Index).Value = -1
                    row.Cells(Me.ColSecondaryUnit.Index).Value = ""
                    row.Cells(Me.ColBatch.Index).Value = rowC.Batch_No
                    row.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                    row.Cells(Me.ColLocation.Index).Value = rowC.Storage_Area_Code
                    row.Cells(Me.ColWareHouse.Index).Value = rowC.Warehouse_Code
                    row.Cells(Me.ColRemark.Index).Value = rowC.Remarks
                    row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = rowC.Primary_PO_Purchase
                    row.Cells(Me.ColPrimaryCodeNo.Index).Value = rowC.Primary_Product
                    row.Cells(Me.ColPrimaryLocation.Index).Value = rowC.Primary_Storage_Area
                    row.Cells(Me.ColPrimaryUnits.Index).Value = rowC.Primary_Units
                    row.Cells(Me.ColPrimarySecondaryUnits.Index).Value = rowC.Primary_Secondary_Units
                    row.Cells(Me.ColPrimaryPdO.Index).Value = rowC.Primary_PdO
                    row.Cells(Me.ColQtyMulti.Index).Value = rowC.Qty_Multi_Unit
                    row.Cells(Me.ColUnitsCodeMulti.Index).Value = rowC.unit_multi
                    row.Cells(Me.ColPrimaryUnitsMulti.Index).Value = rowC.Primary_Multi_Units
                    row.Cells(Me.ColRatioDefault.Index).Value = rowC.Ratio_Default
                    row.Cells(Me.ColRatioReference.Index).Value = rowC.Rasio
                    If rowC.Primary_Units <> rowC.Primary_Multi_Units Then
                        row.Cells(Me.ColRatioDesc.Index).Value = rowC.Rasio _
                        & " " & rowC.unit_multi & " = " & rowC.Ratio_Default & " " & rowC.unit
                    Else
                        row.Cells(Me.ColRatioDesc.Index).Value = "-"
                    End If

                    'Me.dgvItem.Rows.Add("", Me.txtSalesOrderNo.Text, "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatchDefault.Text, _
                    'Me.lblUnitsDefault.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.cboBatchNo.Text, _
                    'ExpiredDateStr, Me.cboLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnitsDefault.Text, _
                    'Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If
                'kosongkan gbo item
                'Me.KosongkanItem()
                'If LewatiFocusBatch = False Then
                '    Me.TxtCodeNo.Focus()
                'End If
                'format dgv
                Me.FormatDGV()
                'fokuskan ke baris terakhir
                If LewatiFocusBatch = False Then
                    While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                                    Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If

                'DO tidak bisa diubah lagi
                'Me.txtDONo.Enabled = False
                'Me.btnFindDO.Enabled = False
            Next

            If Me.dgvItem.Rows.Count > 0 Then
                Me.txtDONo.Enabled = False
                Me.btnFindDO.Enabled = False
            Else
                Me.txtDONo.Enabled = True
                Me.btnFindDO.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex2(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
  ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date, ByVal vPrimarySO As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And SecondaryQty = -1 _
                        And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO Then
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
End Class