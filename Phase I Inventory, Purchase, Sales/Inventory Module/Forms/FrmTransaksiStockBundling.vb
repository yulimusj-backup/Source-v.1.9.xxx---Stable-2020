Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiStockBundling
    Public UserCode As String = ""
    Public ObjStockBundling As New ClsTransaksiStockBundling
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim SedangAddRM As Boolean = False
    Dim ObjItemInOut As New clsTransaksiStockInOut
    Dim ObjItemBundling As New clsTransaksiStockBundlingItem
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCustom As New CustomSMESS.ClsTransaksiStockBundling
    Dim ObjProses As New Core_Module.ClsMasterProses
    Dim ObjWorkUnit As New Core_Module.ClsMasterWorkUnit
    Dim ObjWorkShift As New Core_Module.ClsMasterWorkShift
    Dim TBLLocationPriority As New Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable
    Dim ObjLocationPriority As New Core_Module.ClsMasterLocationPriority
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjUser As New ClsUser
    Dim ObjLicense As New Core_Module.ClsLicense

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

    Dim TBL_BOM As New Core_Module.dseMasterProduct_BOM.tblmaster_product_item_bom1DataTable
    Dim row_BOM As Core_Module.dseMasterProduct_BOM.tblmaster_product_item_bom1Row
    Dim QuickAddByRefDoc As Boolean = False
    Dim QuickAddByBarcode As Boolean = False 

    Public TblItemBefCustom As New CustomSMESS.DseTransaksiStockBundling.tbltrans_custom_stock_bundling_item_beforeDataTable
    Public TblItemAftCustom As New CustomSMESS.DseTransaksiStockBundling.tbltrans_custom_stock_bundling_item_afterDataTable

    Private Sub AktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.tabB4.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.tabAfter.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabB4.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabAfter.Controls, Me.TxtPrimaryStockIn.BackColor, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItemb4.Rows.Count = 0 Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
            End If
            If Me.dgvItemAfter.Rows.Count = 0 Then
                Me.btnUpAfter.Enabled = False
                Me.btnDownAfter.Enabled = False
            End If
            Me.dgvItemb4.AllowUserToDeleteRows = True
            Me.dgvItemb4.ReadOnly = False
            Me.dgvItemAfter.AllowUserToDeleteRows = True
            Me.dgvItemAfter.ReadOnly = False
            Me.DtpExpiredb4.Enabled = False
            Me.llb_BOM.Enabled = True

            Me.SetGboItemAfter()
            Me.SetGboItemb4()
            Me.FormatDGV(Me.dgvItemAfter)
            Me.FormatDGV(Me.dgvItemb4)

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_QuickInputDocRef") = True Then
                Me.QuickAddByRefDoc = True
            Else
                Me.QuickAddByRefDoc = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_QuickInputBarcode") = True Then
                Me.QuickAddByBarcode = True
            Else
                Me.QuickAddByBarcode = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_AddByCustomProduct") = True Then
                Me.QuickAddByBarcode = True

                Me.SplitContainer2.Enabled = False
                Me.SplitContainer3.Enabled = False
                Me.dgvItemb4.ReadOnly = True
                Me.dgvItemb4.AllowUserToDeleteRows = False
                Me.dgvItemAfter.ReadOnly = True
                Me.dgvItemAfter.AllowUserToDeleteRows = False
                Me.llb_BOM.Enabled = False
            Else
                Me.QuickAddByBarcode = False

                Me.SplitContainer2.Enabled = True
                Me.SplitContainer3.Enabled = True
                Me.dgvItemb4.ReadOnly = False
                Me.dgvItemb4.AllowUserToDeleteRows = True
                Me.dgvItemAfter.ReadOnly = False
                Me.dgvItemAfter.AllowUserToDeleteRows = True

                'disini harusnya set menu add by customnya enable false
                Me.AddItemByCustom.Enabled = True
            End If

            If Me.QuickAddByBarcode = True Or Me.QuickAddByRefDoc = True Then
                Me.llbAdvance.Enabled = True
            Else
                Me.llbAdvance.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.tabB4.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.tabAfter.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            'Me.SplitContainer1.TabStop = True
            Me.CancelButton = Me.btnExit
            Me.dgvItemb4.AllowUserToDeleteRows = False
            Me.dgvItemb4.ReadOnly = True
            Me.dgvItemAfter.AllowUserToDeleteRows = False
            Me.dgvItemAfter.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabB4.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabAfter.Controls, Me.BackColor, False)
            Me.llb_BOM.Enabled = False
            Me.llbAdvance.Enabled = False
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_AddByCustomProduct") = True Then
                Me.llbAdvance.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.dtpDate.Value = Today
            Me.KosongkanItemB4()
            Me.KosongkanItemAfter()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            Me.TxtCodeProses.Clear()
            Me.TxtDescProses.Clear()
            Me.TxtCodeWorkUnit.Clear()
            Me.TxtDescWorkUnit.Clear()
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.TxtPrimaryProses.Text = "-1"
            Me.TxtPrimaryWorkUnit.Text = "-1"
            Me.TxtPrimaryWorkShift.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItemb4.Rows.Clear()
            Me.dgvItemAfter.Rows.Clear()
            Me.lblUnits2B4.Text = Me.lblUnitsb4.Text

            Me.TblItemBefCustom.Rows.Clear()
            Me.TblItemAftCustom.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItemAfter()
        Try
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNoAfter.Text = "-1"
            Me.txtPrimaryUnitsAfter.Text = "-1"
            Me.txtPrimarySecondaryUnitsAfter.Text = "-1"
            Me.txtPrimaryLocationAfter.Text = "-1"
            Me.txtPrimaryCategoryAfter.Text = "-1"
            Me.TxtQtyperBatchAfter.Text = "0"
            Me.txtSecondaryQtyAfter.Text = "-1"
            Me.txtFaktorPemberat.Text = "0"
            Me.chkBatchNoControlledAfter.Checked = False
            Me.chkExpiryDateControlledAfter.Checked = False
            Me.chkLocationControlledAfter.Checked = False
            Me.chkPdoControlledAfter.Checked = False
            Me.chkUseSecondaryUnitAfter.Checked = False
            Me.txtSecondaryQtyAfter.Visible = False
            Me.lblSecondaryUnitsAfter.Visible = False
            Me.lblUnitsAfter.Left = Me.lblSecondaryUnitsAfter.Left
            Me.TxtQtyperBatchAfter.Width = Me.lblUnitsAfter.Left - Me.TxtQtyperBatchAfter.Left - 6
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
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNob4.Text = "-1"
            Me.txtPrimaryUnitsb4.Text = "-1"
            Me.txtPrimarySecondaryUnitsb4.Text = "-1"
            Me.txtPrimaryLocationb4.Text = "-1"
            Me.txtPrimaryCategoryb4.Text = "-1"
            Me.TxtQtyperBatchb4.Text = "0"
            Me.txtSecondaryQtyb4.Text = "-1"
            Me.chkBatchNoControlledb4.Checked = False
            Me.chkExpiryDateControlledb4.Checked = False
            Me.chkLocationControlledb4.Checked = False
            Me.chkPdoControlledb4.Checked = False
            Me.chkUseSecondaryUnitb4.Checked = False
            Me.lblSecondaryUnitsb4.Visible = False
            'Me.lblUnitsb4.Left = Me.lblSecondaryUnitsb4.Left
            'Me.TxtQtyperBatchb4.Width = Me.lblUnitsb4.Left - Me.TxtQtyperBatchb4.Left - 6
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

    Private Function ApakahDataItemB4Benar() As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDescb4.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNob4.Text = ""
                Me.TxtCodeNob4.Focus()
                Return False
            End If
            '1.b. hanya boleh memasukkan goods
            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNob4.Text) = False Then
                MsgBox("Please enter only goods code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNob4.Text = ""
                Me.TxtCodeNob4.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchb4.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchb4.Focus()
                Return False
            End If
            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlledb4.Checked = True Then
                If Me.cboBatchNob4.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.cboBatchNob4.Focus()
                    Return False
                End If
            End If
            '4.a Untuk transaksi proses (in/back/out proses) no Pdo harus diisikan
            '4.b Untuk transaksi yang berkaitan dengan purchase (stock out purchase return dan stout in purchase)
            '    harus menyertakan no PO bila terhubungkan dengan software purchase
            '4.c Untuk transaksi yang berkaitan dengan sales (stock out sales dan stock in sales return) harrus
            '    menyertakan no DO bila terhubungkan dengan software sales
            '5. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlledb4.Checked = True Then
                If Me.txtPrimaryLocationb4.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, _
                    "Information")
                    Me.cboLocationb4.Focus()
                    Return False
                End If
            End If
            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlledb4.Checked = True Then
                If Me.DtpExpiredb4.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpiredb4.Focus()
                        Return False
                    End If
                End If
            End If
            '7. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnitb4.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.txtSecondaryQtyb4.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.txtSecondaryQtyb4.Focus()
                    Return False
                End If
            End If
            '8a. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlledb4.Checked = True And Me.chkExpiryDateControlledb4.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNob4.Text, Me.cboBatchNob4.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpiredb4.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.cboBatchNob4.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpiredb4.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItemb4.Rows
                    If row.Cells(7).Value = Me.cboBatchNob4.Text And Me.cboBatchNob4.Text <> "" And row.Cells(12).Value = _
                    Me.TxtPrimaryCodeNob4.Text Then
                        If CDate(row.Cells(8).Value) <> Me.DtpExpiredb4.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNob4.Text + " on " + _
                            Format(CDate(row.Cells(8).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.DtpExpiredb4.Focus()
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
            Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV, JmlObjLamaAft As Decimal
            JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockBundlingItem.GetJmlB4(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNob4.Text, _
            Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, Me.DtpExpiredb4.Value)
            JmlObjLamaAft = Me.objStockBundlingItem.GetJmlAfter(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNob4.Text, _
            Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, Me.DtpExpiredb4.Value)
            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
            Me.cboBatchNob4.Text, Me.chkBatchNoControlledb4.Checked, _
            Me.chkLocationControlledb4.Checked, Me.chkUseSecondaryUnitb4.Checked, Me.dgvItemb4)
            'JmlDGVAft = Me.GetJmlDGV(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
            'Me.cboBatchNob4.Text, Me.chkBatchNoControlledb4.Checked, _
            'Me.chkLocationControlledb4.Checked, Me.chkUseSecondaryUnitb4.Checked, Me.dgvItemAfter)
            If JmlGudang >= JmlObjLamaAft Then
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text) - JmlObjLamaAft
            Else
                JmlSetelahTambah = JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text)
            End If


            If JmlSetelahTambah < 0 Then
                If Me.chkUseSecondaryUnitb4.Checked = True Then
                    If JmlGudang >= JmlObjLamaAft Then
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - _
                        JmlDGV - JmlObjLamaAft), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text) + " @ " + Me.txtSecondaryQtyb4.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNob4.Text) + ".", MsgBoxStyle.Information, _
                        "Information")
                    Else
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlObjLama - _
                        JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text) + " @ " + Me.txtSecondaryQtyb4.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNob4.Text) + ".", MsgBoxStyle.Information, _
                        "Information")
                    End If
                Else
                    If JmlGudang >= JmlObjLamaAft Then
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV - JmlObjLamaAft), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text)) + ".", MsgBoxStyle.Information, _
                        "Information")
                    Else
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text)) + ".", MsgBoxStyle.Information, _
                        "Information")
                    End If
                End If
                Me.TxtQtyperBatchb4.Focus()
                Return False
            End If
            '9.b. Cek stock pada tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(Me.TxtPrimaryCodeNob4.Text, Me.dtpDate.Value)
            JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text)
            If JmlSetelahTambah < 0 Then
                If JmlGudang + JmlObjLama - JmlDGV > 0 Then
                    If Me.chkUseSecondaryUnitb4.Checked = True Then
                        MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text) + " @ " + Me.txtSecondaryQtyb4.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNob4.Text) + ".", MsgBoxStyle.Information, _
                        "Information")
                    Else
                        MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text)) + ".", MsgBoxStyle.Information, _
                        "Information")
                    End If
                Else
                    MsgBox("Stock level on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + _
                    " is 0. Entering this data will cause negative value in your stock.", _
                    MsgBoxStyle.Information, "Information")
                End If

                Me.TxtQtyperBatchb4.Focus()
                Return False
            End If

            '10. Bila ada item bundling after yang sama dengan item bundling before, akan mengakibatkan
            'perhitungan in amount pada bundling ngaco
            'Dim rowDGV As DataGridViewRow
            'For Each rowDGV In Me.dgvItemAfter.Rows
            '    If rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value = Me.TxtPrimaryCodeNob4.Text Then
            '        MsgBox("There product with code no: " + Me.TxtCodeNob4.Text + _
            '        " in item after, we cannot enter this item", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'Next

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function

    Private Function ApakahDataItemAfterBenar() As Boolean
        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDescAfter.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNoAfter.Text = ""
                Me.TxtCodeNoAfter.Focus()
                Return False
            End If
            '1.b. hanya boleh memasukkan goods
            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNoAfter.Text) = False Then
                MsgBox("Please enter only goods code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNoAfter.Text = ""
                Me.TxtCodeNoAfter.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchAfter.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchAfter.Focus()
                Return False
            End If
            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlledAfter.Checked = True Then
                If Me.txtBatchNoAfter.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtBatchNoAfter.Focus()
                    Return False
                End If
            End If
            '4.a Untuk transaksi proses (in/back/out proses) no Pdo harus diisikan
            '4.b Untuk transaksi yang berkaitan dengan purchase (stock out purchase return dan stout in purchase)
            '    harus menyertakan no PO bila terhubungkan dengan software purchase
            '4.c Untuk transaksi yang berkaitan dengan sales (stock out sales dan stock in sales return) harrus
            '    menyertakan no DO bila terhubungkan dengan software sales
            '5. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlledAfter.Checked = True Then
                If Me.txtPrimaryLocationAfter.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtLocationAfter.Focus()
                    Return False
                End If
            End If
            '7. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnitAfter.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.txtSecondaryQtyAfter.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.txtSecondaryQtyAfter.Focus()
                    Return False
                End If
            End If
            '8a. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlledAfter.Checked = True And Me.chkExpiryDateControlledAfter.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNoAfter.Text, Me.txtBatchNoAfter.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpiredAfter.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.txtBatchNoAfter.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpiredAfter.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItemAfter.Rows
                    If row.Cells(7).Value = Me.txtBatchNoAfter.Text And Me.txtBatchNoAfter.Text <> "" And row.Cells(12).Value = _
                    Me.TxtPrimaryCodeNoAfter.Text Then
                        If CDate(row.Cells(8).Value) <> Me.DtpExpiredAfter.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.txtBatchNoAfter.Text + " on " + _
                            Format(CDate(row.Cells(8).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.DtpExpiredAfter.Focus()
                            Return False
                        End If
                    End If
                Next
                ''8.c Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam transaksi lainnya: stock in adjust, stock in purchase, 
                '    stock in sales return, stock tag, transaksi bundling (item before) 
                '    (untuk stock out tidak dicek karena semua harus masuk dulu sebelum keluar, jadi dimasukka ke dalam pengecekan
                '    adalah semua transaksi yang menjadi pintu masuk inventory)
                '1. stock in (semua stock in)
                Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                For loopInOut As Integer = 0 To 4 Step 1
                    Select Case loopInOut
                        Case 0
                            '   a. stock in purchase
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                            Exit Select
                        Case 1
                            '   b. stock in sales return
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            Exit Select
                        Case 2
                            '   c. stock in adjust
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                            Exit Select
                        Case 3
                            '   d. stock in back
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                            Exit Select
                        Case 4
                            '   e. stock in proses
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                            Exit Select
                    End Select
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.TxtPrimaryCodeNoAfter.Text, Me.txtBatchNoAfter.Text, _
                    JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> Me.DtpExpiredAfter.Value Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNoAfter.Text) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.DtpExpiredAfter.Focus()
                            Return False
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.TxtPrimaryCodeNoAfter.Text, Me.txtBatchNoAfter.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpiredAfter.Value Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNoAfter.Text) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpiredAfter.Focus()
                        Return False
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")
            End If

            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlledAfter.Checked = True Then
                If Me.DtpExpiredAfter.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpiredAfter.Focus()
                        Return False
                    End If
                End If
            End If

            '7. Jika item before, sudah ada item 
            'Dim rowDGV As DataGridViewRow
            'For Each rowDGV In Me.dgvItemb4.Rows
            '    If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNoAfter.Text Then
            '        MsgBox("There product with code no: " + Me.TxtCodeNoAfter.Text + _
            '        " in item before, we cannot enter this item", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'Next

            'cost factor harus positif
            'If CDec(Me.txtFaktorPemberat.Text) <= 0 Then
            '    MsgBox("Cost factor must be a positif value", MsgBoxStyle.Information, "Information")
            '    Me.txtFaktorPemberat.Focus()
            '    Return False
            'End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Decimal, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal BatchNoControlled As Boolean, ByVal LocationControlled As Boolean, ByVal UseSecondaryUnit _
    As Boolean, ByVal vDGV As DataGridView) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In vDGV.Rows
                If row.Cells(12).Value = vPrimaryCodeNo Then
                    OK = True
                    'cek secondary qty
                    If UseSecondaryUnit = True And OK = True Then
                        If row.Cells(5).Value <> Secondary_Qty Then
                            OK = False
                        End If
                    End If
                    'cek location
                    If LocationControlled = True And OK = True Then
                        If vPrimaryLocation <> row.Cells(13).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If BatchNoControlled = True Then
                        If vBatchNo <> row.Cells(7).Value Then
                            OK = False
                        End If
                    End If
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    If OK = True Then
                        Jml = Jml + CDbl(row.Cells(3).Value)
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
        Dim row As DataGridViewRow
        Dim Jml As Decimal = 0
        Dim ConvertionDate As Date = #1/1/1950#
        Dim ObjAcc As New Core_Module.ClsTranskasiAccClosing

        Try
            ConvertionDate = ObjAcc.GetSMESSConvertionDate
            If ConvertionDate.Date > Me.dtpDate.Value.Date Then
                MsgBox("Minimum transaction date is on SMESS convertion date: " + Format(ConvertionDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Me.dtpDate.Focus()
                Return False
            End If
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("stock bundling number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            KoleksiNamaData.Add("Process code")
            KoleksiTextBox.Add(Me.TxtCodeProses)
            KoleksiNamaData.Add("Work Unit code")
            KoleksiTextBox.Add(Me.TxtCodeWorkUnit)
            'KoleksiNamaData.Add("Workshift code")
            'KoleksiTextBox.Add(Me.TxtWorkshiftCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item after dan 1 item before
            If Me.dgvItemAfter.RowCount = 0 Then
                MsgBox("Please enter at least 1 item as bundling result.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 1
                Return False
            End If
            If Me.dgvItemb4.RowCount = 0 Then
                MsgBox("Please enter at least 1 item as bundling source.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 0
                Return False
            End If

            'cek berkaitan dengan tanggal transaksi yang sudah dikunci
            'cek timestamp karena bisa saja sudah ada orang yang menyimpan perubahan data ketika user sedang
            'melakukan proses edit
            'msgbox("belum")

            ''ditutup dulu sementara by @ahdar
            'nilai cost factor jumlah totalnya harus sama dengan 100%, terlebih dahulu hitung jika semua baris cost factornya 0
            If Me.ObjStockBundling.TipeBundling <> ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                If Me.dgvItemAfter.ReadOnly = False Then
                    If MsgBox("Calculate cost factor automatically?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
                        ClsTransaksiSetting.EnumNamaParameter.Use_Custom_CostFaktorCalc) = "1" Then
                            'hitung menggunakan data custom
                            If Me.ObjCustom.HitungCostFactor(Me.dgvItemAfter, Me.ColQtyAfter.Index, Me.ColPrimaryCodeNoAfter.Index, _
                            Me.ColAfter_CostFactor.Index, Me.ColCodeNoAfter.Index, Me.ColPrimaryUnitAfter.Index) = False Then
                                MsgBox("Logic error in cost factor calculation, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        Else
                            'hitung berdasarkan qty
                            Me.HitungCostFactor()
                        End If
                    End If
                End If

                For Each row In Me.dgvItemAfter.Rows
                    Jml += CDec(row.Cells(Me.ColAfter_CostFactor.Index).Value)
                Next
                If Jml <> 100 Then
                    MsgBox("Sum of cost factor data must be 100%", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Return False
                End If
            End If
            '}

            '2016/10/10: setiap baris pada row item cost factornya harus positif
            For Each row In Me.dgvItemAfter.Rows
                If row.Cells(Me.ColAfter_CostFactor.Index).Value <= 0 Then
                    MsgBox("Invalid cost factor on row " + (row.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'berkaitan denga tanggal closing max
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.ObjStockBundling.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjStockBundling.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtStockInNo.Text = Me.ObjStockBundling.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjStockBundling.TanggalTransaksi
                        Me.TxtStockInNo.Text = Me.ObjStockBundling.NoTransaksi
                        Return False
                    End If
                End If
            End If

            'custom
            If Me.ObjCustom.ApakahDataBenar_Custom(Me.dgvItemAfter, Me.ColQtyAfter.Index, Me.ColPrimaryCodeNoAfter.Index, _
            Me.ColAfter_CostFactor.Index, Me.ColCodeNoAfter.Index, Me.ColPrimaryUnitAfter.Index, Me.TxtPrimaryProses.Text, _
            Me.ColPrimaryLocationAfter.Index, Me.ColPrimaryLocation.Index, Me.dgvItemb4) = False Then
                Return False
            End If

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation_Before()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_BUNDLING_BEFORE, Me.TBLDetailLocation) = False Then
                    Return False
                End If
            End If
            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation_After()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_BUNDLING_AFTER, Me.TBLDetailLocation) = False Then
                    Return False
                End If
            End If

            'cek berkaitan perhitungan tanpa batas pada item bundling after dan before, dipindahkan ke kelas, karena
            'ada fasilitas antrian untuk transaksi di kelas.

            'kalo sampe sini return true

            Dim ObjCust As New CustomSMESS.ClsTransaksiStockBundling
            Dim rd As DataGridViewRow
            Dim TblDummyBef As New CustomSMESS.DseTransaksiStockBundling.tbltrans_custom_stock_bundling_item_beforeDataTable
            TblDummyBef = Me.TblItemBefCustom.Copy
            'If TblDummyBef.Rows.Count > 0 Then
            For i As Integer = 0 To Me.dgvItemb4.RowCount - 1
                rd = Me.dgvItemb4.Rows(i)
                With rd
                    Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    Dim PrimaryPOSO As Double = -1
                    Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                    Dim Batch As String = .Cells(Me.ColBatch.Index).Value
                    Dim QtyDef As Decimal = .Cells(Me.ColQty.Index).Value
                    Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                    If ObjCust.CompareQtySMESS_CustomBefore(PrimaryProduct, PrimaryLocation, Batch, QtyDef, TblDummyBef) = False Then
                        If String.IsNullOrEmpty(Batch) = False Then
                            ProdCode &= " batch: " & Batch
                        End If
                        MsgBox("Inconsistent data, please cek product: " & ProdCode & "between main form and custom form.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemb4.ClearSelection()
                        rd.Selected = True
                        Return False
                    End If
                End With
            Next
            'End If

            Dim TblDummyAft As New CustomSMESS.DseTransaksiStockBundling.tbltrans_custom_stock_bundling_item_afterDataTable
            TblDummyAft = Me.TblItemAftCustom.Copy
            'If TblDummyAft.Rows.Count > 0 Then
            For i As Integer = 0 To Me.dgvItemAfter.RowCount - 1
                rd = Me.dgvItemAfter.Rows(i)
                With rd
                    Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    Dim PrimaryPOSO As Double = -1
                    Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                    Dim Batch As String = .Cells(Me.ColBatch.Index).Value
                    Dim QtyDef As Decimal = .Cells(Me.ColQty.Index).Value
                    Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                    If ObjCust.CompareQtySMESS_CustomAfter(PrimaryProduct, PrimaryLocation, Batch, QtyDef, TblDummyAft) = False Then
                        If String.IsNullOrEmpty(Batch) = False Then
                            ProdCode &= " batch: " & Batch
                        End If
                        MsgBox("Inconsisten data, please cek product: " & ProdCode & "between main form and custom form.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemAfter.ClearSelection()
                        rd.Selected = True
                        Return False
                    End If
                End With
            Next
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            ObjAcc = Nothing
        End Try
    End Function

    Private Sub LoadLocation_After()
        Dim RowDetailLocation As DseUserLocationPrivilegeSetting.tblmaster_storage_areaRow
        Dim rowDGV As DataGridViewRow
        Dim ListPrimary As New List(Of Double)

        Try
            For Each rowDGV In Me.dgvItemAfter.Rows
                If rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value <> -1 Then
                    If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value) = False Then
                        RowDetailLocation = Me.TBLDetailLocation.Newtblmaster_storage_areaRow
                        With RowDetailLocation
                            .Primary = rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value
                            .Storage_Area_Code = Me.ObjLocation.GetStorageAreaCode(.Primary)
                            .Primary_Warehouse = Me.ObjLocation.GetPrimaryWarehouse(.Primary)
                        End With
                        Me.TBLDetailLocation.Rows.Add(RowDetailLocation)
                        ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadLocation_Before()
        Dim RowDetailLocation As DseUserLocationPrivilegeSetting.tblmaster_storage_areaRow
        Dim rowDGV As DataGridViewRow
        Dim ListPrimary As New List(Of Double)

        Try
            For Each rowDGV In Me.dgvItemb4.Rows
                If rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value <> -1 Then
                    If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value) = False Then
                        RowDetailLocation = Me.TBLDetailLocation.Newtblmaster_storage_areaRow
                        With RowDetailLocation
                            .Primary = rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value
                            .Storage_Area_Code = rowDGV.Cells(Me.ColLocation.Index).Value
                            .Primary_Warehouse = Me.ObjLocation.GetPrimaryWarehouse(.Primary)
                        End With
                        Me.TBLDetailLocation.Rows.Add(RowDetailLocation)
                        ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryLocationAfter.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungCostFactor()
        Dim rowDGV As DataGridViewRow
        Dim TotalJml As Decimal = 0
        Dim Dummy As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAfter.Rows
                TotalJml += CDec(rowDGV.Cells(Me.ColQtyAfter.Index).Value)
            Next
            For Each rowDGV In Me.dgvItemAfter.Rows
                If rowDGV.Index = Me.dgvItemAfter.Rows.Count - 1 Then
                    'untuk baris terakhir = 100 - sub total
                    rowDGV.Cells(Me.ColAfter_CostFactor.Index).Value = Me.ObjCD.DigitGrouping(100 - Dummy, 2)
                Else
                    rowDGV.Cells(Me.ColAfter_CostFactor.Index).Value = ObjCD.DigitGrouping((CDec(rowDGV.Cells(Me.ColQtyAfter.Index).Value) * 100) / TotalJml, 2)
                    Dummy += CDec(rowDGV.Cells(Me.ColAfter_CostFactor.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HilangkanItem(ByVal TBLDummy As dseTransaksiStockBundling.DataTable3DataTable)
        Dim Filter As String = ""
        Filter = "[Primary Transaksi] = " + Me.TxtPrimaryStockIn.Text

        Try
            TBLDummy.DefaultView.RowFilter = Filter
            While TBLDummy.DefaultView.Count > 0
                TBLDummy.Rows.Remove(TBLDummy.DefaultView(0).Row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanItem(ByVal TBLDummy As dseTransaksiStockBundling.DataTable3DataTable)
        Dim rowDGV_B4 As DataGridViewRow
        Dim rowDGV_After As DataGridViewRow
        Dim rowBundling As dseTransaksiStockBundling.DataTable3Row

        Try
            For Each rowDGV_B4 In Me.dgvItemb4.Rows
                For Each rowDGV_After In Me.dgvItemAfter.Rows
                    rowBundling = TBLDummy.NewDataTable3Row
                    rowBundling.Primary_Product_After = rowDGV_After.Cells(Me.ColPrimaryCodeNoAfter.Index).Value
                    rowBundling.Primary_Product_Before = rowDGV_B4.Cells(Me.ColPrimaryCodeNo.Index).Value
                    rowBundling.Primary_Transaksi = Me.TxtPrimaryStockIn.Text
                    rowBundling.Tanggal_Bundling = Me.dtpDate.Value.Date
                    TBLDummy.Rows.Add(rowBundling)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjStockBundling
                .NoTransaksi = Me.TxtStockInNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimaryProses = Me.TxtPrimaryProses.Text
                .PrimaryWorkUnit = Me.TxtPrimaryWorkUnit.Text
                .PrimaryWorkshift = Me.TxtPrimaryWorkShift.Text
                .Remarks = Me.TxtRemark.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjStockBundling
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryProses.Text = .PrimaryProses
                Me.TxtCodeProses.Text = Me.ObjProses.GetCode(Me.TxtPrimaryProses.Text)
                Me.TxtDescProses.Text = Me.ObjProses.GetDesc(Me.TxtPrimaryProses.Text)
                Me.TxtPrimaryWorkUnit.Text = .PrimaryWorkUnit
                Me.TxtCodeWorkUnit.Text = Me.ObjWorkUnit.GetCode(Me.TxtPrimaryWorkUnit.Text)
                Me.TxtDescWorkUnit.Text = Me.ObjWorkUnit.GetDesc(Me.TxtPrimaryWorkUnit.Text)
                Me.TxtPrimaryWorkShift.Text = .PrimaryWorkshift
                Me.TxtWorkshiftCode.Text = Me.ObjWorkShift.GetCodeByPrimary(Me.TxtPrimaryWorkShift.Text)
                Me.TxtWorkshiftDesc.Text = Me.ObjWorkShift.GetDescription(Me.TxtPrimaryWorkShift.Text)
                Me.TxtPrimaryStockIn.Text = .Primary

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV(Me.dgvItemb4)
            Me.FormatDGV(Me.dgvItemAfter)
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItemb4.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
                If IsNumeric(row.Cells(5).Value) = True Then
                    row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
                End If
            Next
            For Each row In Me.dgvItemAfter.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
                If IsNumeric(row.Cells(5).Value) = True Then
                    row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()

            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Purchase1 Then
                'Me.gboItemAfter.Height = 156
                'Me.TxtRemarkItemAfter.Width = 285
                'Me.BtnInputAfter.Top = 121
                'Me.lblFaktorPemberat.Visible = True
                'Me.lblFaktorPemberat.Text = "Cost Factor (%)"
                'Me.txtFaktorPemberat.Visible = True
                'Me.dgvItemAfter.Top = 166
                'Me.dgvItemAfter.Height = Me.tabAfter.Height - Me.dgvItemAfter.Top - Me.btnUpAfter.Height - 5
            Else
                Me.dgvItemAfter.Columns(16).Visible = False
                Me.txtFaktorPemberat.TabStop = False
            End If
            Me.dgvItemAfter.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemb4.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.SetBtnUpDown()
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItemAfter()
                Me.KosongkanItemB4()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjStockBundling.SimpanBaru(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) Then
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
                        If Me.ObjStockBundling.SimpanPerubahan(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) Then
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
            If Me.TxtRemark.Text = "" And Me.dgvItemb4.RowCount = 0 And Me.dgvItemAfter.RowCount = 0 And _
            Me.dtpDate.Value.Date = Today And Me.TxtCodeNob4.Text = "" And Me.TxtCodeNoAfter.Text = "" And _
            Me.cboBatchNob4.Text = "" And Me.txtBatchNoAfter.Text = "" And Me.cboLocationb4.Text = "" And _
            Me.TxtLocationAfter.Text = "" And Me.TxtRemarkItemb4.Text = "" And Me.TxtRemarkItemAfter.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItemAfter()
                    Me.KosongkanItemB4()

                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjStockBundling.CariDetail(Me.TxtPrimaryStockIn.Text, Me.dgvItemb4, Me.dgvItemAfter, Me)
                        Me.KosongkanItemAfter()
                        Me.KosongkanItemB4()
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
            Me.ObjStockBundling.ResetObject(Me.dgvItemb4, Me.dgvItemAfter)
            Me.TxtStockInNo.Text = Me.ObjStockBundling.GetNoBaru(Me.dtpDate.Value, ClsTransaksiStockBundling.EnumTipeBundling.Purchase1)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiStockBundling

        Try
            ij.TipeBundling = Me.ObjStockBundling.TipeBundling
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryTransaksiStockBundling = -1 Then
                If Me.ObjStockBundling.CariDetail(ij.PrimaryTransaksiStockBundling, _
                Me.dgvItemb4, Me.dgvItemAfter, Me) = True Then
                    Me.SalinKeField()
                    Me.TabControl1.SelectedIndex = 0
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                If Me.ObjStockBundling.ApakahBolehEdit(Me.UserCode) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItemb4.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim ListProd As New List(Of Double)
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable

        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                If Me.ObjStockBundling.ApakahBolehHapus(Me.UserCode, False, tblProdQty, ListProd) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjStockBundling.Hapus(Me.TxtPrimaryStockIn.Text, Me.UserCode, True, True) = True Then
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
                            If Me.ObjStockBundling.SimpanBaru(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) = True Then
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
                            If Me.ObjStockBundling.SimpanPerubahan(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) = True Then
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
            Me.TBL.DefaultView.Sort = "[Expiry Date] ASC, [Lot No] ASC"
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
                'Me.lblUnitsb4.Left = Me.lblSecondaryUnitsb4.Left
                'Me.TxtQtyperBatchb4.Width = Me.lblUnitsb4.Left - Me.TxtQtyperBatchb4.Left - 6
                Me.txtSecondaryQtyb4.Text = "-1"
                Me.lblSecQtyb4.Visible = False
                Me.cboSecondaryUnitsb4.Visible = False
                Me.cboSecondaryUnitsb4.DataSource = Nothing
                Me.cboSecondaryUnitsb4.Items.Clear()
            Else
                Me.lblSecondaryUnitsb4.Visible = True
                Me.lblSecQtyb4.Visible = True
                'Me.lblUnitsb4.Left = Me.lblSecQtyb4.Left - 37
                'Me.TxtQtyperBatchb4.Width = Me.lblUnitsb4.Left - Me.TxtQtyperBatchb4.Left - 6
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

    Private Sub TxtCodeNob4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNob4.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0
        Dim PrimaryProduct As Double = -1

        Try
            i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only

            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNob4.Text)
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)

                If Jml = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                    Me.TxtPrimaryCodeNob4.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)
                    Me.txtPrimaryCategoryb4.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNob4.Text)
                    Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                    Me.SetGboItemb4()
                    If Me.cboBatchNob4.Visible = True Then
                        Me.cboBatchNob4.Focus()
                    ElseIf Me.cboLocationb4.Visible = True Then
                        Me.cboLocationb4.Focus()
                    Else
                        Me.TxtQtyperBatchb4.Focus()
                    End If

                    'ahdar
                    Me.GetQtyAvailable()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNob4.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNob4.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNob4.Text = i.PrimaryProduct
                        Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
                        Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                        Me.SetGboItemb4()
                        If Me.cboBatchNob4.Visible = True Then
                            Me.cboBatchNob4.Focus()
                        ElseIf Me.cboLocationb4.Visible = True Then
                            Me.cboLocationb4.Focus()
                        Else
                            Me.TxtQtyperBatchb4.Focus()
                        End If

                        'ahdar
                        Me.GetQtyAvailable()
                    Else
                        Me.TxtCodeNob4.Clear()
                        Me.TxtPrimaryCodeNob4.Text = "-1"
                        Me.txtPrimaryCategoryb4.Text = "-1"
                        Me.KosongkanItemB4()
                        Me.TxtCodeNob4.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNob4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNob4.Leave
        Dim i As New frmCariMasterProduct
        Dim PrimaryProduct As Double = -1

        Try
            i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
            If Me.TxtCodeNob4.Text <> "" Then
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)

                If Me.ObjProduct.GetJml(Me.TxtCodeNob4.Text) = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                    Me.TxtPrimaryCodeNob4.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)
                    Me.txtPrimaryCategoryb4.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNob4.Text)
                    Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                    Me.SetGboItemb4()
                    If Me.cboBatchNob4.Visible = True Then
                        Me.cboBatchNob4.Focus()
                    ElseIf Me.cboLocationb4.Visible = True Then
                        Me.cboLocationb4.Focus()
                    Else
                        Me.TxtQtyperBatchb4.Focus()
                    End If

                    'ahdar
                    Me.GetQtyAvailable()
                Else
                    If Me.TxtCodeNob4.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNob4.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNob4.Text = i.PrimaryProduct
                        Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
                        Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                        Me.SetGboItemb4()
                        If Me.cboBatchNob4.Visible = True Then
                            Me.cboBatchNob4.Focus()
                        ElseIf Me.cboLocationb4.Visible = True Then
                            Me.cboLocationb4.Focus()
                        Else
                            Me.TxtQtyperBatchb4.Focus()
                        End If

                        'ahdar
                        Me.GetQtyAvailable()
                    Else
                        Me.TxtCodeNob4.Clear()
                        Me.TxtPrimaryCodeNob4.Text = "-1"
                        Me.txtPrimaryCategoryb4.Text = "-1"
                        Me.KosongkanItemB4()
                        Me.TxtCodeNob4.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNob4.Text = "-1"
                Me.txtDescb4.Clear()
                Me.txtPrimaryCategoryb4.Text = "-1"
                Me.KosongkanItemB4()
            End If
            Me.SetGboItemb4()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProductb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNob4.Click
        Dim i As New frmCariMasterProduct

        Try
            i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Goods_Only
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNob4.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNob4.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNob4.Text = i.PrimaryProduct
                Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
                Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                Me.SetGboItemb4()
                If Me.cboBatchNob4.Visible = True Then
                    Me.cboBatchNob4.Focus()
                ElseIf Me.cboLocationb4.Visible = True Then
                    Me.cboLocationb4.Focus()
                Else
                    Me.TxtQtyperBatchb4.Focus()
                End If

                'ahdar
                Me.GetQtyAvailable()
            Else
                Me.TxtCodeNob4.Clear()
                Me.TxtPrimaryCodeNob4.Text = "-1"
                Me.txtPrimaryCategoryb4.Text = "-1"
                Me.KosongkanItemB4()
                Me.TxtCodeNob4.Focus()
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
    ByVal SecondaryQty As Decimal, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date, ByVal vDGV As DataGridView) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In vDGV.Rows
                If IsDate(row.Cells(8).Value) = False Then
                    If IsNumeric(row.Cells(5).Value) = False Then
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And row.Cells(5).Value = SecondaryQty And _
                        ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(5).Value) = False Then
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(8).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And row.Cells(5).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(8).Value) Then
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

    Private Sub BtnInputb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInputb4.Click
        Try
            Me.InputB4(True, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputB4(ByVal vFollowUpFormatDGV As Boolean, ByVal vCekApakahDataBenar As Boolean)
        Dim Lanjut As Boolean = False

        Try
            If vCekApakahDataBenar = True Then
                Lanjut = Me.ApakahDataItemB4Benar
            Else
                Lanjut = True
            End If

            If Lanjut = True Then
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNob4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, _
                 Me.txtSecondaryQtyb4.Text, Me.txtPrimaryPDOb4.Text, Me.DtpExpiredb4.Value, Me.dgvItemb4)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItemb4.Rows(Index)
                    row.Cells(3).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(3).Value) + CDbl(Me.TxtQtyperBatchb4.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text))
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpiredb4.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpiredb4.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItemb4.Rows.Add("", Me.TxtCodeNob4.Text, Me.txtDescb4.Text, Me.TxtQtyperBatchb4.Text, _
                    Me.lblUnitsb4.Text, Me.txtSecondaryQtyb4.Text, Me.lblSecondaryUnitsb4.Text, Me.cboBatchNob4.Text, _
                    ExpiredDateStr, Me.cboLocationb4.Text, Me.TxtWarehouseb4.Text, Me.TxtRemarkItemb4.Text, _
                    Me.TxtPrimaryCodeNob4.Text, Me.txtPrimaryLocationb4.Text, Me.txtPrimaryUnitsb4.Text, _
                    Me.txtPrimarySecondaryUnitsb4.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItemb4)
                End If

                If vFollowUpFormatDGV = True Then
                    'kosongkan gbo item
                    Me.KosongkanItemB4()
                    Me.TxtCodeNob4.Focus()
                    'format dgv
                    Me.FormatDGV(Me.dgvItemb4)
                    'fokuskan ke baris terakhir
                    While Me.dgvItemb4.Rows(Me.dgvItemb4.Rows.Count - 1).Displayed = False _
                    Or Me.dgvItemb4.DisplayedRowCount(True) <> Me.dgvItemb4.DisplayedRowCount(False)
                        Me.dgvItemb4.FirstDisplayedScrollingRowIndex = Me.dgvItemb4.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV(ByVal vDGV As DataGridView)
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            'yang diformat dgv-nya:
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each row In vDGV.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(8).Value) = False Then
                    row.Cells(8).Value = "-"
                Else
                    If CDate(row.Cells(8).Value).Date = #1/1/1950# Then
                        row.Cells(8).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(8).Value = Format(CDate(row.Cells(8).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(5).Value) = True Then
                    If row.Cells(5).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(5).Value = "-1" Then
                    row.Cells(5).Value = "-"
                    row.Cells(6).Value = "-"
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(13).Value = "-1" Then
                    row.Cells(9).Value = "-"
                    row.Cells(10).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(7).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                vDGV.Columns(5).Visible = False
                vDGV.Columns(6).Visible = False
            Else
                vDGV.Columns(5).Visible = True
                vDGV.Columns(6).Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                vDGV.Columns(9).Visible = False
                vDGV.Columns(10).Visible = False
            Else
                vDGV.Columns(9).Visible = True
                vDGV.Columns(10).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                vDGV.Columns(7).Visible = False
            Else
                vDGV.Columns(7).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                vDGV.Columns(8).Visible = False
            Else
                vDGV.Columns(8).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchb4, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchb4.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchb4, sender, e)
    End Sub

    Private Sub TxtQtyperBatchb4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchb4.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchb4, sender, e)
    End Sub

    Private Sub txtSecondaryQtyb4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyb4.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQtyb4, sender, e)
    End Sub

    Private Sub txtSecondaryQtyb4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyb4.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQtyb4, sender, e)
    End Sub

    Private Sub dgvItemb4_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemb4.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Location Code
            '5. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = 3 Or e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 _
                Or e.ColumnIndex = 11 Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = 5 Then
                        'cek secondary qty
                        If CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 7 Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItemb4.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 8 Then
                        'cek expiry date
                        If Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 9 Then
                        'cek location controlled
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItemb4.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value

                    If e.ColumnIndex = Me.ColQty.Index Then
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
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

    Private Sub dgvItemb4_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemb4.CellEndEdit
        Try
            Dim SecondaryQty, JmlDGV As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItemb4.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> 8 Then
                        If Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> 11 Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(8).Value) = True Then
                            ExpiryDate = CDate(.Cells(8).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(5).Value) = True Then
                            SecondaryQty = .Cells(5).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = 3 Then
                            Me.EndEditQtyb4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                            Me.GetQtyAvailable()
                        ElseIf e.ColumnIndex = 5 Then
                            Me.EndEditSecondaryUnitb4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 7 Then
                            Me.EndEditBatchNob4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                            Me.GetQtyAvailable()
                        ElseIf e.ColumnIndex = 9 Then
                            Me.EndEditLocationb4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = 11 Then
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQtyb4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'digit grouping
                Me.dgvItemb4(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItemb4(e.ColumnIndex, e.RowIndex).Value, _
                Me.ObjProduct.GetAngkaDecimal(Me.dgvItemb4.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                JmlDGV = .Cells(3).Value
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnitb4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, .Cells(5).Value, JmlDGV) = False Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemb4)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemb4)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNob4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan diisikan
                If CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'batch no
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItemb4.Item(8, e.RowIndex).Value) = True Then
                        'basis data
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItemb4.Item(12, e.RowIndex).Value, _
                        Me.dgvItemb4.Item(7, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItemb4(8, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemb4.Item(7, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItemb4.Item(8, e.RowIndex).Value)
                                    Me.dgvItemb4.Item(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        'data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItemb4.Rows
                            If row1.Cells(7).Value = Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells(12).Value = Me.dgvItemb4.Item(12, e.RowIndex).Value Then
                                If CDate(row1.Cells(8).Value) <> CDate(Me.dgvItemb4.Item(8, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemb4.Item(7, e.RowIndex).Value) + " is " + _
                                        Format(CDate(row1.Cells(8).Value), "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", _
                                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItemb4.Item(8, e.RowIndex).Value)
                                        Me.dgvItemb4.Item(8, e.RowIndex).Value = row1.Cells(8).Value
                                        Exit For
                                    Else
                                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '3. cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemb4)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemb4.Item(8, e.RowIndex).Value = BufferExpiryDate
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemb4)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocationb4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim Obj As New clsMasterStorageArea
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                If Obj.ApakahSudahAda(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterStorageArea
                    If CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) <> "" Then
                        i.txtCari.Text = Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea <> -1 Then
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetStorageAreaCode(i.PrimaryStorageArea)).ToUpper
                    Else
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        i = Nothing
                        Exit Sub
                    End If
                    i = Nothing
                Else
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                End If
                'update data primary location
                .Cells(13).Value = Obj.GetPrimary(.Cells(9).Value)
                'update data gudang
                .Cells(10).Value = Obj.GetWarehouseCode(.Cells(13).Value)

                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    'kembalikan ke data semula
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    .Cells(13).Value = Obj.GetPrimary(.Cells(9).Value)
                    'update data gudang
                    .Cells(10).Value = Obj.GetWarehouseCode(.Cells(13).Value)
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemb4)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan ke data semula
                            Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'update data primary location
                            .Cells(13).Value = Obj.GetPrimary(.Cells(9).Value)
                            'update data gudang
                            .Cells(10).Value = Obj.GetWarehouseCode(.Cells(13).Value)
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemb4)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Function ApakahAdaBarisYangSama(ByVal vDgv As DataGridView) As Boolean
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= vDgv.RowCount - 2
                row = vDgv.Rows(i)
                j = i + 1
                While j <= vDgv.RowCount - 1
                    row1 = vDgv.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse

                    If row.Cells(12).Value = row1.Cells(12).Value And row.Cells(2).Value = row1.Cells(2).Value And _
                    row.Cells(3).Value = row1.Cells(3).Value And row.Cells(4).Value = row1.Cells(4).Value And _
                    CStr(row.Cells(7).Value) = CStr(row1.Cells(7).Value) And row.Cells(9).Value = _
                    row1.Cells(9).Value Then
                        If IsNumeric(row.Cells(5).Value) = True And IsNumeric(row1.Cells(5).Value) = True Then
                            If row.Cells(5).Value = row1.Cells(5).Value Then
                                Return True
                            Else
                                Return False
                            End If
                        ElseIf IsNumeric(row.Cells(5).Value) = False And IsNumeric(row1.Cells(5).Value) = False Then
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
    Private Sub GabungkanBaris(ByVal vdGV As DataGridView)
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= vdGV.RowCount - 2
                row = vdGV.Rows(i)
                j = i + 1
                While j <= vdGV.RowCount - 1
                    row1 = vdGV.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse
                    If CStr(row.Cells(1).Value) = CStr(row1.Cells(1).Value) And CStr(row.Cells(2).Value) = CStr(row1.Cells(2).Value) And _
                    CStr(row.Cells(3).Value) = CStr(row1.Cells(3).Value) And CStr(row.Cells(12).Value) = CStr(row1.Cells(12).Value) And _
                    CStr(row.Cells(7).Value) = CStr(row1.Cells(7).Value) And _
                    CStr(row.Cells(9).Value) = CStr(row1.Cells(9).Value) Then
                        If IsNumeric(row.Cells(5).Value) = True And IsNumeric(row1.Cells(5).Value) = True Then
                            If CDbl(row.Cells(5).Value) = CDbl(row1.Cells(5).Value) Then
                                Me.GabungkanBaris(row, row1, vdGV)
                                Me.ObjCD.SetNumber(vdGV)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(5).Value) = False And IsNumeric(row1.Cells(5).Value) = False Then
                            Me.GabungkanBaris(row, row1, vdGV)
                            Me.ObjCD.SetNumber(vdGV)
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

    Private Sub GabungkanBaris(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow, _
    ByVal vDGV As DataGridView)
        Try
            'gabungkan data qty.
            row.Cells(3).Value = CDbl(row.Cells(3).Value) + CDbl(row1.Cells(3).Value)
            'gabungkan data remarks
            If CStr(row.Cells(1).Value).EndsWith(".") = False Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(11).Value = CStr(row.Cells(11).Value) + ". "
            End If
            row.Cells(11).Value = CStr(row.Cells(11).Value) + CStr(row1.Cells(11).Value)
            'hapus baris ybs
            vDGV.Rows.RemoveAt(row1.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSahb4(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
    ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain, JmlObjAfter As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockBundlingItem.GetJmlB4(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
            JmlObjAfter = Me.objStockBundlingItem.GetJmlAfter(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary, Me.dgvItemb4)

            If JmlGudang >= JmlObjAfter Then
                JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal - JmlObjAfter
            Else
                JmlSetelahEdit = JmlObjLama - JmlDGVTotal
            End If
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    If JmlGudang >= JmlObjAfter Then
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGVLain - JmlObjAfter), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
                        Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlObjLama - JmlDGVLain), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
                        Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                    End If
                End If
                Return False
            End If

            'cek stock untuk tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(PrimaryProduct, Me.dtpDate.Value)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal - JmlObjAfter
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + "  minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGVLain - JmlObjAfter), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
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

    Private Sub dgvItemb4_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemb4.RowsAdded
        Me.dgvItemb4.TabStop = True
        If Me.TxtRemark.Enabled = True Then
            Me.btnUPb4.Enabled = True
            Me.btnDownb4.Enabled = True
        End If
    End Sub

    Private Sub dgvItemb4_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemb4.RowsRemoved
        If Me.dgvItemb4.Rows.Count = 0 Then
            Me.dgvItemb4.TabStop = False
            If Me.TxtRemark.Enabled = True Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
            End If
        End If
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemb4.SelectionChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItemb4.SelectedRows.Count > 0 Then
                    Me.btnUPb4.Enabled = True
                    Me.btnDownb4.Enabled = True
                Else
                    Me.btnUPb4.Enabled = False
                    Me.btnDownb4.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemb4_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItemb4.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemb4)
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean, ByVal vdGV As DataGridView)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < vdGV.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > vdGV.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While vdGV.Rows(row.Index).Displayed = False
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If vdGV.DisplayedRowCount(True) <> vdGV.DisplayedRowCount(False) Then
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUPb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUPb4.Click
        Try
            If Me.dgvItemb4.SelectedRows.Count > 0 Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemb4.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemb4.Rows.Remove(row)
                    Me.dgvItemb4.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemb4.Focus()
                    Me.dgvItemb4.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemb4)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True, Me.dgvItemb4)
                    End If
                End If
                Me.btnUPb4.Enabled = True
                Me.btnDownb4.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDownb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownb4.Click
        Try
            If Me.dgvItemb4.SelectedRows.Count > 0 Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemb4.SelectedRows(0)
                If row.Index <> Me.dgvItemb4.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemb4.Rows.Remove(row)
                    Me.dgvItemb4.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemb4.Focus()
                    Me.dgvItemb4.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemb4)
                    Me.ShowRow(row, False, Me.dgvItemb4)
                End If
                Me.btnUPb4.Enabled = True
                Me.btnDownb4.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemb4_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemb4.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemb4)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemb4_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemb4.UserDeletingRow
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
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtRemark.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtRemark.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            Me.TxtRemark.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNob4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNob4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtCodeNob4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtCodeNob4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNob4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNob4.LostFocus
        Try
            Me.TxtCodeNob4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtQtyperBatchb4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtQtyperBatchb4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.LostFocus
        Try
            Me.TxtQtyperBatchb4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyb4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyb4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.txtSecondaryQtyb4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.txtSecondaryQtyb4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyb4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyb4.LostFocus
        Try
            Me.txtSecondaryQtyb4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemb4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemb4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtRemarkItemb4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtRemarkItemb4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemb4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemb4.LostFocus
        Try
            Me.TxtRemarkItemb4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNob4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNob4.SelectedIndexChanged
        Try
            If Me.SedangAddRM = True Then
                Exit Sub
            End If

            If Me.chkLocationControlledb4.Checked = True Then
                If Me.cboBatchNob4.Text <> "" Then
                    'isikan data lokasi
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockIn.Text, Me.cboBatchNob4.Text)
                End If
            Else
                'kosongkan data lokasi
                Me.cboLocationb4.DataSource = Nothing
                Me.cboLocationb4.Items.Clear()
            End If
            'isikan data expiry date 
            If Me.chkExpiryDateControlledb4.Checked = True Then
                Me.DtpExpiredb4.Value = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNob4.Text, Me.cboBatchNob4.Text)
            Else
                If Me.DtpExpiredb4.Visible = True Then
                    Me.DtpExpiredb4.Value = Today
                Else
                    Me.DtpExpiredb4.Value = #1/1/1950#
                End If
            End If
            'isikan secondary unit
            If Me.chkLocationControlledb4.Checked = False And _
            Me.chkUseSecondaryUnitb4.Checked = True And Me.cboBatchNob4.Text <> "" Then
                Me.IsiCboSecondary(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockIn.Text, _
                Me.cboBatchNob4.Text, Me.txtPrimaryLocationb4.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocationb4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocationb4.SelectedIndexChanged
        Dim objLoc As New clsMasterStorageArea

        Try
            If Me.SedangAddRM = True Then
                Exit Sub
            End If

            If Me.cboLocationb4.Text = "" Or objLoc.ApakahSudahAda(Me.cboLocationb4.Text) = False Then
                Me.txtPrimaryLocationb4.Text = "-1"
                Me.TxtWarehouseAfter.Clear()
            Else
                Me.txtPrimaryLocationb4.Text = objLoc.GetPrimary(Me.cboLocationb4.Text)
                Me.TxtWarehouseb4.Text = objLoc.GetWarehouseCode(Me.txtPrimaryLocationb4.Text)
                Me.IsiCboSecondary(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockIn.Text, Me.cboBatchNob4.Text, _
                Me.txtPrimaryLocationb4.Text)

                'ahdar
                Me.GetQtyAvailable()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objLoc = Nothing
        End Try
    End Sub

    Private Sub cboSecondaryUnitsb4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSecondaryUnitsb4.SelectedIndexChanged
        Try
            If Me.cboSecondaryUnitsb4.SelectedIndex <> -1 Then
                Me.txtSecondaryQtyb4.Text = Me.cboSecondaryUnitsb4.Text
            Else
                Me.txtSecondaryQtyb4.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub BtnInputAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInputAfter.Click
        Try
            If Me.ApakahDataItemAfterBenar = True Then
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNoAfter.Text, Me.txtPrimaryLocationAfter.Text, Me.txtBatchNoAfter.Text, _
                 Me.txtSecondaryQtyAfter.Text, Me.txtPrimaryPDOAfter.Text, Me.DtpExpiredAfter.Value, Me.dgvItemAfter)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItemAfter.Rows(Index)
                    row.Cells(3).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(3).Value) + CDbl(Me.TxtQtyperBatchAfter.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNoAfter.Text))
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpiredAfter.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpiredAfter.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItemAfter.Rows.Add("", Me.TxtCodeNoAfter.Text, Me.txtDescAfter.Text, Me.TxtQtyperBatchAfter.Text, _
                    Me.lblUnitsAfter.Text, Me.txtSecondaryQtyAfter.Text, Me.lblSecondaryUnitsAfter.Text, Me.txtBatchNoAfter.Text, _
                    ExpiredDateStr, Me.TxtLocationAfter.Text, Me.TxtWarehouseAfter.Text, Me.TxtRemarkItemAfter.Text, _
                    Me.TxtPrimaryCodeNoAfter.Text, Me.txtPrimaryLocationAfter.Text, Me.txtPrimaryUnitsAfter.Text, _
                    Me.txtPrimarySecondaryUnitsAfter.Text, Me.txtFaktorPemberat.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItemAfter)
                End If

                'kosongkan gbo item
                Me.KosongkanItemAfter()
                Me.TxtCodeNoAfter.Focus()
                'format dgv
                Me.FormatDGV(Me.dgvItemAfter)
                'fokuskan ke baris terakhir
                While Me.dgvItemAfter.Rows(Me.dgvItemAfter.Rows.Count - 1).Displayed = False _
                Or Me.dgvItemAfter.DisplayedRowCount(True) <> Me.dgvItemAfter.DisplayedRowCount(False)
                    Me.dgvItemAfter.FirstDisplayedScrollingRowIndex = Me.dgvItemAfter.FirstDisplayedScrollingRowIndex + 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchAfter, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchAfter.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchAfter, sender, e)
    End Sub

    Private Sub TxtQtyperBatchAfter_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchAfter.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchAfter, sender, e)
    End Sub

    Private Sub txtSecondaryQtyAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyAfter.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQtyAfter, sender, e)
    End Sub

    Private Sub txtSecondaryQtyAfter_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyAfter.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQtyAfter, sender, e)
    End Sub

    Private Sub dgvItemAfter_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemAfter.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Expiry Date
            '5. Location Code
            '6. Remarks
            '7. Cost Factor/Faktor Pemberat

            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = 3 Or e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 _
                Or e.ColumnIndex = 11 Or e.ColumnIndex = 8 Or e.ColumnIndex = 16 Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = 5 Then
                        'cek secondary qty
                        If CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 7 Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItemAfter.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 8 Then
                        'cek expiry date
                        If Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 9 Then
                        'cek location controlled
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItemAfter.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If                
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value

                    If e.ColumnIndex = Me.ColQtyAfter.Index Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
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

    Private Sub dgvItemAfter_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemAfter.CellEndEdit
        Try
            Dim SecondaryQty, JmlDGV As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItemAfter.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> 8 Then
                        If Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> 11 Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(8).Value) = True Then
                            ExpiryDate = CDate(.Cells(8).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(5).Value) = True Then
                            SecondaryQty = .Cells(5).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = 3 Then
                            Me.EndEditQtyAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 5 Then
                            Me.EndEditSecondaryUnitAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 7 Then
                            Me.EndEditBatchNoAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 9 Then
                            Me.EndEditLocationAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 8 Then
                            Me.EndEditExpiryDateAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 16 Then
                            Me.EndEditCostFactor(sender, e, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = 13 Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditExpiryDateAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan format yang dituliskan adalah format tanggal
                If IsDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'format ulang tanggal sehingga formatnya konsisten
                Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value), _
                "dd MMM yyyy")
                'cek apakah back date terhadap tanggal transaksi, konfirmasikan data ini
                If CDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'pastikan tanggal expiry date tidak konflik dengan data yang sudah ada
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'bandingkan dengan yang sudah ada didalam dgv
                Dim row1 As DataGridViewRow
                For Each row1 In Me.dgvItemAfter.Rows
                    If row1.Cells(7).Value = row.Cells(7).Value And CStr(row.Cells(7).Value) <> "" _
                    And row1.Index <> row.Index And row1.Cells(12).Value = row.Cells(12).Value Then
                        If CDate(row1.Cells(8).Value).Date <> CDate(row.Cells(8).Value).Date Then
                            MsgBox("You already entered expiry date for batch no: " + CStr(row1.Cells(7).Value) + " on " + _
                            Format(CDate(row1.Cells(8).Value), "dd MMM yyyy"), MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                'bandingkan dengan yang sudah ada didalam transaksi lain. Detail lihat point no 8.c pada 
                'fungsi apakah data item benar
                'fungsi apakah data item benar
                '1. stock in (semua stock in)
                Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                For loopInOut As Integer = 0 To 4 Step 1
                    Select Case loopInOut
                        Case 0
                            '   a. stock in purchase
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                            Exit Select
                        Case 1
                            '   b. stock in sales return
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            Exit Select
                        Case 2
                            '   c. stock in adjust
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                            Exit Select
                        Case 3
                            '   d. stock in back
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                            Exit Select
                        Case 4
                            '   e. stock in proses
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                            Exit Select
                    End Select
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                    Me.dgvItemAfter.Item(7, e.RowIndex).Value, JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value) Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")
                'expiry date, tidak perlu cek kemungkinan merger, karena tidak mungkin
                'hanya expiry date saja yang berbeda sedangkan batch no-nya sama
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, CDate(Me.buffer), SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCostFactor(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for cost factor data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'digit grouping
                Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value, 2)
                'cek nilai cost factor maximum adalah 100
                If CDbl(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    MsgBox("The maximum value is 100%.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQtyAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'digit grouping
                Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value, _
                Me.ObjProduct.GetAngkaDecimal(Me.dgvItemAfter.Item(Me.ColPrimaryCodeNoAfter.Index, e.RowIndex).Value))
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                JmlDGV = .Cells(3).Value
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnitAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemAfter)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemAfter)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNoAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                If IsNothing(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                    End If
                End If
                'batch no selalu huruf besar
                Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper.Trim
                'pastikan diisikan
                If CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                Me.buffer, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value) = True Then
                        'basis data
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                        Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(8, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value)
                                    Me.dgvItemAfter.Item(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If

                        'cek juga ke transaksi:
                        '1. stock in (semua stock in)
                        Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                        For loopInOut As Integer = 0 To 4 Step 1
                            Select Case loopInOut
                                Case 0
                                    '   a. stock in purchase
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                                    Exit Select
                                Case 1
                                    '   b. stock in sales return
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                                    Exit Select
                                Case 2
                                    '   c. stock in adjust
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                                    Exit Select
                                Case 3
                                    '   d. stock in back
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                                    Exit Select
                                Case 4
                                    '   e. stock in proses
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                                    Exit Select
                            End Select
                            ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                            Me.dgvItemAfter.Item(7, e.RowIndex).Value, JenisTransaksi)
                            If ExpiryDateLama <> #1/1/1950# Then
                                If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(8, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                    Format(ExpiryDateLama, "dd MMM yyyy") + _
                                    ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    Else
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItemAfter(8, e.RowIndex).Value)
                                        Me.dgvItemAfter(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                    End If
                                End If
                            End If
                        Next
                        '2. transaksi stock bundling untuk item hasil bundling
                        ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                        Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(8, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItemAfter(8, e.RowIndex).Value)
                                    Me.dgvItemAfter(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        '3. proses stock taking
                        'msgbox("belum")

                        'data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItemAfter.Rows
                            If row1.Cells(7).Value = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells(12).Value = Me.dgvItemAfter.Item(12, e.RowIndex).Value Then
                                If CDate(row1.Cells(8).Value) <> CDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                        Format(CDate(row1.Cells(8).Value), "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value)
                                        Me.dgvItemAfter.Item(8, e.RowIndex).Value = row1.Cells(8).Value
                                        Exit For
                                    Else
                                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '3. cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemAfter)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemAfter.Item(8, e.RowIndex).Value = Format(CDate(BufferExpiryDate), "dd MMM yyyy")
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemAfter)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocationAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                Dim Obj As New clsMasterStorageArea
                If Obj.ApakahSudahAda(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterStorageArea
                    If CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) <> "" Then
                        i.txtCari.Text = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea <> -1 Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetStorageAreaCode(i.PrimaryStorageArea)).ToUpper
                    Else
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        i = Nothing
                        Exit Sub
                    End If
                    i = Nothing
                Else
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                End If
                Obj = Nothing
                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'update data primary location
                    Dim Obj1 As New clsMasterStorageArea
                    .Cells(13).Value = Obj1.GetPrimary(.Cells(9).Value)
                    'update data gudang
                    .Cells(10).Value = Obj1.GetWarehouseCode(.Cells(13).Value)

                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemAfter)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'update primary location
                            .Cells(13).Value = Obj1.GetPrimary(.Cells(9).Value)
                            'update data gudang
                            .Cells(10).Value = Obj1.GetWarehouseCode(.Cells(13).Value)
                            Obj1 = Nothing
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemAfter)
                    End If
                    Obj1 = Nothing
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSahAfter(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
   ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
   ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockBundlingItem.GetJmlAfter(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)

            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary, Me.dgvItemAfter)
            JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV

            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Minimum value you must can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlObjLama - JmlGudang), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + _
                    " " + Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemAfter_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemAfter.RowsAdded
        Me.dgvItemAfter.TabStop = True
        If Me.TxtRemark.Enabled = True Then
            Me.btnUpAfter.Enabled = True
            Me.btnDownAfter.Enabled = True
        End If
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemAfter.RowsRemoved
        If Me.dgvItemAfter.Rows.Count = 0 Then
            Me.dgvItemAfter.TabStop = False
            If Me.TxtRemark.Enabled = True Then
                Me.btnUpAfter.Enabled = False
                Me.btnDownAfter.Enabled = False
            End If
        End If
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
                    Me.txtPrimaryLocationAfter.Text = obj.GetPrimary(Me.TxtLocationAfter.Text)
                    Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                    'Me.TxtQtyperBatchAfter.Focus()
                Else
                    If Me.TxtLocationAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocationAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocationAfter.Text = i.PrimaryStorageArea
                        Me.TxtLocationAfter.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                        ' Me.TxtQtyperBatchAfter.Focus()
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

    Private Sub dgvItemAfter_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemAfter.SelectionChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.dgvItemAfter.SelectedRows.Count > 0 Then
                    Me.btnUpAfter.Enabled = True
                    Me.btnDownAfter.Enabled = True
                Else
                    Me.btnUpAfter.Enabled = False
                    Me.btnDownAfter.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItemAfter.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemAfter)
    End Sub

    Private Sub btnUPAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpAfter.Click
        Try
            If Me.dgvItemAfter.SelectedRows.Count > 0 Then
                Me.btnUpAfter.Enabled = False
                Me.btnDownAfter.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemAfter.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemAfter.Rows.Remove(row)
                    Me.dgvItemAfter.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemAfter.Focus()
                    Me.dgvItemAfter.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemAfter)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True, Me.dgvItemAfter)
                    End If
                End If
                Me.btnUpAfter.Enabled = True
                Me.btnDownAfter.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDownAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownAfter.Click
        Try
            If Me.dgvItemAfter.SelectedRows.Count > 0 Then
                Me.btnUpAfter.Enabled = False
                Me.btnDownAfter.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemAfter.SelectedRows(0)
                If row.Index <> Me.dgvItemAfter.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemAfter.Rows.Remove(row)
                    Me.dgvItemAfter.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemAfter.Focus()
                    Me.dgvItemAfter.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemAfter)
                    Me.ShowRow(row, False, Me.dgvItemAfter)
                End If
                Me.btnUpAfter.Enabled = True
                Me.btnDownAfter.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemAfter_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemAfter.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemAfter)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemAfter.UserDeletingRow
        Try
            'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
            Dim PrimaryLocation, PrimaryPdO, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
            JmlDGV, JmlSetelahHapus As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = e.Row

            With row
                PrimaryLocation = .Cells(13).Value
                PrimaryPdO = -1
                PrimaryProduct = .Cells(12).Value
                BatchNo = .Cells(7).Value
                If IsDate(.Cells(8).Value) = True Then
                    ExpiryDate = CDate(.Cells(8).Value)
                Else
                    ExpiryDate = #1/1/1950#
                End If
                If IsNumeric(.Cells(5).Value) = True Then
                    SecondaryQty = .Cells(5).Value
                Else
                    SecondaryQty = -1
                End If
                JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                'karena akan dihapus, perhitungan menggunakan nilai jumlah dgv =0
                JmlDGV = 0
                JmlObjLama = Me.objStockBundlingItem.GetJmlAfter(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
                SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
                JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
                If JmlSetelahHapus < 0 Then
                    MsgBox("You cannot delete this data, because this will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
            End With
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
            If Me.TxtRemark.Enabled = True Then
                Me.txtBatchNoAfter.Text = Me.txtBatchNoAfter.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
            End If
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
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNoAfter.Text <> "" And Me.TxtPrimaryCodeNoAfter.Text <> "" Then
                    Me.TxtQtyperBatchAfter.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchAfter.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNoAfter.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjStockBundling.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjStockBundling.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtStockInNo.Text = Me.ObjStockBundling.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtStockInNo.Text = Me.ObjStockBundling.GetNoBaru(Me.dtpDate.Value, ClsTransaksiStockBundling.EnumTipeBundling.Purchase1)
                    End If
                Else
                    Me.TxtStockInNo.Text = Me.ObjStockBundling.GetNoBaru(Me.dtpDate.Value, ClsTransaksiStockBundling.EnumTipeBundling.Purchase1)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtQtyperBatchb4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNob4.Text <> "" And Me.TxtPrimaryCodeNob4.Text <> "-1" Then
                    Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchb4.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TabControl1.SelectedIndex = 0 Then
                    ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
                    ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
                    Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
                Else
                    ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
                    ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
                    Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
                End If
            Else
                If Me.TabControl1.SelectedIndex = 0 Then
                    ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
                    ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
                Else
                    ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
                    ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.BackColor, False)
                End If
            End If
            Me.SetBtnUpDown()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetBtnUpDown()
        Try
            If Me.TabControl1.SelectedTab.Name = Me.tabB4.Name Then
                Me.btnUpAfter.Visible = False
                Me.btnDownAfter.Visible = False
                Me.btnUPb4.Visible = True
                Me.btnDownb4.Visible = True
            Else
                Me.btnUpAfter.Visible = True
                Me.btnDownAfter.Visible = True
                Me.btnUPb4.Visible = False
                Me.btnDownb4.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddByBillOfMaterial()
        Dim ijk As New FrmTransaksiStockBundling_BOM

        Try
            If Me.TxtRemark.Enabled = True Then
                ijk.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan nilai true, jika ada salah satu rm dari item2 after yang location control
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahAda_ItemRM_yg_LocationControl() As Boolean
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.DataTable7TableAdapter
        Dim TBL As New dseTransaksiStockBundling.DataTable7DataTable
        Dim row As dseTransaksiStockBundling.DataTable7Row
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowDGV In Me.dgvItemAfter.Rows
                DTA.Fill(TBL, rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value)
                For Each row In TBL.Rows
                    If row.Location_Controlled = 1 Then
                        Return True
                    End If
                Next
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub llb_BOM_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llb_BOM.LinkClicked
        Dim ijk As New Core_Module.FrmMasterLocationPriority
        Dim Lanjut As Boolean = False

        Try
            If Me.TxtRemark.Enabled = True And Me.dgvItemAfter.Rows.Count > 0 Then
                If MsgBox("Add item before by bill of material?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    If Me.ApakahAda_ItemRM_yg_LocationControl = True Then
                        'isikan prioritas item before
                        ijk.BackColorAll = Me.BackColorAll
                        ijk.UserCode = Me.UserCode
                        ijk.vTipeTransaksi = ClsMasterLocationPriority.EnumTipeTransaksi.StockBundling
                        If ijk.ShowDialog() = Windows.Forms.DialogResult.OK Then
                            Lanjut = True
                            Me.ObjLocationPriority.GetLocationPrioriry(Me.TBLLocationPriority, Me.UserCode, ClsMasterLocationPriority.EnumTipeTransaksi.StockBundling)
                        Else
                            Lanjut = False
                        End If
                    Else
                        Lanjut = True
                    End If

                    If Lanjut = True Then
                        'kosongkan item before
                        Me.dgvItemb4.Rows.Clear()
                        Me.KosongkanItemB4()
                        Me.SetGboItemb4()
                        Me.TabControl1.SelectedIndex = 0

                        For Each rowDGV As DataGridViewRow In Me.dgvItemAfter.Rows
                            Me.AddItemRawMaterial2(rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value, rowDGV.Cells(Me.ColQtyAfter.Index).Value, Me.TBLLocationPriority)
                        Next

                        Me.KosongkanItemB4()
                        Me.SetGboItemb4()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk = Nothing
        End Try
    End Sub

    Private Sub IsikanDataSesuaiPrioritas()
        Dim rowLocation As dseMasterProduct.stock_location_viewRow
        Dim idx As Integer = 0
        Dim StartIdx As Integer = 0
        Dim Lanjut As Boolean = True

        Try
            'menentukan data yg diisikan berdasarkan location priority
            If Me.cboBatchNob4.Visible = True Then
                'kombinasi batch no + expiry + location
                'pilih berdasarkan location priority, jika ada 2 item barang dengan lokasi yg sama tetapi expiry yang berbeda
                'maka pilih yang expiry paling dekat terlebih dahulu
                'jika sudah tidak termasuk location priority, maka tetap pilih yang expiry-nya palin dekat terlebih dahulu, alias tidak ada perubahan

                '1. catat index lama
                StartIdx = Me.cboBatchNob4.SelectedIndex
                '2. start looping
                While Lanjut = True
                    For Each rowPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row In Me.TBLLocationPriority.Rows
                        idx = 0
                        For Each DV As Data.DataRowView In Me.TBL2.DefaultView
                            rowLocation = DV.Row
                            If rowLocation.Location = rowPriority.Storage_Area_Code Then
                                Me.cboLocationb4.SelectedIndex = idx
                                Exit Sub
                            End If
                            idx += 1
                        Next
                    Next
                    If Me.cboBatchNob4.SelectedIndex + 1 <= Me.cboBatchNob4.Items.Count - 1 Then
                        'jika ada batch selanjutnya, maka cek untuk batch selanjutnya
                        Lanjut = True
                        Me.cboBatchNob4.SelectedIndex += 1
                    Else
                        'jika tidak ada batch selanjutnya maka kembalikan ke batch awal
                        Lanjut = False
                        Me.cboBatchNob4.SelectedIndex = StartIdx
                    End If
                End While
            Else
                'hanya location control
                'pilih lokasi sesuai dengan location priority, jika ada
                'jika tidak ada, maka pilih location berdasarkan abjad saja, alias tidak di ubah
                For Each rowPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row In Me.TBLLocationPriority.Rows
                    idx = 0
                    For Each DV As Data.DataRowView In Me.TBL2.DefaultView
                        rowLocation = DV.Row
                        If rowLocation.Location = rowPriority.Storage_Area_Code Then
                            Me.cboLocationb4.SelectedIndex = idx
                            Exit Sub
                        End If
                        idx += 1
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanItemQty(ByVal vTBLProductQty As DseTrans_StockBundling.DataTable1DataTable, _
   ByVal vTBLItemBundling As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before4DataTable, _
   ByVal vPrimaryProduct As Double)
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim rowItemBundling As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before4Row
        Dim DV As Data.DataRowView
        Dim ketemu As Boolean = False
        Dim CodeNo As String = ""
        Dim Desc As String = ""
        Dim DecimalPoint As Integer = 0
        Dim PrimaryUnit As Double = -1
        Dim UnitCode As String = ""

        Try
            'tambahkan dengan item db
            vTBLItemBundling.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString
            For Each DV In vTBLItemBundling.DefaultView
                rowItemBundling = DV.Row
                ketemu = False

                For Each rowProductQty In vTBLProductQty.Rows
                    If rowProductQty.Lot_No = rowItemBundling.Batch_No And rowProductQty.Expiry_Date = rowItemBundling.Expiry_Date And _
                    rowProductQty.Primary_Storage_Area = rowItemBundling.Primary_Storage_Area Then
                        ketemu = True
                        rowProductQty.Qty += rowItemBundling.Qty
                        Exit For
                    End If
                Next

                If ketemu = False Then
                    If CodeNo = "" Then
                        CodeNo = Me.ObjProduct.GetProductCode(rowItemBundling.Primary_Product)
                        Desc = Me.ObjProduct.GetDescription(rowItemBundling.Primary_Product)
                        DecimalPoint = Me.ObjProduct.GetAngkaDecimal(rowItemBundling.Primary_Product)
                        PrimaryUnit = Me.ObjProduct.GetPrimaryUnit(rowItemBundling.Primary_Product)
                        UnitCode = Me.ObjProduct.GetUnitCode(rowItemBundling.Primary_Product)
                    End If
                    rowProductQty = vTBLProductQty.NewDataTable1Row
                    With rowProductQty
                        .Description = Desc
                        .Product_Code = CodeNo
                        .Expiry_Date = rowItemBundling.Expiry_Date
                        .Lot_No = rowItemBundling.Batch_No
                        .Primary_Product = rowItemBundling.Primary_Product
                        .Primary_Storage_Area = rowItemBundling.Primary_Storage_Area
                        .Qty = rowItemBundling.Qty
                        ._Decimal = DecimalPoint
                        .Primary_Unit = PrimaryUnit
                        .Unit_Code = UnitCode
                        .Storage_Area_Code = rowItemBundling.Storage_Area_Code
                        .Warehouse_Code = rowItemBundling.Warehouse_Code
                    End With
                    vTBLProductQty.Rows.Add(rowProductQty)
                End If
            Next

            'kurangkan dengan dgv
            For Each rowDGV As DataGridViewRow In Me.dgvItemb4.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    ketemu = False
                    For Each rowProductQty In vTBLProductQty.Rows
                        If rowProductQty.Lot_No = rowDGV.Cells(Me.ColBatch.Index).Value And _
                        rowProductQty.Primary_Storage_Area = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value Then
                            ketemu = True
                            rowProductQty.Qty -= CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                            Exit For
                        End If
                    Next

                    If ketemu = False Then
                        'tambahkan baris dengan qty minus
                        If CodeNo = "" Then
                            CodeNo = rowDGV.Cells(Me.ColCodeNo.Index).Value
                            Desc = rowDGV.Cells(Me.ColDescription.Index).Value
                            DecimalPoint = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct)
                            PrimaryUnit = Me.ObjProduct.GetPrimaryUnit(vPrimaryProduct)
                            UnitCode = Me.ObjProduct.GetUnitCode(vPrimaryProduct)
                        End If
                        rowProductQty = vTBLProductQty.NewDataTable1Row
                        With rowProductQty
                            .Description = Desc
                            .Product_Code = CodeNo
                            If IsDate(rowDGV.Cells(Me.ColExpiredDate.Index).Value) = True Then
                                .Expiry_Date = rowDGV.Cells(Me.ColExpiredDate.Index).Value
                            Else
                                .Expiry_Date = #1/1/1950#
                            End If
                            .Lot_No = rowDGV.Cells(Me.ColBatch.Index).Value
                            .Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                            .Primary_Storage_Area = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value
                            .Qty = CDec(rowDGV.Cells(Me.ColQty.Index).Value) * (-1)
                            ._Decimal = DecimalPoint
                            .Primary_Unit = PrimaryUnit
                            .Unit_Code = UnitCode
                            .Storage_Area_Code = rowDGV.Cells(Me.ColLocation.Index).Value
                            .Warehouse_Code = rowDGV.Cells(Me.ColWareHouse.Index).Value
                        End With
                        vTBLProductQty.Rows.Add(rowProductQty)
                    End If
                End If
            Next

            For ij As Integer = vTBLProductQty.Rows.Count - 1 To 0 Step -1
                rowProductQty = vTBLProductQty.Rows(ij)
                If rowProductQty.Qty <= 0 Then
                    vTBLProductQty.Rows.Remove(rowProductQty)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrioritasStrictLocation(ByVal vTBLProductQty As DseTrans_StockBundling.DataTable1DataTable, _
    ByVal vTBLLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable)
        Dim TblDummy As New DseTrans_StockBundling.DataTable1DataTable
        Dim rowDummy As DseTrans_StockBundling.DataTable1Row
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim RowLocation As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row
        Dim DV As Data.DataRowView

        Try
            'pindahkan ke tabel dummy
            vTBLProductQty.DefaultView.Sort = "[Expiry Date], [Lot No]"
            For Each RowLocation In vTBLLocationPriority.Rows
                vTBLProductQty.DefaultView.RowFilter = "[Primary Storage Area] = " + RowLocation.primary_location.ToString

                For Each DV In vTBLProductQty.DefaultView
                    rowProductQty = DV.Row
                    rowDummy = TblDummy.NewDataTable1Row
                    For ij As Integer = 0 To TblDummy.Columns.Count - 1
                        rowDummy.Item(ij) = rowProductQty.Item(ij)
                    Next
                    TblDummy.Rows.Add(rowDummy)
                Next

                While vTBLProductQty.DefaultView.Count > 0
                    rowProductQty = vTBLProductQty.DefaultView(0).Row
                    vTBLProductQty.Rows.Remove(rowProductQty)
                End While
            Next

            'pindahkan sisanya ke tabel dummy
            For Each rowProductQty In vTBLProductQty.Rows
                rowDummy = TblDummy.NewDataTable1Row
                For ij As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ij) = rowProductQty.Item(ij)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next

            'kembali pindahkan ke tabel utama
            vTBLProductQty.Rows.Clear()
            For Each rowDummy In TblDummy.Rows
                rowProductQty = vTBLProductQty.NewDataTable1Row
                For ij As Integer = 0 To vTBLProductQty.Columns.Count - 1
                    rowProductQty.Item(ij) = rowDummy.Item(ij)
                Next
                vTBLProductQty.Rows.Add(rowProductQty)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterStrictLocation(ByVal vTBLProductQty As DseTrans_StockBundling.DataTable1DataTable, _
    ByVal vTBLLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable)
        Dim TblDummy As New DseTrans_StockBundling.DataTable1DataTable
        Dim rowDummy As DseTrans_StockBundling.DataTable1Row
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim RowLocation As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row
        Dim DV As Data.DataRowView

        Try
            'pindahkan ke tabel dummy
            vTBLProductQty.DefaultView.Sort = "[Expiry Date], [Lot No]"
            For Each RowLocation In vTBLLocationPriority.Rows
                vTBLProductQty.DefaultView.RowFilter = "[Primary Storage Area] = " + RowLocation.primary_location.ToString

                For Each DV In vTBLProductQty.DefaultView
                    rowProductQty = DV.Row
                    rowDummy = TblDummy.NewDataTable1Row
                    For ij As Integer = 0 To TblDummy.Columns.Count - 1
                        rowDummy.Item(ij) = rowProductQty.Item(ij)
                    Next
                    TblDummy.Rows.Add(rowDummy)
                Next
            Next

            'kembali pindahkan ke tabel utama
            vTBLProductQty.Rows.Clear()
            For Each rowDummy In TblDummy.Rows
                rowProductQty = vTBLProductQty.NewDataTable1Row
                For ij As Integer = 0 To vTBLProductQty.Columns.Count - 1
                    rowProductQty.Item(ij) = rowDummy.Item(ij)
                Next
                vTBLProductQty.Rows.Add(rowProductQty)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy.Dispose()
        End Try
    End Sub

    Private Sub AddItemRawMaterial2(ByVal vPrimaryProduct As Double, ByVal vQtyAfter As Decimal, _
    ByVal vTBLLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable)
        Dim ij As Integer = 0
        Dim QtyEntry As Decimal = 0
        'Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
        Dim Dummy As String = ""
        Dim GoodPercentage As Decimal = 0
        Dim DummyQty1 As Decimal = 0
        Dim NST As Decimal = 0
        Dim ObjUSER As New Core_Module.ClsUser
        Dim LocationControl As Boolean = True
        Dim DTAProductQty As New DseTrans_StockBundlingTableAdapters.DataTable1TableAdapter
        Dim TBLProductQty As New DseTrans_StockBundling.DataTable1DataTable
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim DTAItemBundling As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before4TableAdapter
        Dim TBLItemBundling As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before4DataTable
        Dim StrictLocation As Boolean = False
        Dim jk As Integer = 0

        Try
            Me.SedangAddRM = True

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_NoStrictLocation") = True Then
                StrictLocation = False
            Else
                StrictLocation = True
            End If

            DTAProductQty.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemBundling.Connection.ConnectionString = DTAProductQty.Connection.ConnectionString

            Me.dgvItemb4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemb4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'isikan tblrm berdasarkan bill of material
            Me.ObjProduct.Isikan_TBLRM(vPrimaryProduct, Me.TBL_BOM)
            'kalikan qty raw berdasarkan good percentage umum, good percentage detail, qty hasil
            For Each Me.row_BOM In Me.TBL_BOM.Rows
                'kalikan good percentage item
                DummyQty1 = Math.Round((((((Me.row_BOM.qty_raw * 100) / Me.row_BOM.good_percentage) * 100) / Me.row_BOM.Good_Percentage_General) * vQtyAfter) / Me.row_BOM.Qty_Ref_Product, _
                Me.row_BOM._Decimal, MidpointRounding.AwayFromZero)

                Me.row_BOM.qty_raw = DummyQty1
            Next

            'input ke dalam product, berdasarkan location priority
            TBLItemBundling.Rows.Clear()
            If Me.baru = True Then
                DTAItemBundling.Fill(TBLItemBundling, Me.TxtPrimaryStockIn.Text)
            End If
            Me.KosongkanItemB4()

            While ij < Me.TBL_BOM.Rows.Count
                Me.row_BOM = Me.TBL_BOM.Rows(ij)
                LocationControl = Me.ObjProduct.ApakahLocationControlled(Me.row_BOM.primary_product_raw)

                Me.ObjProduct.IsiCheckBoxControll(Me.row_BOM.primary_product_raw, Me.chkLocationControlledb4, _
                Me.chkBatchNoControlledb4, Me.chkExpiryDateControlledb4, Me.chkUseSecondaryUnitb4, _
                Me.chkPdoControlledb4, Me.txtPrimaryUnitsb4, Me.txtPrimarySecondaryUnitsb4)

                'load item product qty
                DTAProductQty.Fill(TBLProductQty, Me.row_BOM.primary_product_raw)
                Me.IsikanItemQty(TBLProductQty, TBLItemBundling, Me.row_BOM.primary_product_raw)

                If LocationControl = True And StrictLocation = True Then
                    'jika location control dan strict location, maka buang yang lokasinya tidak ada dalam tabel location control
                    Me.FilterStrictLocation(TBLProductQty, vTBLLocationPriority)
                ElseIf LocationControl = True And StrictLocation = False Then
                    'jika location control dan tidak strict location, maka pindahkan lokasi yang menjadi prioritas ke baris paling atas
                    Me.PrioritasStrictLocation(TBLProductQty, vTBLLocationPriority)
                End If

                'dari baris paling atas tabel item product qty
                jk = 0
                While Me.row_BOM.qty_raw > 0 And jk <= TBLProductQty.Rows.Count - 1
                    rowProductQty = TBLProductQty.Rows(jk)
                    If Me.row_BOM.qty_raw > rowProductQty.Qty Then
                        'masukkan ke dgv b4
                        Me.TxtPrimaryCodeNob4.Text = rowProductQty.Primary_Product
                        Me.TxtCodeNob4.Text = rowProductQty.Product_Code
                        Me.txtDescb4.Text = rowProductQty.Description
                        Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(rowProductQty.Qty, rowProductQty._Decimal)
                        Me.txtPrimaryUnitsb4.Text = rowProductQty.Primary_Unit
                        Me.txtPrimaryLocationb4.Text = rowProductQty.Primary_Storage_Area
                        Me.cboLocationb4.Items.Clear()
                        Me.cboLocationb4.Items.Add(rowProductQty.Storage_Area_Code)
                        Me.cboLocationb4.Text = rowProductQty.Storage_Area_Code
                        Me.TxtWarehouseb4.Text = rowProductQty.Warehouse_Code
                        Me.txtSecondaryQtyb4.Text = "-1"
                        Me.lblUnitsb4.Text = rowProductQty.Unit_Code
                        Me.cboBatchNob4.Items.Clear()
                        Me.cboBatchNob4.Items.Add(rowProductQty.Lot_No)
                        Me.cboBatchNob4.Text = rowProductQty.Lot_No
                        Me.DtpExpiredb4.Value = rowProductQty.Expiry_Date
                        Me.InputB4(False, False)

                        'edit qty rm item bom
                        Me.row_BOM.qty_raw -= rowProductQty.Qty

                        'ambil semua item product qty
                        rowProductQty.Qty = 0

                        'pindah ke baris item product qty selanjutnya
                        jk += 1
                    Else
                        'masukkan ke dgv b4
                        Me.TxtPrimaryCodeNob4.Text = rowProductQty.Primary_Product
                        Me.TxtCodeNob4.Text = rowProductQty.Product_Code
                        Me.txtDescb4.Text = rowProductQty.Description
                        Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(Me.row_BOM.qty_raw, rowProductQty._Decimal)
                        Me.txtPrimaryUnitsb4.Text = rowProductQty.Primary_Unit
                        Me.lblUnitsb4.Text = rowProductQty.Unit_Code
                        Me.txtSecondaryQtyb4.Text = "-1"
                        Me.txtPrimaryLocationb4.Text = rowProductQty.Primary_Storage_Area
                        Me.cboLocationb4.Items.Clear()
                        Me.cboLocationb4.Items.Add(rowProductQty.Storage_Area_Code)
                        Me.cboLocationb4.Text = rowProductQty.Storage_Area_Code
                        Me.TxtWarehouseb4.Text = rowProductQty.Warehouse_Code

                        Me.cboBatchNob4.Items.Clear()
                        Me.cboBatchNob4.Items.Add(rowProductQty.Lot_No)
                        Me.cboBatchNob4.Text = rowProductQty.Lot_No
                        Me.DtpExpiredb4.Value = rowProductQty.Expiry_Date
                        Me.InputB4(False, False)

                        'edit qty item product qty
                        rowProductQty.Qty -= Me.row_BOM.qty_raw

                        'ambil semua qty item bom
                        Me.row_BOM.qty_raw = 0
                    End If
                End While

                ij += 1
            End While

            Me.ObjCD.SetNumber(Me.dgvItemb4, False)
            'kosongkan gbo item
            Me.KosongkanItemB4()
            Me.TxtCodeNob4.Focus()
            'format dgv
            Me.FormatDGV(Me.dgvItemb4)
            'fokuskan ke baris terakhir
            If Me.dgvItemb4.RowCount > 0 Then
                While Me.dgvItemb4.Rows(Me.dgvItemb4.Rows.Count - 1).Displayed = False _
                Or Me.dgvItemb4.DisplayedRowCount(True) <> Me.dgvItemb4.DisplayedRowCount(False)
                    Me.dgvItemb4.FirstDisplayedScrollingRowIndex = Me.dgvItemb4.FirstDisplayedScrollingRowIndex + 1
                End While
            End If

            'berikan peringatan bila qty raw material yang dimasukkan tidak cukup
            For ij = TBL_BOM.Rows.Count - 1 To 0 Step -1
                Me.row_BOM = TBL_BOM.Rows(ij)
                If Me.row_BOM.qty_raw <= 0 Then
                    TBL_BOM.Rows.Remove(Me.row_BOM)
                End If
            Next
            If Me.TBL_BOM.Rows.Count > 0 Then
                jk = 0
                For Each Me.row_BOM In Me.TBL_BOM.Rows
                    jk += 1
                    If Dummy = "" Then
                        Dummy = jk.ToString + ". " + Me.row_BOM.RM_Code + " (" + Me.ObjCD.DigitGrouping(Me.row_BOM.qty_raw, Me.row_BOM._Decimal) + " " + Me.row_BOM.Unit_Code + ")"
                    Else
                        Dummy = Dummy + vbCrLf + jk.ToString + ". " + Me.row_BOM.RM_Code + " (" + Me.ObjCD.DigitGrouping(Me.row_BOM.qty_raw, Me.row_BOM._Decimal) + " " + Me.row_BOM.Unit_Code + ")"
                    End If
                Next
                MsgBox("Not enough raw material: " + vbCrLf + Dummy + ".", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUSER = Nothing
            DTAProductQty.Dispose()
            TBLProductQty.Dispose()
            DTAItemBundling.Dispose()
            TBLItemBundling.Dispose()
            Me.SedangAddRM = False
        End Try
    End Sub

    'Private Sub AddItemRawMaterial(ByVal vPrimaryProduct As Double, ByVal vQtyAfter As Decimal)
    '    Dim ij As Integer = 0
    '    Dim QtyEntry As Decimal = 0
    '    Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
    '    Dim Dummy As String = ""
    '    Dim GoodPercentage As Decimal = 0
    '    Dim DummyQty As Decimal = 0
    '    Dim DummyQty1 As Decimal = 0
    '    Dim NST As Decimal = 0

    '    Try

    ''isikan tblrm berdasarkan bill of material
    '        Me.ObjProduct.Isikan_TBLRM(vPrimaryProduct, Me.TBL_BOM)
    ''kalikan qty raw berdasarkan good percentage umum, good percentage detail, qty hasil
    '        For Each Me.row_BOM In Me.TBL_BOM.Rows
    ''kalikan good percentage item
    '            DummyQty = (((((Me.row_BOM.qty_raw * 100) / Me.row_BOM.good_percentage) * 100) / Me.row_BOM.Good_Percentage_General) * vQtyAfter) / Me.row_BOM.Qty_Ref_Product
    '            DummyQty1 = Math.Round((((((Me.row_BOM.qty_raw * 100) / Me.row_BOM.good_percentage) * 100) / Me.row_BOM.Good_Percentage_General) * vQtyAfter) / Me.row_BOM.Qty_Ref_Product, _
    '            Me.row_BOM._Decimal, MidpointRounding.AwayFromZero)

    '            If DummyQty1 >= DummyQty Then
    '                Me.row_BOM.qty_raw = DummyQty1
    '            Else
    '                NST = (1 / (10 ^ Me.row_BOM.Decimal_Product)) / 2
    '                DummyQty1 += NST
    '                Me.row_BOM.qty_raw = Math.Round(DummyQty1, Me.row_BOM._Decimal, MidpointRounding.AwayFromZero)
    '            End If
    '        Next

    '        'input ke dalam product
    '        While ij < Me.TBL_BOM.Rows.Count
    '            Me.row_BOM = Me.TBL_BOM.Rows(ij)
    '            Me.TxtPrimaryCodeNob4.Text = Me.row_BOM.primary_product_raw
    '            Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
    '            Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
    '            Me.SetGboItemb4()

    '            If Me.cboBatchNob4.Visible = True Then
    '                Me.cboBatchNob4.Focus()
    '            ElseIf Me.cboLocationb4.Visible = True Then
    '                Me.cboLocationb4.Focus()
    '            Else
    '                Me.TxtQtyperBatchb4.Focus()
    '            End If

    '            If Me.cboLocationb4.Visible = True Then
    '                Me.IsikanDataSesuaiPrioritas()
    '            End If

    '            'tentukan qty yang dimasukkan
    '            JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
    '            Me.cboBatchNob4.Text)
    '            'perhitungan menggunakan nilai jumlah dgv yang baru
    '            JmlObjLama = Me.objStockBundlingItem.GetJmlB4(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNob4.Text, _
    '            Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, Me.DtpExpiredb4.Value)
    '            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
    '            Me.cboBatchNob4.Text, Me.chkBatchNoControlledb4.Checked, _
    '            Me.chkLocationControlledb4.Checked, Me.chkUseSecondaryUnitb4.Checked, Me.dgvItemb4)
    '            JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text)
    '            QtyEntry = JmlGudang + JmlObjLama - JmlDGV
    '            If QtyEntry > Me.row_BOM.qty_raw Then
    '                QtyEntry = Me.row_BOM.qty_raw
    '            End If
    '            If QtyEntry <> 0 Then
    '                Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(QtyEntry, Me.row_BOM._Decimal)
    '                Me.row_BOM.qty_raw -= QtyEntry
    '                Me.BtnInputb4.PerformClick()
    '                If Me.row_BOM.qty_raw = 0 Then
    '                    Me.TBL_BOM.Rows.Remove(Me.row_BOM)
    '                End If
    '            Else
    '                'tidak ada lagi yang bisa diinput utnuk rm ybs, pindah ke baris selanjutnya
    '                'kasus terjadi misalnya pada qty raw yang diminta lebih banyak dari qty rm yang ada
    '                ij += 1
    '            End If
    '        End While

    '        'berikan peringatan bila qty raw material yang dimasukkan tidak cukup
    '        If Me.TBL_BOM.Rows.Count > 0 Then
    '            For Each Me.row_BOM In Me.TBL_BOM.Rows
    '                If Dummy = "" Then
    '                    Dummy = Me.row_BOM.RM_Code + "(" + Me.ObjCD.DigitGrouping(Me.row_BOM.qty_raw, Me.row_BOM._Decimal) + " " + Me.row_BOM.Unit_Code + ")"
    '                Else
    '                    Dummy = Dummy + ", " + Me.row_BOM.RM_Code + "(" + Me.ObjCD.DigitGrouping(Me.row_BOM.qty_raw, Me.row_BOM._Decimal) + " " + Me.row_BOM.Unit_Code + ")"
    '                End If
    '            Next
    '            MsgBox("Not enough raw material: " + Dummy + ".", MsgBoxStyle.Information, "Information")
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'iwan: 2016-04-10
    'obselete version of the above coding
    'Private Sub AddItemRawMaterial(ByVal vPrimaryProduct As Double, ByVal vQtyAfter As Decimal)
    '    Dim ij As Integer = 0
    '    Dim QtyEntry As Decimal = 0
    '    Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
    '    Dim Dummy As String = ""
    '    Dim GoodPercentage As Decimal = 0
    '    Dim DummyQty As Decimal = 0
    '    Dim DummyQty1 As Decimal = 0
    '    Dim NST As Decimal = 0

    '    Try

    '        'isikan tblrm berdasarkan bill of material
    '        Me.ObjProduct.Isikan_TBLRM(vPrimaryProduct, Me.TBL_BOM)
    '        'kalikan qty raw berdasarkan good percentage umum, good percentage detail, qty hasil
    '        For Each Me.row_BOM In Me.TBL_BOM.Rows
    '            'kalikan good percentage item
    '            DummyQty = (((((Me.row_BOM.qty_raw * 100) / Me.row_BOM.good_percentage) * 100) / Me.row_BOM.Good_Percentage_General) * vQtyAfter) / Me.row_BOM.Qty_Ref_Product
    '            DummyQty1 = Math.Round((((((Me.row_BOM.qty_raw * 100) / Me.row_BOM.good_percentage) * 100) / Me.row_BOM.Good_Percentage_General) * vQtyAfter) / Me.row_BOM.Qty_Ref_Product, _
    '            Me.row_BOM._Decimal, MidpointRounding.AwayFromZero)

    '            If DummyQty1 >= DummyQty Then
    '                Me.row_BOM.qty_raw = DummyQty1
    '            Else
    '                NST = (1 / (10 ^ Me.row_BOM.Decimal_Product)) / 2
    '                DummyQty1 += NST
    '                Me.row_BOM.qty_raw = Math.Round(DummyQty1, Me.row_BOM._Decimal, MidpointRounding.AwayFromZero)
    '            End If
    '        Next

    '        'input ke dalam product
    '        While ij < Me.TBL_BOM.Rows.Count
    '            Me.row_BOM = Me.TBL_BOM.Rows(ij)
    '            Me.TxtPrimaryCodeNob4.Text = Me.row_BOM.primary_product_raw
    '            Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
    '            Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
    '            Me.SetGboItemb4()

    '            If Me.cboBatchNob4.Visible = True Then
    '                Me.cboBatchNob4.Focus()
    '            ElseIf Me.cboLocationb4.Visible = True Then
    '                Me.cboLocationb4.Focus()
    '            Else
    '                Me.TxtQtyperBatchb4.Focus()
    '            End If

    '            'tentukan qty yang dimasukkan
    '            JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
    '            Me.cboBatchNob4.Text)
    '            'perhitungan menggunakan nilai jumlah dgv yang baru
    '            JmlObjLama = Me.objStockBundlingItem.GetJmlB4(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNob4.Text, _
    '            Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, Me.DtpExpiredb4.Value)
    '            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
    '            Me.cboBatchNob4.Text, Me.chkBatchNoControlledb4.Checked, _
    '            Me.chkLocationControlledb4.Checked, Me.chkUseSecondaryUnitb4.Checked, Me.dgvItemb4)
    '            JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text)
    '            QtyEntry = JmlGudang + JmlObjLama - JmlDGV
    '            If QtyEntry > Me.row_BOM.qty_raw Then
    '                QtyEntry = Me.row_BOM.qty_raw
    '            End If
    '            If QtyEntry <> 0 Then
    '                Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(QtyEntry, Me.row_BOM._Decimal)
    '                Me.row_BOM.qty_raw -= QtyEntry
    '                Me.BtnInputb4.PerformClick()
    '                If Me.row_BOM.qty_raw = 0 Then
    '                    Me.TBL_BOM.Rows.Remove(Me.row_BOM)
    '                End If
    '            Else
    '                'tidak ada lagi yang bisa diinput utnuk rm ybs, pindah ke baris selanjutnya
    '                'kasus terjadi misalnya pada qty raw yang diminta lebih banyak dari qty rm yang ada
    '                ij += 1
    '            End If
    '        End While

    '        'berikan peringatan bila qty raw material yang dimasukkan tidak cukup
    '        If Me.TBL_BOM.Rows.Count > 0 Then
    '            For Each Me.row_BOM In Me.TBL_BOM.Rows
    '                If Dummy = "" Then
    '                    Dummy = Me.row_BOM.RM_Code + "(" + Me.ObjCD.DigitGrouping(Me.row_BOM.qty_raw, Me.row_BOM._Decimal) + " " + Me.row_BOM.Unit_Code + ")"
    '                Else
    '                    Dummy = Dummy + ", " + Me.row_BOM.RM_Code + "(" + Me.ObjCD.DigitGrouping(Me.row_BOM.qty_raw, Me.row_BOM._Decimal) + " " + Me.row_BOM.Unit_Code + ")"
    '                End If
    '            Next
    '            MsgBox("Not enough raw material: " + Dummy + ".", MsgBoxStyle.Information, "Information")
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim frm As New FrmCetakTransaksiStockTransfer
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtStockInNo.Text <> "" Then
                frm.primary = Me.TxtPrimaryStockIn.Text
                frm.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockBundling
                frm.UserCode = Me.UserCode
                frm.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    '''FindMaster Proses
    Private Sub BtnFindProses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindProses.Click
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim hj As New Core_Module.FrmMasterProsesCari
        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtCodeProses.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeProses.Text
            End If

            hj.ShowDialog()
            If hj.mPrimary > 0 Then
                Me.TxtPrimaryProses.Text = hj.mPrimary
                Me.TxtCodeProses.Text = ObjProses.GetCode(Me.TxtPrimaryProses.Text)
                Me.TxtDescProses.Text = ObjProses.GetDesc(Me.TxtPrimaryProses.Text)
                Me.TxtCodeWorkUnit.Focus()
            Else
                Me.TxtPrimaryProses.Text = "-1"
                Me.TxtCodeProses.Clear()
                Me.TxtDescProses.Clear()
                Me.TxtCodeProses.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeProses_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeProses.KeyDown
        Dim objProses As New Core_Module.ClsMasterProses
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeProses.Text <> "" Then
                    If objProses.ApakahSudahAda(Me.TxtCodeProses.Text) = True Then
                        Me.TxtPrimaryProses.Text = objProses.GetPrimary(Me.TxtCodeProses.Text)
                        Me.TxtCodeProses.Text = objProses.GetCode(Me.TxtPrimaryProses.Text)
                        Me.TxtDescProses.Text = objProses.GetDesc(Me.TxtPrimaryProses.Text)
                        Me.TxtCodeWorkUnit.Focus()
                    Else
                        Me.BtnFindProses.PerformClick()
                    End If
                Else
                    Me.BtnFindProses.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeProses_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeProses.Leave
        Dim objproses As New Core_Module.ClsMasterProses
        Dim hj As New Core_Module.FrmMasterProsesCari
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCodeProses.Text <> "" Then
                If objproses.ApakahSudahAda(Me.TxtCodeProses.Text) = True Then
                    Me.TxtPrimaryProses.Text = objproses.GetPrimary(Me.TxtCodeProses.Text)
                    Me.TxtDescProses.Text = objproses.GetDesc(Me.TxtPrimaryProses.Text)
                ElseIf objproses.ApakahSudahAda(Me.TxtCodeProses.Text) = False Then
                    If Me.TxtCodeProses.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeProses.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimary > 0 Then
                        Me.TxtPrimaryProses.Text = hj.mPrimary
                        Me.TxtCodeProses.Text = objproses.GetCode(Me.TxtPrimaryProses.Text)
                        Me.TxtDescProses.Text = objproses.GetDesc(Me.TxtPrimaryProses.Text)
                    Else
                        Me.TxtPrimaryProses.Text = "-1"
                        Me.TxtCodeProses.Clear()
                        Me.TxtDescProses.Clear()
                        Me.TxtCodeProses.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryProses.Text = "-1"
                Me.TxtCodeProses.Clear()
                Me.TxtDescProses.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeProses_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeProses.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtCodeProses.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtCodeProses.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeProses_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeProses.LostFocus
        Try
            Me.TxtCodeProses.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '''Find Master WorkUnit 
    Private Sub TxtCodeWorkUnit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeWorkUnit.KeyDown
        Dim ObjWorkUnit As New Core_Module.ClsMasterWorkUnit
        Try
            If e.KeyValue = 13 Then
                If Me.TxtCodeWorkUnit.Text <> "" Then
                    If ObjWorkUnit.ApakahSudahAda(Me.TxtCodeWorkUnit.Text) = True Then
                        Me.TxtPrimaryWorkUnit.Text = ObjWorkUnit.GetPrimary(Me.TxtCodeWorkUnit.Text)
                        Me.TxtCodeWorkUnit.Text = ObjWorkUnit.GetCode(Me.TxtPrimaryWorkUnit.Text)
                        Me.TxtDescWorkUnit.Text = ObjWorkUnit.GetDesc(Me.TxtPrimaryWorkUnit.Text)
                        Me.TxtWorkshiftCode.Focus()
                    Else
                        Me.BtnFindWorkUnit.PerformClick()
                    End If
                Else
                    Me.BtnFindWorkUnit.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeWorkUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeWorkUnit.Leave
        Dim ObjWorkUnit As New Core_Module.ClsMasterWorkUnit
        Dim hj As New Core_Module.FrmMasterWorkUnit_Cari
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCodeWorkUnit.Text <> "" Then
                If ObjWorkUnit.ApakahSudahAda(Me.TxtCodeWorkUnit.Text) = True Then
                    Me.TxtPrimaryWorkUnit.Text = ObjWorkUnit.GetPrimary(Me.TxtCodeWorkUnit.Text)
                    Me.TxtDescWorkUnit.Text = ObjWorkUnit.GetDesc(Me.TxtPrimaryWorkUnit.Text)
                ElseIf ObjWorkUnit.ApakahSudahAda(Me.TxtCodeWorkUnit.Text) = False Then
                    If Me.TxtCodeWorkUnit.Text <> "" Then
                        hj.txtCari.Text = Me.TxtCodeWorkUnit.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryWorkUnit > 0 Then
                        Me.TxtPrimaryWorkUnit.Text = hj.PrimaryWorkUnit
                        Me.TxtCodeWorkUnit.Text = ObjWorkUnit.GetCode(Me.TxtPrimaryWorkUnit.Text)
                        Me.TxtDescWorkUnit.Text = ObjWorkUnit.GetDesc(Me.TxtPrimaryWorkUnit.Text)
                    Else
                        Me.TxtPrimaryWorkUnit.Text = "-1"
                        Me.TxtCodeWorkUnit.Clear()
                        Me.TxtDescWorkUnit.Clear()
                        Me.TxtCodeWorkUnit.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryWorkUnit.Text = "-1"
                Me.TxtCodeWorkUnit.Clear()
                Me.TxtDescWorkUnit.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindWorkUnit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindWorkUnit.Click
        Dim ObjWorkUnit As New Core_Module.ClsMasterWorkUnit
        Dim hj As New Core_Module.FrmMasterWorkUnit_Cari

        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtCodeWorkUnit.Text <> "" Then
                hj.txtCari.Text = Me.TxtCodeWorkUnit.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryWorkUnit > 0 Then
                Me.TxtPrimaryWorkUnit.Text = hj.PrimaryWorkUnit
                Me.TxtCodeWorkUnit.Text = ObjWorkUnit.GetCode(Me.TxtPrimaryWorkUnit.Text)
                Me.TxtDescWorkUnit.Text = ObjWorkUnit.GetDesc(Me.TxtPrimaryWorkUnit.Text)
                Me.TxtWorkshiftCode.Focus()
            Else
                Me.TxtPrimaryWorkUnit.Text = "-1"
                Me.TxtCodeWorkUnit.Clear()
                Me.TxtDescWorkUnit.Clear()
                Me.TxtCodeWorkUnit.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeWorkUnit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeWorkUnit.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtCodeWorkUnit.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtCodeWorkUnit.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeWorkUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeWorkUnit.LostFocus
        Try
            Me.TxtCodeWorkUnit.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkshiftCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtWorkshiftCode.KeyDown
        Dim ObjWorkshift As New Core_Module.ClsMasterWorkShift
        Try
            If e.KeyValue = 13 Then
                If Me.TxtWorkshiftCode.Text <> "" Then
                    If ObjWorkshift.ApakahSudahAda(Me.TxtWorkshiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production) = True Then
                        Me.TxtPrimaryWorkShift.Text = ObjWorkshift.GetPrimary(Me.TxtWorkshiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production)
                        Me.TxtWorkshiftCode.Text = ObjWorkshift.GetCodeByPrimary(Me.TxtPrimaryWorkShift.Text)
                        Me.TxtWorkshiftDesc.Text = ObjWorkshift.GetDescription(Me.TxtPrimaryWorkShift.Text)
                        Me.TxtRemark.Focus()
                    Else
                        Me.btnFindWorkshift.PerformClick()
                    End If
                Else
                    Me.btnFindWorkshift.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkshiftCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkshiftCode.Leave
        Dim ObjWorkshift As New Core_Module.ClsMasterWorkShift
        Dim hj As New Core_Module.FrmMasterWorkShift_Cari(ClsMasterWorkShift.EnumTipeWorkshift.Production)
        hj.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtWorkshiftCode.Text <> "" Then
                If ObjWorkshift.ApakahSudahAda(Me.TxtWorkshiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production) = True Then
                    Me.TxtPrimaryWorkShift.Text = ObjWorkshift.GetPrimary(Me.TxtWorkshiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production)
                    Me.TxtWorkshiftDesc.Text = ObjWorkshift.GetDescription(Me.TxtPrimaryWorkShift.Text)
                ElseIf ObjWorkshift.ApakahSudahAda(Me.TxtWorkshiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production) = False Then
                    If Me.TxtWorkshiftCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtWorkshiftCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryWorkshift > 0 Then
                        Me.TxtPrimaryWorkShift.Text = hj.PrimaryWorkshift
                        Me.TxtWorkshiftCode.Text = ObjWorkshift.GetCodeByPrimary(Me.TxtPrimaryWorkShift.Text)
                        Me.TxtWorkshiftDesc.Text = ObjWorkshift.GetDescription(Me.TxtPrimaryWorkShift.Text)
                    Else
                        Me.TxtPrimaryWorkShift.Text = "-1"
                        Me.TxtWorkshiftCode.Clear()
                        Me.TxtWorkshiftDesc.Clear()
                        Me.TxtWorkshiftCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryWorkShift.Text = "-1"
                Me.TxtWorkshiftCode.Clear()
                Me.TxtWorkshiftDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindWorkshift_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindWorkshift.Click
        Dim ObjWorkshift As New Core_Module.ClsMasterWorkShift
        Dim hj As New Core_Module.FrmMasterWorkShift_Cari(ClsMasterWorkShift.EnumTipeWorkshift.Production)

        Try
            hj.BackColorAll = Me.BackColorAll
            If Me.TxtWorkshiftCode.Text <> "" Then
                hj.txtCari.Text = Me.TxtWorkshiftCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryWorkshift > 0 Then
                Me.TxtPrimaryWorkShift.Text = hj.PrimaryWorkshift
                Me.TxtWorkshiftCode.Text = ObjWorkshift.GetCodeByPrimary(Me.TxtPrimaryWorkShift.Text)
                Me.TxtWorkshiftDesc.Text = ObjWorkshift.GetDescription(Me.TxtPrimaryWorkShift.Text)
                Me.TxtRemark.Focus()
            Else
                Me.TxtPrimaryWorkShift.Text = "-1"
                Me.TxtWorkshiftCode.Clear()
                Me.TxtWorkshiftDesc.Clear()
                Me.TxtWorkshiftCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtWorkshiftCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkshiftCode.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Inventory Then
                Me.TxtWorkshiftCode.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtWorkshiftCode.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkshiftCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkshiftCode.LostFocus
        Try
            Me.TxtWorkshiftCode.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Dim ij As New FrmCariTransaksiStockBundling

        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.ObjStockBundling.CariDetail(vPrimary, _
                Me.dgvItemb4, Me.dgvItemAfter, Me) = True Then
                    Me.SalinKeField()
                    Me.TabControl1.SelectedIndex = 0
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Public Sub EditTanggal(ByVal vTglTujuan As Date)
        'Dim hasil As MsgBoxResult

        Try
            If Me.TxtStockInNo.Text <> "" Then
                If Me.ObjStockBundling.ApakahBolehEdit(Me.UserCode) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItemb4.Focus()

                    'ganti tanggal
                    Me.dtpDate.Value = vTglTujuan

                    'If Me.ApakahDataBenar() = True Then
                    Me.KosongkanItemAfter()
                    Me.KosongkanItemB4()

                    Me.SalinKeObjek()
                    If baru = True Then
                        ''konfirmasi disimpan atau tidak
                        'hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        'If hasil = MsgBoxResult.Yes Then
                        '    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        '    If Me.ObjStockBundling.SimpanBaru(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) Then
                        '        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        '        Me.SalinKeField()
                        '        Me.NonAktifkanDetail()
                        '        Me.btnFind.Focus()
                        '    End If
                        'End If
                    Else
                        'konfirmasi disimpan atau tidak
                        'hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        'If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjStockBundling.SimpanPerubahan(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) Then
                            'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                        'End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetQtyAvailable()
        Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV, JmlObjLamaAft As Decimal
        Try
            JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockBundlingItem.GetJmlB4(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNob4.Text, _
            Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, Me.DtpExpiredb4.Value)
            JmlObjLamaAft = Me.objStockBundlingItem.GetJmlAfter(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNob4.Text, _
            Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, Me.DtpExpiredb4.Value)
            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
            Me.cboBatchNob4.Text, Me.chkBatchNoControlledb4.Checked, _
            Me.chkLocationControlledb4.Checked, Me.chkUseSecondaryUnitb4.Checked, Me.dgvItemb4)

            If JmlGudang >= JmlObjLamaAft Then
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text) - JmlObjLamaAft
            Else
                JmlSetelahTambah = JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text)
            End If

            If JmlGudang >= JmlObjLamaAft Then
                Me.TxtQty_OnHandB4.Text= Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV - JmlObjLamaAft), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) 
            Else
                Me.TxtQty_OnHandB4.Text= Me.ObjCD.DigitGrouping(CStr(JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) 
            End If
            Me.lblUnits2B4.Text = Me.lblUnitsb4.Text

            'Me.TxtQty_OnHandB4.Text = Me.ObjCD.DigitGrouping(Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text) - JmlDGV, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_StockBundling_Create(ByVal vPrimaryTransaksi As Double)
        Try
            Me.baru = False
            Me.ObjStockBundling.ResetObject(Me.dgvItemb4, Me.dgvItemAfter)

            Me.ObjStockBundling.CariDetail(vPrimaryTransaksi, Me.dgvItemb4, Me.dgvItemAfter, Me)
            Me.UserCode = Me.ObjStockBundling.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_StockBundling_Save() As Double
        Try
            Me.baru = True
            Me.ObjStockBundling.SimpanBaru1(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me, False, False, False, True)
            Return Me.ObjStockBundling.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub llbAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAdvance.LinkClicked
        Try
            Me.cms1.Show(Me.llbAdvance, New Point((Me.cms1.Width * -1) - 5 + (Me.llbAdvance.Width / 2), Me.llbAdvance.Height + 5))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsm_AddByRefDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsm_AddByRefDoc.Click
        Dim Obj As New ClsTransaksiBarcode
        Dim ij As New FrmTransaksiStockAdvanceAdd
        Dim TBLData As New Core_Module.DseBarcodeInput.DataTable1DataTable
        Dim InputType As CustomSMESS.ClsCustomBarcode.EnumInputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add
        Dim Index As Integer = -1
        Dim rowData As Core_Module.DseBarcodeInput.DataTable1Row

        Try
            'ij.BackColorAll = Me.BackColorAll
            'ij.UserCode = Me.UserCode
            'If Me.TabControl1.SelectedIndex = 0 Then
            '    ij.ParentForm1 = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_Before
            'Else
            '    ij.ParentForm1 = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_After
            'End If
            'ij.ShowDialog()
            If Me.TabControl1.SelectedIndex = 0 Then
                Obj.GetDataInput(TBLData, InputType, False, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_Before, Me.dgvItemb4, Me.TxtPrimaryStockIn.Text, Me.UserCode)
            Else
                Obj.GetDataInput(TBLData, InputType, False, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Bundling_Item_After, Me.dgvItemAfter, Me.TxtPrimaryStockIn.Text, Me.UserCode)
            End If


            If InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add Then

            ElseIf InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Replace Then
                If TBLData.Rows.Count > 0 Then
                    If Me.TabControl1.SelectedIndex = 0 Then
                        Me.dgvItemb4.Rows.Clear()
                    Else
                        Me.dgvItemAfter.Rows.Clear()
                    End If
                End If
            End If

            If Me.TabControl1.SelectedIndex = 0 Then
                For Each rowData In TBLData
                    With rowData
                        Index = Me.GetIndex2(.Primary_Product, .Batch, -1, -1, Me.dgvItemb4, .Primary_Location, .Location_Controlled)
                        'cari data yang sama persis
                        If Index <> -1 Then
                            'tambahkan nilai qty
                            Dim row As DataGridViewRow = Me.dgvItemb4.Rows(Index)
                            row.Cells(3).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(3).Value) + CDec(.Qty), Me.ObjProduct.GetAngkaDecimal(.Primary_Product))
                        Else
                            'tambahkan baris baru
                            Dim ExpiredDateStr As String
                            If .Expired_Date = #1/1/1950# Then
                                ExpiredDateStr = "-"
                            Else
                                ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                            End If
                            Me.dgvItemb4.Rows.Add("", .Code, .Description, Me.ObjCD.DigitGrouping(.Qty, ._Decimal), _
                      .Unit, "-", "-", .Batch, _
                      ExpiredDateStr, .Location, .Warehouse, .Remarks, _
                      .Primary_Product, .Primary_Location, .Primary_Unit, _
                      -1)
                        End If
                    End With
                Next
                Me.ObjCD.SetNumber(Me.dgvItemb4)
                'format dgv
                Me.FormatDGV(dgvItemb4)

                'fokuskan ke baris terakhir
                If Me.dgvItemb4.Rows.Count > 0 Then
                    While Me.dgvItemb4.Rows(Me.dgvItemb4.Rows.Count - 1).Displayed = False _
                             Or Me.dgvItemb4.DisplayedRowCount(True) <> Me.dgvItemb4.DisplayedRowCount(False)
                        Me.dgvItemb4.FirstDisplayedScrollingRowIndex = Me.dgvItemb4.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            Else

                For Each rowData In TBLData
                    With rowData
                        Index = Me.GetIndex2(.Primary_Product, .Batch, -1, -1, Me.dgvItemAfter, .Primary_Location, .Location_Controlled)
                        'cari data yang sama persis
                        If Index <> -1 Then
                            'tambahkan nilai qty
                            Dim row As DataGridViewRow = Me.dgvItemAfter.Rows(Index)
                            row.Cells(3).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(3).Value) + CDec(.Qty), ._Decimal)
                        Else
                            'tambahkan baris baru
                            Dim ExpiredDateStr As String
                            If .Expired_Date = #1/1/1950# Then
                                ExpiredDateStr = "-"
                            Else
                                ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                            End If
                            Me.dgvItemAfter.Rows.Add("", .Code, .Description, Me.ObjCD.DigitGrouping(.Qty, ._Decimal), _
                      .Unit, "-", "-", .Batch, _
                      ExpiredDateStr, .Location, .Warehouse, .Remarks, _
                      .Primary_Product, .Primary_Location, .Primary_Unit, _
                      -1)
                        End If
                    End With
                Next
                Me.ObjCD.SetNumber(Me.dgvItemAfter)
                'format dgv
                Me.FormatDGV(dgvItemAfter)

                'fokuskan ke baris terakhir
                If Me.dgvItemAfter.Rows.Count > 0 Then
                    While Me.dgvItemAfter.Rows(Me.dgvItemAfter.Rows.Count - 1).Displayed = False _
                             Or Me.dgvItemAfter.DisplayedRowCount(True) <> Me.dgvItemAfter.DisplayedRowCount(False)
                        Me.dgvItemAfter.FirstDisplayedScrollingRowIndex = Me.dgvItemAfter.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsm_AddByBarcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsm_AddByBarcode.Click
        Dim Obj As New CustomSMESS.ClsCustomBarcode
        Dim TBLData As New Core_Module.DseBarcodeInput.DataTable1DataTable
        Dim InputType As CustomSMESS.ClsCustomBarcode.EnumInputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add
        Dim rowData As Core_Module.DseBarcodeInput.DataTable1Row
        Dim Index As Integer = -1
        Dim AdaService As Boolean = False

        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Obj.GetDataInput(TBLData, InputType, False, CustomSMESS.ClsCustomBarcode.EnumTransaksiAcuan.BundlingBefore)
            Else
                Obj.GetDataInput(TBLData, InputType, False, CustomSMESS.ClsCustomBarcode.EnumTransaksiAcuan.BundlingAfter)
            End If

            If InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add Then

            ElseIf InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Replace Then
                If TBLData.Rows.Count > 0 Then
                    If Me.TabControl1.SelectedIndex = 0 Then
                        Me.dgvItemb4.Rows.Clear()
                    Else
                        Me.dgvItemAfter.Rows.Clear()
                    End If
                End If
            End If

            If Me.TabControl1.SelectedIndex = 0 Then
                For Each rowData In TBLData
                    With rowData
                        If rowData.Is_Service = 1 Then
                            AdaService = True
                        Else
                            Index = Me.GetIndex2(.Primary_Product, .Batch, -1, -1, Me.dgvItemb4, .Primary_Location, .Location_Controlled)
                            'cari data yang sama persis
                            If Index <> -1 Then
                                'tambahkan nilai qty
                                Dim row As DataGridViewRow = Me.dgvItemb4.Rows(Index)
                                row.Cells(3).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(3).Value) + CDec(.Qty), ._Decimal)
                            Else
                                'tambahkan baris baru
                                Dim ExpiredDateStr As String
                                If .Expired_Date = #1/1/1950# Then
                                    ExpiredDateStr = "-"
                                Else
                                    ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                                End If
                                Me.dgvItemb4.Rows.Add("", .Code, .Description, Me.ObjCD.DigitGrouping(.Qty, ._Decimal), _
                          .Unit, "-", "-", .Batch, _
                          ExpiredDateStr, .Location, .Warehouse, .Remarks, _
                          .Primary_Product, .Primary_Location, .Primary_Unit, _
                          -1)
                            End If
                        End If
                    End With
                Next
                Me.ObjCD.SetNumber(Me.dgvItemb4)
                'format dgv
                Me.FormatDGV(dgvItemb4)

                'fokuskan ke baris terakhir
                If Me.dgvItemb4.Rows.Count > 0 Then
                    While Me.dgvItemb4.Rows(Me.dgvItemb4.Rows.Count - 1).Displayed = False _
                             Or Me.dgvItemb4.DisplayedRowCount(True) <> Me.dgvItemb4.DisplayedRowCount(False)
                        Me.dgvItemb4.FirstDisplayedScrollingRowIndex = Me.dgvItemb4.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            Else

                For Each rowData In TBLData
                    With rowData
                        If rowData.Is_Service = 1 Then
                            AdaService = True
                        Else
                            Index = Me.GetIndex2(.Primary_Product, .Batch, -1, -1, Me.dgvItemAfter, .Primary_Location, .Location_Controlled)
                            'cari data yang sama persis
                            If Index <> -1 Then
                                'tambahkan nilai qty
                                Dim row As DataGridViewRow = Me.dgvItemAfter.Rows(Index)
                                row.Cells(3).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(3).Value) + CDec(.Qty), Me.ObjProduct.GetAngkaDecimal(.Primary_Product))
                            Else
                                'tambahkan baris baru
                                Dim ExpiredDateStr As String
                                If .Expired_Date = #1/1/1950# Then
                                    ExpiredDateStr = "-"
                                Else
                                    ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                                End If
                                Me.dgvItemAfter.Rows.Add("", .Code, .Description, .Qty, _
                          .Unit, "-", "-", .Batch, _
                          ExpiredDateStr, .Location, .Warehouse, .Remarks, _
                          .Primary_Product, .Primary_Location, .Primary_Unit, _
                          -1)
                            End If
                        End If
                    End With
                Next
                Me.ObjCD.SetNumber(Me.dgvItemAfter)
                'format dgv
                Me.FormatDGV(dgvItemAfter)

                'fokuskan ke baris terakhir
                If Me.dgvItemAfter.Rows.Count > 0 Then
                    While Me.dgvItemAfter.Rows(Me.dgvItemAfter.Rows.Count - 1).Displayed = False _
                             Or Me.dgvItemAfter.DisplayedRowCount(True) <> Me.dgvItemAfter.DisplayedRowCount(False)
                        Me.dgvItemAfter.FirstDisplayedScrollingRowIndex = Me.dgvItemAfter.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            End If
            If AdaService = True Then
                MsgBox("All service is excluded.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cms1_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cms1.Opening
        Try 
            Me.tsm_AddByRefDoc.Enabled = Me.QuickAddByRefDoc
            Me.tsm_AddByBarcode.Enabled = Me.QuickAddByBarcode 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex2(ByVal vPrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Decimal, ByVal PrimaryPdO As Double, ByVal vDGV As DataGridView, ByVal vPrimaryLocation As Double, ByVal vLocationControl As Integer) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In vDGV.Rows
                If vLocationControl = 1 Then
                    If IsNumeric(row.Cells(5).Value) = False Then
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(7).Value = BatchNo And row.Cells(13).Value = vPrimaryLocation Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(7).Value = BatchNo And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(5).Value) = False Then
                        If row.Cells(12).Value = vPrimaryProduct And _
                        row.Cells(7).Value = BatchNo Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(12).Value = vPrimaryProduct And _
                        row.Cells(7).Value = BatchNo Then
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

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiStockBundling_AddByCustomProduct

        Try
            If Me.TxtRemark.Enabled = True Then
                'If Me.TxtPrimaryPihakAc.Text = -1 And Me.TxtPrimaryPihakAcuan.Text <> "" Then
                '    MsgBox("To start entering item data, you must specify supplier data first.", MsgBoxStyle.Information, "Information")
                '    Me.TxtPihakAcuan.Focus()
                '    Exit Sub
                'End If
            End If

            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TblItemBefCustom = Me.TblItemBefCustom.Copy
            i.TblItemAftCustom = Me.TblItemAftCustom.Copy
            i.PrimaryTransaksi = Me.TxtPrimaryStockIn.Text
            i.TglTransaksi = Me.dtpDate.Value
            i.Baru = Me.baru
            'i.PrimaryPackingList = Me.txtPrimaryDO.Text
            If Me.TxtRemark.Enabled = False Then
                i.CanEditForm = False
            End If

            If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblItemBefCustom.Clear()
                Me.TblItemAftCustom.Clear()
                Me.TblItemBefCustom = i.TblItemBefCustom
                Me.TblItemAftCustom = i.TblItemAftCustom
                Me.IsikanDgvFromTblDummy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini dibuat untuk admiral
    Private Function IsikanDgvFromTblDummy() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowBef As CustomSMESS.DseTransaksiStockBundling.tbltrans_custom_stock_bundling_item_beforeRow
        Dim rowAft As CustomSMESS.DseTransaksiStockBundling.tbltrans_custom_stock_bundling_item_afterRow
        Dim ObjP As New Core_Module.clsMasterProduct
        Dim rd As DataGridViewRow
        Dim DummyPO As String = ""
        Dim PrimaryLocationLama As Double = -1
        Dim IdxRows As Integer = -1

        Try
            'ITEM BEFORE
            Me.dgvItemb4.Rows.Clear()
            Me.dgvItemb4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowBef In TblItemBefCustom.Rows
                With rowBef
                    IdxRows = Me.GetIndex(.Primary_Product, .Primary_Storage_Area, .Batch_No, _
                    -1, -1, .Expiry_Date, Me.dgvItemb4)

                    'cari data yang sama persis
                    If IdxRows <> -1 Then
                        'tambahkan nilai qty
                        rd = Me.dgvItemb4.Rows(IdxRows)
                        If .is_kain = True Then
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQty.Index).Value) + (.Qty_Roll * .Qty_At), ._Decimal)
                        Else
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQty.Index).Value) + .Qty, ._Decimal)
                        End If
                    Else
                        'tambahkan baris baru
                        rd = Me.dgvItemb4.Rows(Me.dgvItemb4.Rows.Add())
                        rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description

                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        If .is_kain = True Then
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                            rd.Cells(Me.ColSecondaryQty.Index).Value = "-1" 'Me.ObjCD.DigitGrouping(.Qty_At, ._Decimal)
                            rd.Cells(Me.ColSecondaryUnit.Index).Value = "" '.Unit_Code
                        Else
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                            rd.Cells(Me.ColSecondaryQty.Index).Value = "-1"
                            rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        End If
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code

                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        If .Expiry_Date = #1/1/1950# Then
                            rd.Cells(Me.ColExpiredDate.Index).Value = "" '#1/1/1950#
                        Else
                            rd.Cells(Me.ColExpiredDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If

                        rd.Cells(Me.ColLocation.Index).Value = .location_code
                        rd.Cells(Me.ColWareHouse.Index).Value = .warehouse_code
                        rd.Cells(Me.ColRemark.Index).Value = .Remarks
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = .Primary_Secondary_Units
                    End If
                End With
            Next
            Me.dgvItemb4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'ITEM AFTER
            Me.dgvItemAfter.Rows.Clear()
            Me.dgvItemAfter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowAft In TblItemAftCustom.Rows
                With rowAft
                    IdxRows = Me.GetIndex(.Primary_Product, .Primary_Storage_Area, .Batch_No, _
                    .Secondary_Qty, -1, .Expiry_Date, Me.dgvItemAfter)

                    'cari data yang sama persis
                    If IdxRows <> -1 Then
                        'tambahkan nilai qty
                        rd = Me.dgvItemAfter.Rows(IdxRows)
                        If .is_kain = True Then
                            rd.Cells(Me.ColQtyAfter.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQtyAfter.Index).Value) + (.Qty_Roll * .Qty_At), ._Decimal)
                        Else
                            rd.Cells(Me.ColQtyAfter.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQtyAfter.Index).Value) + .Qty, ._Decimal)
                        End If
                        rd.Cells(Me.ColAfter_CostFactor.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColAfter_CostFactor.Index).Value) + .Faktor_Pemberat, 2)
                    Else
                        'tambahkan baris baru
                        rd = Me.dgvItemAfter.Rows(Me.dgvItemAfter.Rows.Add())
                        rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColCodeNoAfter.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description

                        rd.Cells(Me.ColPrimaryCodeNoAfter.Index).Value = .Primary_Product
                        If .is_kain = True Then
                            rd.Cells(Me.ColQtyAfter.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                            rd.Cells(Me.ColSecondaryQty.Index).Value = "-1" 'Me.ObjCD.DigitGrouping(.Qty_At, ._Decimal)
                            rd.Cells(Me.ColSecondaryUnit.Index).Value = "" '.Unit_Code
                        Else
                            rd.Cells(Me.ColQtyAfter.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                            rd.Cells(Me.ColSecondaryQty.Index).Value = ""
                            rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        End If
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code

                        rd.Cells(Me.ColBatchNoAfter.Index).Value = .Batch_No
                        If .Expiry_Date = #1/1/1950# Then
                            rd.Cells(Me.ColExpiryAfter.Index).Value = "" '#1/1/1950#
                        Else
                            rd.Cells(Me.ColExpiryAfter.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If

                        rd.Cells(Me.ColLocation.Index).Value = .location_code
                        rd.Cells(Me.ColWareHouse.Index).Value = .warehouse_code
                        rd.Cells(Me.ColRemark.Index).Value = .Remarks
                        rd.Cells(Me.ColPrimaryCodeNoAfter.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocationAfter.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnitAfter.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = .Primary_Secondary_Units
                        rd.Cells(Me.ColAfter_CostFactor.Index).Value = Me.ObjCD.DigitGrouping(.Faktor_Pemberat, 2)
                    End If
                End With
            Next
            Me.dgvItemAfter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.FormatDGV(Me.dgvItemb4)
            Me.FormatDGV(Me.dgvItemAfter)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class