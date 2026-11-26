Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiDO3

    'Update SMESS 2026
    Public vQuickPrint As Boolean = False
    Public vFromQuick As Boolean = False
    Public vStopCancellation As Boolean = False
    Public vDirectlySave As Boolean = False
    Public vDirectlyPrint As Boolean = False
    Public vShowDialog As Boolean = True
    Public vMinDate As Date = #1/1/1950#

    Public UserCode As String = ""
    Public BackColorAll As Color = Me.BackColorAll
    Public FormNo As Integer = 0
    Public vPrimaryDO As Double = -1
    Public vDONo As String = ""

    Dim baru As Boolean
    Dim objDO As New ClsTransaksiDO
    Public ObjCD As New ClsCekData 'Update SMESS 2026
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjDOItem As New ClsTransaksiDOItem
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSOItem As New ClsTransaksiSalesOrderItem
    Dim ObjBookingInv As New ClsTransaksiBookingInvItem
    Dim ObjBookingGIT As New ClsTransaksiBookingGITitem
    Dim ObjCL As New ClsTransaksiPurchaseClosing
    Dim ObjInvoiceItem As New ClsTransaksiSalesInvoiceItem
    Dim ObjLicense As New ClsLicense
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim IsInventoryInstalled As Boolean = False
    Dim ObjOtorisasi As New Core_Module.ClsMasterTipeOtorisasi
    Dim ObjOtorisasi_DOExceedCreditLimit As New Core_Module.ClsMasterTipeOtorisasi
    Dim ObjOtorisasi_Overdue_Invoice As New Core_Module.ClsMasterTipeOtorisasi
    Dim ApakahBolehDOTanpaStock As Boolean = False
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable

    Dim TBL As New DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable
    Dim TBL_CustPO As New DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable
    Dim TBL2 As New DseTransaksiDO.DataTable2DataTable
    Dim TBL3 As New DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable
    Dim TBL4 As New DseTransaksiDO.DataTable2DataTable
    Dim TBL5 As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL6 As New DseTransaksiDO.tbltrans_do_item1DataTable
    Dim TBL7 As New dseMasterProduct.tblmaster_productqty3DataTable
    Dim TBL8 As New DseTransaksiDO.tbltrans_do_item2DataTable

    Dim PersenPajakAcuan, PrimaryCurrencySO As Decimal
    Dim JenisFakturPajakAcuan As ClsTransaksiSalesOrder.EnumJenisFakturPajak = ClsTransaksiSalesOrder.EnumJenisFakturPajak.FakturPajakSederhana
    Dim SamaSO As Boolean = False
    Dim SamaCodeNo As Boolean = False
    Dim buffer As String = ""
    Dim bufferProduct As String = ""
    Dim BufferSO As String = ""
    Dim SelectedCBOSOChange As Boolean = True
    Dim SelectedCBOPOChange As Boolean = True
    Dim ShowDeliveryDate As Boolean = False

    Dim TBLCustomItem As New CustomSMESS.DseTransaksiDO_AddByCustomProduct.tbltrans_do_itemDataTable
    Public TBLCustomItem_Loc As New CustomSMESS.DseTransaksiDO_Location.tbltrans_custom_doDataTable
    Dim objCustomDO As New CustomSMESS.ClsTransaksiDO 'ahdar

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryDO.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryDO.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryDO.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryDO.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryDO.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryDO.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, False)
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
            'Me.llbAddress.Enabled = True
            Me.llbDetailStock.Enabled = True
            Me.chkKwitansi.Enabled = True
            Me.chkFakturPajak.Enabled = True
            If Me.baru = True Then
                Me.llbAddBooking.Enabled = True
                'Me.cboSalesOrderNo.Width = Me.txtDesc.Width - Me.llbAddBooking.Width - 5
            Else
                Me.cboSalesOrderNo.Width = Me.txtDesc.Width
            End If

            If Me.dgvItem.RowCount > 0 Then
                Me.txtCustomerCode.Enabled = False
                Me.btnFindCustomer.Enabled = False
            End If

            If Me.txtPrimaryCustomer.Text <> "-1" And Me.txtPrimaryCustomer.Text <> "" Then
                Me.SetGboItem()
            End If

            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                Me.TxtDONo.ReadOnly = True
                Me.TxtDONo.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtDONo.TabStop = False
                Me.dtpDate.Focus()
            Else
                Me.TxtDONo.ReadOnly = False
                Me.TxtDONo.BackColor = Me.TxtPrimaryDO.BackColor
                Me.TxtDONo.TabStop = True
                Me.TxtDONo.Focus()
            End If
            Me.llbAddByLocation.Enabled = True
            Me.ChkIgnoreTolerance.Enabled = True

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_EditShipTo") = True Then
                Me.CboAddressShipTo.Enabled = True
            Else
                Me.CboAddressShipTo.Enabled = False
            End If
            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_AddItembyCustom") = True Then
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
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            'Me.llbAddress.Enabled = False
            Me.llbDetailStock.Enabled = False
            Me.chkKwitansi.Enabled = False
            Me.chkFakturPajak.Enabled = False
            Me.llbAddBooking.Enabled = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)

            If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiDO_New") = 0 Then
                Me.btnNew.Enabled = False
            End If
            If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiDO_Edit") = 0 Then
                Me.btnEdit.Enabled = False
            End If
            If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiDO_Hapus") = 0 Then
                Me.btnDelete.Enabled = False
            End If
            If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiDO_Cetak") = 0 Then
                Me.btnPrint.Enabled = False
            End If
            Me.llbAddByLocation.Enabled = False
            Me.ChkIgnoreTolerance.Enabled = False
            Me.CboAddressShipTo.Enabled = False
            If vFromQuick Then
                Me.btnStopProcess.Enabled = True
                Me.btnStopProcess.Visible = True
                Me.prgProgress.Visible = True
            End If
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
            ModCekData.KosongkanControl(Me.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.Panel2.Controls, True, True, False)

            Me.txtPrimarySO.Text = "-1"
            Me.txtPrimaryCodeNo.Text = "-1"
            Me.TxtPrimaryDO.Text = "-1"
            Me.txtPrimaryCustomer.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.cboCustomerAddress.DataSource = Nothing
            Me.cboCustomerAddress.Text = ""
            Me.chkKwitansi.Checked = False
            Me.chkFakturPajak.Checked = False
            Me.dgvItem.Rows.Clear()
            Me.cboCustomerAddress.DataSource = Nothing
            Me.cboCustomerAddress.Items.Clear()
            Me.CboAddressShipTo.Items.Clear()
            Me.CboAddressCustomerShipTo.Items.Clear()
            Me.TxtPrimaryCustomerShipTo.Text = "-1"

            Me.TBLCustomItem_Loc.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.txtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.txtPrimarySO.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyperBatchMulti.Text = "0"
            Me.TxtQtyperBatchDefault.Text = "0"
            Me.txtSecondaryQty.Text = "-1"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnitsMulti.Left = Me.lblSecondaryUnits.Left
            'Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
            Me.lblSecQty.Visible = False
            Me.lblExpiryDate.Visible = False
            Me.DtpExpired.Visible = True
            Me.DtpExpired.Value = Today
            Me.lblBatchNo.Visible = False
            Me.cboBatchNo.Visible = True
            Me.cboBatchNo.TabStop = True
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Items.Clear()
            Me.cboSecondaryQty.Visible = False
            Me.cboSalesOrderNo.DataSource = Nothing
            Me.cboCodeNo.DataSource = Nothing

            Me.txtPendingDO.Text = "0"
            Me.txtAvailableMulti.Text = "0"
            Me.txtAvailable.Text = "0"
            Me.LlbQty.Text = "Qty"
            Me.TxtQtyperBatchMulti.Width = Me.txtAvailableMulti.Right - Me.TxtQtyperBatchMulti.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar(Optional ByVal vCekOnhand As Boolean = True) As Boolean
        Dim JmlAcuan As Decimal = 0

        Try
            'data item salah bila:
            '1. Bila product code tidak diisi
            If Me.cboCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                MsgBox("Please specify product code", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '2.b. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlled.Checked = True And Me.IsInventoryInstalled = False Then
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

            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlled.Checked = True And Me.DtpExpired.Visible = True Then
                If Me.DtpExpired.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
            End If

            '7. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnit.Checked = True And Me.IsInventoryInstalled = False Then
                If Me.ObjCD.CekBilanganPositif(Me.cboSecondaryQty.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.cboSecondaryQty.Focus()
                    Return False
                End If
            End If

            '8a. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlled.Checked = True And Me.chkExpiryDateControlled.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
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
                    If row.Cells(Me.ColBatchNo.Index).Value = Me.cboBatchNo.Text And Me.cboBatchNo.Text <> "" And row.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                    Me.txtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColExpiryDate.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
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

            '9. Bila setelah diinput jumlah sales order balance menjadi minus
            Dim JmlSOBBaru As Decimal
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlSOBBaru = CDec(Me.txtQtyOrderBalance.Text) - CDec(Me.TxtQtyperBatchMulti.Text) + CDec(Me.TxtToleranceUp.Text)
            If JmlSOBBaru < 0 Then
                JmlAcuan = CDec(Me.txtQtyOrderBalance.Text) + CDec(Me.TxtToleranceUp.Text)
                If Me.ObjProduct.ApakahGoods(Me.txtPrimaryCodeNo.Text) = True Then
                    If Me.ApakahBolehDOTanpaStock = False Then
                        If CDec(Me.txtAvailableMulti.Text) < JmlAcuan Then
                            JmlAcuan = CDec(Me.txtAvailableMulti.Text)
                        End If
                    End If
                End If

                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(JmlAcuan) + " " + _
                Me.lblUnitsMulti.Text + ".", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If

            '9.b. Cek stock pada tanggal yang bersangkutan
            'msgbox("belum"), harus dipertimbangkan kembali

            '9.c Jika inventory diinstall cek kembali berkaitan dengan stock digudang,
            '    jangan sampe malahan bikin minus
            If vCekOnhand = True Then
                If Me.ObjProduct.ApakahGoods(Me.txtPrimaryCodeNo.Text) = True Then
                    If Me.ApakahBolehDOTanpaStock = False Then
                        If CDec(Me.txtAvailableMulti.Text) < CDec(Me.TxtQtyperBatchMulti.Text) Then
                            JmlAcuan = CDec(Me.txtAvailableMulti.Text)
                            If JmlAcuan > CDec(Me.txtQtyOrderBalance.Text) + CDec(Me.TxtToleranceUp.Text) Then
                                JmlAcuan = CDec(Me.txtQtyOrderBalance.Text) + CDec(Me.TxtToleranceUp.Text)
                            End If
                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(JmlAcuan, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) + " " + _
                            Me.lblUnitsMulti.Text + ".", MsgBoxStyle.Information, "Information")
                            Me.TxtQtyperBatchMulti.Focus()
                            Return False
                        End If

                        If CDec(Me.txtAvailable.Text) < CDec(Me.TxtQtyperBatchDefault.Text) Then
                            JmlAcuan = CDec(Me.txtAvailable.Text)
                            MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(JmlAcuan, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)) + " " + _
                            Me.lblUnitsDefault.Text + ".", MsgBoxStyle.Information, "Information")
                            Me.TxtQtyperBatchMulti.Focus()
                            Return False
                        End If
                    End If
                End If
            End If


            '10. Dalam 1 do bisa ada beberapa SO dengan syarat:
            If Me.dgvItem.RowCount > 0 Then
                Dim mTaxNew, mprimaryCurrency As Decimal
                Dim mJenisFaktur As ClsTransaksiSalesOrder.EnumJenisFakturPajak
                Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.txtPrimarySO.Text, _
                mprimaryCurrency, mJenisFaktur, mTaxNew)
                '   a. % pajak yang digunakan sama
                If mTaxNew <> Me.PersenPajakAcuan Then
                    MsgBox("This sales order cannot be in one delivery order with the existing items. " + _
                    "For several sales order item written in one delivery order their currency, tax, and tax invoice must be identical", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '   b. jenis faktur pajak yang digunakan sama
                If mJenisFaktur <> Me.JenisFakturPajakAcuan Then
                    MsgBox("This sales order cannot be in one delivery order with the existing items. " + _
                    "For several sales order item written in one delivery order their currency, tax, and tax invoice must be identical", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '   c. primary currency sama
                If mprimaryCurrency <> Me.PrimaryCurrencySO Then
                    MsgBox("This sales order cannot be in one delivery order with the existing items. " + _
                    "For several sales order item written in one delivery order their currency, tax, and tax invoice must be identical", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            ''tidak bisa input yang ship to nya berbeda
            If Me.dgvItem.RowCount > 0 Then
                Dim DtaDataSO As New DseTransaksiSOTableAdapters.tbltrans_so3TableAdapter
                Dim TblDataSOExist As New DseTransaksiSO.tbltrans_so3DataTable
                Dim TblDataSONew As New DseTransaksiSO.tbltrans_so3DataTable
                Dim RowDataSOExist As DseTransaksiSO.tbltrans_so3Row
                Dim RowDataSONew As DseTransaksiSO.tbltrans_so3Row
                DtaDataSO.Connection.ConnectionString = Core_Module.GetConstring
                If DtaDataSO.Fill(TblDataSOExist, Me.dgvItem.Item(Me.ColPrimarySO.Index, 0).Value) > 0 And _
                        DtaDataSO.Fill(TblDataSONew, Me.txtPrimarySO.Text) > 0 Then
                    RowDataSOExist = TblDataSOExist.Rows(0)
                    RowDataSONew = TblDataSONew.Rows(0)
                    If RowDataSOExist.identifier_address_customer_ship_to = "-1" Then
                        RowDataSOExist.identifier_address_customer_ship_to = ""
                    End If
                    If RowDataSONew.identifier_address_customer_ship_to = "-1" Then
                        RowDataSONew.identifier_address_customer_ship_to = ""
                    End If
                    If RowDataSOExist.identifier_address_ship_to <> RowDataSONew.identifier_address_ship_to Or _
                             RowDataSOExist.identifier_address_customer_ship_to <> RowDataSONew.identifier_address_customer_ship_to Then
                        MsgBox("You have to input SO data that has the same shipping address.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'ahdar
            If objCustomDO.ApakahDataItemBenar_Custom(Me.txtPrimaryCodeNo.Text, True) = False Then
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Decimal, ByVal vBatchNo As String, _
    ByVal vSecQty As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColBatchNo.Index).Value = vBatchNo _
                    And row.Cells(Me.ColSecQty.Index).Value = vSecQty Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColBatchNo.Index).Value = vBatchNo Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Decimal, ByVal vBatchNo As String, _
        ByVal vSecQty As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColBatchNo.Index).Value = vBatchNo _
                    And row.Cells(Me.ColSecQty.Index).Value = vSecQty Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColBatchNo.Index).Value = vBatchNo Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Decimal, ByVal vBatchNo As String, _
    ByVal vSecQty As Decimal, ByVal vPrimarySO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColBatchNo.Index).Value = vBatchNo _
                    And row.Cells(Me.ColSecQty.Index).Value = vSecQty And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColBatchNo.Index).Value = vBatchNo And _
                    row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Decimal, ByVal vBatchNo As String, _
    ByVal vSecQty As Decimal, ByVal vPrimarySO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColBatchNo.Index).Value = vBatchNo _
                    And row.Cells(Me.ColSecQty.Index).Value = vSecQty And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And _
                    row.Cells(Me.ColBatchNo.Index).Value = vBatchNo And _
                    row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml_PrimaryLocation_Minus(ByVal vPrimaryCodeNo As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimaryLocation.Index).Value = -1 Then
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml_Default(ByVal vPrimaryCodeNo As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimarySO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimarySO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV2_Default(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimarySO As Decimal, ByVal vBatchNo As String) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                row.Cells(Me.ColBatchNo.Index).Value = vBatchNo Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV2_Multi(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimarySO As Decimal, ByVal vBatchNo As String) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                row.Cells(Me.ColBatchNo.Index).Value = vBatchNo Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ResetObjOtorisasi_OverdueInvoice()
        Try
            Me.ObjOtorisasi_Overdue_Invoice.Primary = -1
            Me.ObjOtorisasi_Overdue_Invoice.PrimaryApprovedBy = -1
            Me.ObjOtorisasi_Overdue_Invoice.PrimaryDokumenAcuan = -1
            Me.ObjOtorisasi_Overdue_Invoice.PrimaryRequestBy = -1
            Me.ObjOtorisasi_Overdue_Invoice.PrimaryTipeApproval = 0
            Me.ObjOtorisasi_Overdue_Invoice.Reason = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ResetObjOtorisasi_DOExceedCreditLimit()
        Try
            Me.ObjOtorisasi_DOExceedCreditLimit.Primary = -1
            Me.ObjOtorisasi_DOExceedCreditLimit.PrimaryApprovedBy = -1
            Me.ObjOtorisasi_DOExceedCreditLimit.PrimaryDokumenAcuan = -1
            Me.ObjOtorisasi_DOExceedCreditLimit.PrimaryRequestBy = -1
            Me.ObjOtorisasi_DOExceedCreditLimit.PrimaryTipeApproval = 0
            Me.ObjOtorisasi_DOExceedCreditLimit.Reason = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim Tolerance As Decimal = 0
        Dim CreditLimit As Decimal = 0
        Dim Dummy As String = ""
        Dim Dummy1 As String = ""
        Dim vFrmReminder As New FrmTransaksiSO_ReminderAR
        Dim Urutan As Integer = 0
        Dim rowDGV As DataGridViewRow

        Try
            Me.ResetObjOtorisasi_DOExceedCreditLimit()
            Me.ResetObjOtorisasi_OverdueInvoice()

            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("delivery order no")
            KoleksiTextBox.Add(Me.TxtDONo)
            KoleksiNamaData.Add("customer code")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                If Me.baru = False Then
                    If Me.objDO.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                                       Me.objDO.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        Me.TxtDONo.Text = Me.objDO.NoTransaksi
                    Else
                        If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                            Me.dtpDate.Value = Me.objDO.TanggalTransaksi
                            Me.TxtDONo.Text = Me.objDO.NoTransaksi
                            Return False
                        End If
                    End If
                End If
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.objDO.TimeStamp <> Me.objDO.GetTimestamp(Me.objDO.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objDO.CariDetail(Me.objDO.Primary, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                    Return False
                End If

                '   b. invoice
                If Me.ObjInvoiceItem.GetJmlByDO(Me.objDO.Primary) > 0 Then
                    MsgBox("There are sales invoice: " + Me.ObjInvoiceItem.GetSalesInvoiceNo(Me.objDO.Primary) + _
                        " refering to this transaction, cannot edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '   c. sales return 
                If DTA.ScalarQuery_GetJmlSalesReturn(Me.objDO.Primary).HasValue = True Then
                    If DTA.ScalarQuery_GetJmlSalesReturn(Me.objDO.Primary).Value > 0 Then
                        MsgBox("There are sales return refering to this transaction, cannot edit this data.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
                '   d. stock out sales
                If DTA.ScalarQuery_GetPrimaryStockOutSales(Me.objDO.Primary).HasValue = True Then
                    If DTA.ScalarQuery_GetPrimaryStockOutSales(Me.objDO.Primary).Value <> -1 Then
                        MsgBox("There are stock out sales refering to this transaction, cannot edit this data.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'cek exceed credit limit
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
            ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit) = "1" Then
                CreditLimit = Me.ObjCustomer.GetCreditLimit(Me.txtPrimaryCustomer.Text)
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit_Tolerance)
                If Dummy <> "" Then
                    Tolerance = Dummy
                Else
                    Tolerance = 0
                End If

                If CreditLimit <> 0 Then
                    If CreditLimit + Tolerance < Me.ObjSO.GetAR(Me.txtPrimaryCustomer.Text) Then
                        vFrmReminder.dgvItem.Rows.Clear()
                        vFrmReminder.vPrimaryCustomer = Me.txtPrimaryCustomer.Text
                        vFrmReminder.vTipeReminder = FrmTransaksiSO_ReminderAR.EnumTipeReminder.AR_Exceeds_Limit
                        vFrmReminder.vDate = Me.dtpDate.Value
                        vFrmReminder.vCreditLimit = CreditLimit
                        vFrmReminder.Tolerance_ARCreditLimit = Tolerance
                        vFrmReminder.FormAcuan = FrmTransaksiSO_ReminderAR.EnumFormAcuan._DO
                        vFrmReminder.vUserCode = Me.UserCode
                        vFrmReminder.ObjOtorisasiDO_ExceedCreditLimit = Me.ObjOtorisasi_DOExceedCreditLimit
                        If vFrmReminder.ShowDialog <> Windows.Forms.DialogResult.Yes Then
                            Return False
                        End If
                    End If
                End If
            End If

            'over due ar
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.SO, _
            ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue) = "1" Then
                Dummy1 = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, _
                ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue_Tolerance)
                If Dummy1 <> "" Then
                    Tolerance = Dummy1
                End If
                Dummy = Me.ObjSO.GetAR_Detail_OverDue(Me.txtPrimaryCustomer.Text, Me.dtpDate.Value - TimeSpan.FromDays(Tolerance))
                If Dummy <> "" Then
                    vFrmReminder.dgvItem.Rows.Clear()
                    vFrmReminder.vPrimaryCustomer = Me.txtPrimaryCustomer.Text
                    vFrmReminder.vTipeReminder = FrmTransaksiSO_ReminderAR.EnumTipeReminder.AR_Overdue
                    vFrmReminder.vDate = Me.dtpDate.Value - TimeSpan.FromDays(Tolerance)
                    vFrmReminder.FormAcuan = FrmTransaksiSO_ReminderAR.EnumFormAcuan._DO
                    vFrmReminder.vUserCode = Me.UserCode
                    vFrmReminder.ObjOtorisasiDO_OverdueInvoice = Me.ObjOtorisasi_Overdue_Invoice
                    If vFrmReminder.ShowDialog <> Windows.Forms.DialogResult.Yes Then
                        Return False
                    End If
                End If
            End If

            ''cek bisi ada qty default=0
            Dim ListError As String = ""
            For Each rowDGV In Me.dgvItem.Rows
                With rowDGV
                    If CDec(.Cells(Me.ColTotalQty.Index).Value) <= 0 Then
                        ListError += .Cells(Me.ColCodeNo.Index).Value & ", "
                    End If
                End With
            Next
            If ListError <> "" Then
                ListError = ListError.Substring(0, ListError.Length - 3) & "."
                MsgBox("The following product has zero default qty: " & vbNewLine & ListError & vbNewLine & "Please change qty or set ratio.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'jika ada delivery date, cek kombinasi per primary product dan primary so, delivery datenya harus sama
            If Me.ShowDeliveryDate = True Then
                Dim Cocok As Boolean = False
                Dim Ketemu As Boolean = False
                Dim TBL As New DseTransaksiDO.Cek_DeliveryDataTable
                Dim TBL1 As New DseTransaksiDO.Cek_DeliveryDataTable
                Dim Msg As String = ""

                Dim row As DseTransaksiDO.Cek_DeliveryRow
                For Each rowDGV In Me.dgvItem.Rows
                    If IsDate(rowDGV.Cells(Me.ColDeliveryDate.Index).Value) = True Then
                        Cocok = False
                        Ketemu = False

                        For Each row In TBL.Rows
                            If row.Primary_SO = rowDGV.Cells(Me.ColPrimarySO.Index).Value And _
                            row.Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                                Ketemu = True
                                If row.Delivery_Date.Date = CDate(rowDGV.Cells(Me.ColDeliveryDate.Index).Value) Then
                                    Cocok = True
                                Else
                                    Cocok = False
                                End If
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan baris baru untuk pengecekan
                            row = TBL.NewCek_DeliveryRow
                            row.Delivery_Date = CDate(rowDGV.Cells(Me.ColDeliveryDate.Index).Value)
                            row.Primary_SO = rowDGV.Cells(Me.ColPrimarySO.Index).Value
                            row.Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                            TBL.Rows.Add(row)
                        Else
                            If Cocok = False Then
                                Ketemu = False
                                For Each row In TBL1.Rows
                                    If row.Primary_SO = rowDGV.Cells(Me.ColPrimarySO.Index).Value And _
                                    row.Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                                        Ketemu = True
                                        Exit For
                                    End If
                                Next

                                'tambahkan baris baru untuk dikomplen kemudian
                                If Ketemu = False Then
                                    row = TBL1.NewCek_DeliveryRow
                                    row.Delivery_Date = CDate(rowDGV.Cells(Me.ColDeliveryDate.Index).Value)
                                    row.Primary_SO = rowDGV.Cells(Me.ColPrimarySO.Index).Value
                                    row.Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                                    TBL1.Rows.Add(row)
                                End If
                            End If
                        End If
                    End If
                Next
                If TBL1.Rows.Count > 0 Then
                    For Each row In TBL1.Rows
                        Urutan += 1
                        If Msg = "" Then
                            Msg = Urutan.ToString + ". SO: " + Me.ObjSO.GetNoSO(row.Primary_SO) + " - Code No: " + Me.ObjProduct.GetProductCode(row.Primary_Product)
                        Else
                            Msg = Msg + vbCrLf + Urutan.ToString + ". SO: " + Me.ObjSO.GetNoSO(row.Primary_SO) + " - Code No: " + Me.ObjProduct.GetProductCode(row.Primary_Product)
                        End If
                    Next
                    MsgBox("Please specify consistent delivery date for the following item: " + vbCrLf + Msg, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'cek qty multi vs qty total 
            For Each rowDGV In Me.dgvItem.Rows
                If CDec(rowDGV.Cells(Me.ColRatioDefault.Index).Value) = 1 And CDec(rowDGV.Cells(Me.ColRatioReference.Index).Value) = 1 Then
                    If CDec(rowDGV.Cells(Me.ColTotalQty.Index).Value) <> CDec(rowDGV.Cells(Me.ColTotalQtyMulti.Index).Value) Then
                        MsgBox("Inconsistent multi unit qty, ratio and default qty  for row: " + (rowDGV.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        rowDGV.Selected = True
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = rowDGV.Index

                        Return False
                    End If
                End If
            Next

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_OUT_SALES, Me.TBLDetailLocation) = False Then
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            DTA.Dispose()
            vFrmReminder.Dispose()
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
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub IsikanDefPrec(ByVal vTBL As DseTransaksiDO_AddByLocation.TblDefPrecDataTable)
        Dim DTAProduct As New DseTransaksiDO_AddByLocationTableAdapters.tblmaster_productTableAdapter
        Dim TBLProduct As New DseTransaksiDO_AddByLocation.tblmaster_productDataTable
        Dim rowProduct As DseTransaksiDO_AddByLocation.tblmaster_productRow
        Dim PrimaryMin, PrimaryMax As Double
        Dim rowDGV As DataGridViewRow
        Dim ListPrimaryProduct As New List(Of Double)
        Dim TblDefPrec As New DseTransaksiDO_AddByLocation.TblDefPrecDataTable
        Dim rowDefPrec1, rowDefPrec2 As DseTransaksiDO_AddByLocation.TblDefPrecRow

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
            For Each DV As DataRowView In TblDefPrec.DefaultView
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

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objDO
                .NoTransaksi = Me.TxtDONo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .PrimaryCustomer = Me.txtPrimaryCustomer.Text
                .CustAddIdentifier = Me.txtCustomerAddIdentifier.Text
                .Remarks = Me.TxtRemark.Text
                .CetakFaktur = Me.chkFakturPajak.Checked
                .CetakKwitansi = Me.chkKwitansi.Checked
                .IdentifierAddressShipTo = Me.TxtIdentifierAddressShipTo.Text
                .PrimaryCustomerShipTo = Me.TxtPrimaryCustomerShipTo.Text
                .IdentifierAddressCustomerShipTo = Me.TxtIdentifierAddressCustomerShipTo.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objDO
                Me.TxtDONo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryDO.Text = .Primary
                Me.txtPrimaryCustomer.Text = .PrimaryCustomer
                Me.txtCustomerAddIdentifier.Text = .CustAddIdentifier
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                Me.chkKwitansi.Checked = .CetakKwitansi
                Me.chkFakturPajak.Checked = .CetakFaktur

                If .PrimaryCustomerShipTo = -1 Then
                    Me.Panel1.Visible = True
                    Me.TxtIdentifierAddressShipTo.Text = .IdentifierAddressShipTo
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    If CboAddressShipTo.Items.Count > 0 Then
                        Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressShipTo.Text)
                    End If
                    Me.Panel2.Visible = False
                    Me.TxtPrimaryCustomerShipTo.Text = "-1"
                    Me.TxtCustomerCodeShipTo.Text = ""
                    Me.TxtCustomerDescShipTo.Text = ""
                    Me.TxtIdentifierAddressCustomerShipTo.Text = ""
                    Me.CboAddressCustomerShipTo.Items.Clear()
                    Me.Label6.Location = New Point(Me.Label6.Location.X, Me.Label1.Location.Y)
                    Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.TxtCustomerName.Location.Y + 5)
                    Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, (Me.TxtCustomerName.Size.Height * 2) + (Me.cboCustomerAddress.Top - Me.TxtCustomerName.Bottom))
                Else
                    Me.Panel2.Visible = True
                    Me.TxtPrimaryCustomerShipTo.Text = .PrimaryCustomerShipTo
                    Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(.PrimaryCustomerShipTo)
                    Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(.PrimaryCustomerShipTo)
                    Me.TxtIdentifierAddressCustomerShipTo.Text = .IdentifierAddressCustomerShipTo
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, .PrimaryCustomerShipTo)
                    If Me.CboAddressCustomerShipTo.Items.Count > 0 Then
                        Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressCustomerShipTo.Text)
                    End If
                    Me.Panel1.Visible = False
                    TxtIdentifierAddressShipTo.Text = ""
                    CboAddressShipTo.Items.Clear()
                    Me.Label6.Location = New Point(Me.Label6.Location.X, Me.cboCustomerAddress.Location.Y)
                    Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, Me.TxtCustomerName.Size.Height)
                    Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.cboCustomerAddress.Location.Y)
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()
            'digit grouping
            'Dim row As DataGridViewRow
            'For Each row In Me.dgvItem.Rows
            '    row.Cells(12).Value = Me.ObjCD.DigitGrouping(row.Cells(12).Value)
            '    If IsNumeric(row.Cells(14).Value) = True Then
            '        row.Cells(14).Value = Me.ObjCD.DigitGrouping(row.Cells(14).Value)
            '    End If
            'Next
            'set nilai acuan untuk pengecekan SO
            If Me.dgvItem.RowCount > 0 Then
                Dim PrimarySO As Decimal = Me.dgvItem.Item(Me.ColPrimarySO.Index, 0).Value
                Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(PrimarySO, Me.PrimaryCurrencySO, _
                Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
            End If

            'untuk custom admiral
            Dim ObjIdentifierCustom As New CustomSMESS.ClsIdentifier
            Dim IdentifierCustom As String = ObjIdentifierCustom.GetIdentifier()
            If IdentifierCustom = "smess_admiral2" Then
                Dim dta As New CustomSMESS.DseTransaksiDO_LocationTableAdapters.tbltrans_custom_doTableAdapter
                Me.TBLCustomItem_Loc.Rows.Clear()
                dta.Fill(Me.TBLCustomItem_Loc, CDec(Me.TxtPrimaryDO.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjLS As New clsMasterListedSoftware

        Try
            'Me.ColTotalQty.Visible = True
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'cek apakah inventory juga sudah diinstall
            Me.IsInventoryInstalled = ObjLS.IsInventoryInstalled
            If Me.IsInventoryInstalled = True Then
                Me.cboSecondaryQty.DropDownStyle = ComboBoxStyle.DropDownList
                Me.cboBatchNo.DropDownStyle = ComboBoxStyle.DropDownList
            End If
            Control.CheckForIllegalCrossThreadCalls = False
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable

            Me.cboSalesOrderNo.Width = Me.SplitContainer2.Panel1.Width - Me.cboSalesOrderNo.Left
            Me.cboCustPONO.Width = Me.cboSalesOrderNo.Width
            Me.llbDetailStock.Left = Me.cboSalesOrderNo.Left + Me.cboSalesOrderNo.Width - Me.llbDetailStock.Width
            Me.cboCodeNo.Width = Me.llbDetailStock.Left - 5 - Me.cboCodeNo.Left
            Me.txtDesc.Width = Me.cboSalesOrderNo.Width
            Me.LblUnitOB.Left = Me.llbDetailStock.Left
            Me.lblUnitsPendingDO.Left = Me.llbDetailStock.Left
            Me.TxtToleranceUp.Left = Me.cboCodeNo.Left + Me.cboCodeNo.Width - Me.TxtToleranceUp.Width
            Me.Label10.Left = Me.TxtToleranceUp.Left - 5 - Me.Label10.Width
            Me.txtQtyOrderBalance.Width = Me.Label10.Left - 5 - Me.txtQtyOrderBalance.Left
            Me.txtPendingDO.Width = Me.TxtToleranceUp.Left + Me.TxtToleranceUp.Width - Me.txtPendingDO.Left

            Me.cboBatchNo.Width = Me.SplitContainer2.Panel2.Width - 5 - Me.cboBatchNo.Left
            Me.DtpExpired.Width = Me.cboBatchNo.Width
            Me.BtnInput.Left = Me.cboBatchNo.Left + Me.cboBatchNo.Width - Me.BtnInput.Width
            Me.lblAvailableUnits.Left = Me.BtnInput.Left
            Me.lblUnitsDefault.Left = Me.lblAvailableUnits.Left
            Me.TxtRemarkItem.Width = Me.BtnInput.Left - 5 - Me.TxtRemarkItem.Left
            Me.TxtQtyperBatchDefault.Left = Me.lblUnitsDefault.Left - 5 - Me.TxtQtyperBatchDefault.Width
            Me.txtAvailable.Left = Me.TxtQtyperBatchDefault.Left
            Me.lblAvailableUnitsMulti.Left = Me.txtAvailable.Left - Me.lblAvailableUnitsMulti.Width - 20
            Me.lblUnitsMulti.Left = Me.lblAvailableUnitsMulti.Left
            Me.txtAvailableMulti.Width = Me.lblAvailableUnitsMulti.Left - 5 - Me.txtAvailableMulti.Left
            Me.TxtQtyperBatchMulti.Width = Me.txtAvailableMulti.Width
            Me.txtDesc.Width = Me.SplitContainer1.Panel1.Width - 5 - Me.txtDesc.Left
            Me.llbAddByLocation.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.llbAddByLocation.Width

            Me.ChkIgnoreTolerance.TabStop = False
            Me.dgvItem.ColumnHeadersHeight = 25
            Me.txtAvailable.TabStop = False
            Me.txtAvailableMulti.TabStop = False

            Me.buffer = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, ClsTransaksiSetting.EnumNamaParameter.DeliveryOrder_ShowDeliveryDate)
            If Me.buffer = "1" Then
                Me.ColDeliveryDate.Visible = True
                Me.ShowDeliveryDate = True
            Else
                Me.ColDeliveryDate.Visible = False
                Me.ShowDeliveryDate = False
            End If

            Me.ShowIcon = False
            Me.dgvItem.AllowUserToResizeRows = False
            Me.dgvItem.AllowUserToResizeColumns = False
            Me.dgvItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing

            'Update SMESS 2026
            If Me.vFromQuick Then
                Me.btnStopProcess.Enabled = True
                Me.prgProgress.Visible = True
                Me.btnStopProcess.Visible = True
                Me.btnStopProcess.TabStop = True
                Me.dtpDate.Enabled = True
                Me.btnFind.Visible = False
                Me.btnFind.TabStop = False
                Me.btnNew.Visible = False
                Me.btnNew.TabStop = False
                Me.btnEdit.Visible = False
                Me.btnEdit.TabStop = False
                Me.btnDelete.Visible = False
                Me.btnDelete.TabStop = False
                Me.btnExit.Visible = False
                Me.btnExit.TabStop = False
                Me.btnPrint.Visible = False
                Me.btnPrint.TabStop = False
                Me.MinimizeBox = False

                If vShowDialog = True Then
                    'If IsNumeric(Me.TxtPrimaryPihakAcuan.Text) = True Then
                    '    If Me.TxtPrimaryPihakAcuan.Text > 0 Then
                    '        For ij As Integer = Me.dgvItem.Rows.Count - 1 To 0 Step -1
                    '            rowDGV = Me.dgvItem.Rows(ij)
                    '            If rowDGV.Cells(Me.ColQtyMulti.Index).Value = 0 Then
                    '                Me.dgvItem.Rows.RemoveAt(ij)
                    '            End If
                    '        Next
                    '        Me.ObjCD.SetNumber(Me.dgvItem)
                    '    End If
                    'End If
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                Me.btnOK.Enabled = True
                Me.btnCancel.Enabled = True
                'Me.TxtKodePihakAcuan_Leave(sender, e)
                If vDirectlySave = True Then
                    Me.SimpanLangsung()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjLS = Nothing
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_New") = False Then
                        FrmO.AccessName = "FrmTransaksiDO_New"
                        FrmO.Text = "Approve: create new DO"
                        If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                            MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                        With Me.ObjOtorisasi
                            .Primary = 1
                            .PrimaryApprovedBy = Me.ObjUser.GetPrimary(FrmO.cboUserCode.Text)
                            .PrimaryDokumenAcuan = Me.TxtPrimaryDO.Text
                            .PrimaryRequestBy = Me.ObjUser.GetPrimary(Me.UserCode)
                            .PrimaryTipeApproval = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.DO_New
                            .Reason = FrmO.TxtReason.Text
                            .NoDocumentHist = Me.TxtDONo.Text
                        End With
                    End If

                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        'Update SMESS 2026
                        If vFromQuick Then
                            If Me.objDO.SimpanBaru1(Me.dgvItem, Me.UserCode, Me) = True Then
                                Me.KosongkanItem()
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                                Me.vPrimaryDO = Me.objDO.Primary
                                Me.vDONo = Me.objDO.NoTransaksi
                            End If
                            Me.Close()
                        Else
                            If Me.objDO.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.objDO.CariDetail(Me.objDO.Primary, Me.dgvItem)
                                Me.KosongkanItem()
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                                'update property primary dan no dokumen acuan, karena simpan baru
                                Me.ObjOtorisasi.PrimaryDokumenAcuan = Me.TxtPrimaryDO.Text
                                Me.ObjOtorisasi.NoDocumentHist = Me.TxtDONo.Text
                                Me.ObjOtorisasi.SimpanData()
                                'simpan otorisasi lain jika ada
                                Me.SimpanOtorisasiTambahan()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objDO.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            Me.SimpanOtorisasiTambahan()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            FrmO.Dispose()
            Me.ResetObjOtorisasi_DOExceedCreditLimit()
            Me.ResetObjOtorisasi_OverdueInvoice()
        End Try
    End Sub

    Private Sub SimpanOtorisasiTambahan()

        Try
            If Me.ObjOtorisasi_DOExceedCreditLimit.Primary <> -1 Then
                Me.ObjOtorisasi_DOExceedCreditLimit.Primary = 1
                Me.ObjOtorisasi_DOExceedCreditLimit.PrimaryDokumenAcuan = Me.TxtPrimaryDO.Text
                Me.ObjOtorisasi_DOExceedCreditLimit.NoDocumentHist = Me.TxtDONo.Text
                Me.ObjOtorisasi_DOExceedCreditLimit.SimpanData()
            End If
            If Me.ObjOtorisasi_Overdue_Invoice.Primary <> -1 Then
                Me.ObjOtorisasi_Overdue_Invoice.Primary = 1
                Me.ObjOtorisasi_Overdue_Invoice.PrimaryDokumenAcuan = Me.TxtPrimaryDO.Text
                Me.ObjOtorisasi_Overdue_Invoice.NoDocumentHist = Me.TxtDONo.Text
                Me.ObjOtorisasi_Overdue_Invoice.SimpanData()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.dtpDate.Value.Date = Today And _
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
                Me.KosongkanItem()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objDO.CariDetail(Me.TxtPrimaryDO.Text, Me.dgvItem)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()

                'Update SMESS 2026
                If vFromQuick = True Then
                    Me.Close()
                End If
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
            Me.KosongkanItem()
            'Me.dtpDate.Focus()
            Me.objDO.ResetObject(Me.dgvItem)
            Me.TxtDONo.Text = Me.objDO.GetNoBaru1(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, ClsTransaksiSetting.EnumNamaParameter.DeliveryOrder_AllowNoStock) = "1" Then
                Me.ApakahBolehDOTanpaStock = True
            Else
                Me.ApakahBolehDOTanpaStock = False
            End If
            Me.buffer = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.DeliveryOrder, ClsTransaksiSetting.EnumNamaParameter.DeliveryOrder_ShowDeliveryDate)
            If Me.buffer = "1" Then
                Me.ColDeliveryDate.Visible = True
                Me.ShowDeliveryDate = True
            Else
                Me.ColDeliveryDate.Visible = False
                Me.ShowDeliveryDate = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiDO
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                If Me.objDO.CariDetail(hj.PrimaryTransaksi, Me.dgvItem) = True Then
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
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.TxtDONo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDONo.Text <> "" Then
                If Me.objDO.ApakahBolehEdit(Me.TxtPrimaryDO.Text, Me.UserCode, True, True) = True Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_Edit") = False Then
                        FrmO.AccessName = "FrmTransaksiDO_Edit"
                        If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                            MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If

                    If Me.objDO.TimeStamp <> Me.objDO.GetTimestamp(Me.objDO.Primary) Then
                        Me.objDO.CariDetail(Me.objDO.Primary, Me.dgvItem)
                    End If
                    Me.baru = False
                    Me.AktifkanDetail()
                    'Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO.Dispose()
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.TxtDONo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDONo.Text <> "" Then
                If Me.objDO.ApakahBolehHapus(Me.TxtPrimaryDO.Text, Me.UserCode, True, True) = False Then
                    Me.btnFind.Focus()
                Else
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_Hapus") = False Then
                        FrmO.AccessName = "FrmTransaksiDO_Hapus"
                        If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                            MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If

                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objDO.Hapus(Me.TxtPrimaryDO.Text, Me.UserCode) = True Then
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
            FrmO.Dispose()
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
                            If Me.objDO.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.objDO.CariDetail(Me.objDO.Primary, Me.dgvItem)
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objDO.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
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

    Private Sub IsiCboSalesOrder(ByVal vPrimaryCustomer As Decimal)
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data awal
            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()
            Me.TBL5.Rows.Clear()
            Me.TBL6.Rows.Clear()
            Me.TBL7.Rows.Clear()
            Me.TBL8.Rows.Clear()

            'isikan dari basis data-persedian terakhir
            If Me.ChkIgnoreTolerance.Checked = False Then
                Me.ObjSOItem.IsikanTBLSO3_Tolerance(vPrimaryCustomer, Me.TBL, Me.cboSalesOrderNo)
            Else
                Me.ObjSOItem.IsikanTBLSO3_Default(vPrimaryCustomer, Me.TBL, Me.cboSalesOrderNo)
            End If

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiDO.DataTable2Row
            Dim row2 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjDOItem.IsikanTBLSO(Me.TxtPrimaryDO.Text, Me.TBL2)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL2.Rows.Count > 0
                    row1 = Me.TBL2.Rows(0)
                    ketemu = False
                    For Each row2 In Me.TBL.Rows
                        If row1.Primary_Customer = row2.Primary_Customer _
                        And row1.Primary_Product = row2.Primary_Product And _
                        row2.Primary_SO = row1.Primary_SO Then
                            row2.Order_Balance = row2.Order_Balance + row1.Total_Qty
                            row2.SOB_Multi = row2.SOB_Multi + row1.Total_Qty_Multi
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        If vPrimaryCustomer = row1.Primary_Customer Then 'modif 09/07/2020
                            'tambahkan baris baru
                            row2 = TBL.Newlaporan_sales_orderbalance_view9dRow
                            row2.Primary_SO = row1.Primary_SO
                            row2.Cust_PO_No = Me.ObjSO.GetCustPONo(row1.Primary_SO)
                            row2.Primary_Customer = row1.Primary_Customer
                            row2.Primary_Product = row1.Primary_Product
                            row2.Product_Code = row1.Product_Code
                            row2.SO_No = row1.So_No
                            row2.Order_Balance = row1.Total_Qty
                            row2.SOB_Multi = row1.Total_Qty_Multi
                            row2.Primary_Units_Multi = row1.Primary_Units_Multi
                            row2.Ratio_Reference = row1.Ratio_Reference
                            row2.Ratio_Default = row1.Ratio_Default
                            row2.Units_Code_Multi = row1.Units_Code_Multi
                            row2.Primary_Units_Default = row1.Primary_Units_Default
                            row2.Units_Code_Default = row1.Units_Code_Default
                            TBL.Rows.Add(row2)
                            'Me.TBL.Rows.Add(row1.Primary_SO, row1.Primary_Customer, row1.Primary_Product, _
                            'row1.Product_Code, row1.Total_Qty, row1.So_No)
                        End If
                    End If
                    Me.TBL2.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLSO()
            While Me.TBL2.Rows.Count > 0
                row1 = Me.TBL2.Rows(0)
                ketemu = False
                For Each row2 In Me.TBL.Rows
                    If row1.Primary_Customer = row2.Primary_Customer _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_SO = row1.Primary_SO Then
                        row2.Order_Balance = row2.Order_Balance - row1.Total_Qty
                        row2.SOB_Multi = row2.SOB_Multi - row1.Total_Qty_Multi
                        'If row2.Order_Balance = 0 Then
                        '    Me.TBL.Rows.Remove(row2)
                        'End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    'MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL2.Rows.Remove(row1)
            End While

            For i As Integer = TBL.Rows.Count - 1 To 0 Step -1
                row2 = TBL.Rows(i)
                If Me.ChkIgnoreTolerance.Checked = True Then
                    If row2.SOB_Multi + row2.Tolerance_Up <= 0 Then
                        Me.TBL.Rows.Remove(row2)
                    End If
                Else
                    If row2.SOB_Multi - row2.Tolerance_Down <= 0 Then
                        Me.TBL.Rows.Remove(row2)
                    End If
                End If
            Next

            'pastikan hanya 1 so disebutkan hanya 1x
            Dim idx As Integer = 0
            Dim idx2 As Integer = 0
            Dim dummyRow As DseTransaksiSO.laporan_sales_orderbalance_view9dRow
            While idx < Me.TBL.Rows.Count
                row2 = Me.TBL(idx)
                idx2 = idx + 1
                While idx2 <= Me.TBL.Rows.Count - 1
                    dummyRow = Me.TBL.Rows(idx2)
                    If dummyRow.SO_No = row2.SO_No Then
                        Me.TBL.Rows.Remove(dummyRow)
                    Else
                        idx2 += 1
                    End If
                End While

                'pindah ke baris selanjutnya
                idx += 1
            End While

            Me.TBL.DefaultView.Sort = "[SO No] ASC"
            Me.cboSalesOrderNo.DataSource = Me.TBL.DefaultView
            Me.cboSalesOrderNo.ValueMember = "SO No"
            Me.cboSalesOrderNo.DisplayMember = "SO No"
            Me.cboSalesOrderNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub IsikanTBLSO()
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiDO.DataTable2Row
            For Each row In Me.dgvItem.Rows
                Dim ketemu As Boolean = False
                For Each row1 In Me.TBL2.Rows
                    If row1.Primary_Customer = Me.txtPrimaryCustomer.Text And _
                    row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                    row1.Primary_SO = row.Cells(Me.ColPrimarySO.Index).Value Then
                        'kalo ketemu ubah data order balance
                        row1.Total_Qty = row1.Total_Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                        row1.Total_Qty_Multi = row1.Total_Qty_Multi + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                        ketemu = True
                        Exit For
                    End If
                Next
                'kalo ga ketemu tambahkan baris baru
                If ketemu = False Then
                    Me.TBL2.Rows.Add(Me.TxtPrimaryDO.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimarySO.Index).Value, row.Cells(Me.ColTotalQty.Index).Value, _
                        Me.txtPrimaryCustomer.Text, row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColSONO.Index).Value, Me.ObjSO.GetCustPONo(row.Cells(Me.ColPrimarySO.Index).Value), _
                        row.Cells(Me.ColTotalQtyMulti.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value, row.Cells(Me.ColRatioReference.Index).Value, _
                        row.Cells(Me.ColRatioDefault.Index).Value, row.Cells(Me.ColUnitsCodeMulti.Index).Value, row.Cells(Me.ColPrimaryUnit.Index).Value, _
                        row.Cells(Me.ColUnit.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboProduct(ByVal vPrimarySO As Decimal)
        Try
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()
            Me.TBL5.Rows.Clear()
            Me.TBL6.Rows.Clear()
            Me.TBL7.Rows.Clear()
            Me.TBL8.Rows.Clear()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'isikan dari basis data-persedian terakhir
            Me.cboCodeNo.DataSource = Nothing
            If Me.ChkIgnoreTolerance.Checked = False Then
                Me.ObjSOItem.IsikanTBLProduct2_Tolerance(Me.txtPrimarySO.Text, Me.TBL3)
            Else
                Me.ObjSOItem.IsikanTBLProduct2_Default(Me.txtPrimarySO.Text, Me.TBL3)
            End If

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiDO.DataTable2Row
            Dim row2 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjDOItem.IsikanTBLProductBySO(Me.TxtPrimaryDO.Text, Me.TBL4, Me.txtPrimarySO.Text)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    ketemu = False
                    For Each row2 In Me.TBL3.Rows
                        If row1.Primary_Customer = row2.Primary_Customer _
                        And row1.Primary_Product = row2.Primary_Product And _
                        row2.Primary_SO = row1.Primary_SO Then
                            row2.Order_Balance = row2.Order_Balance + row1.Total_Qty
                            row2.SOB_Multi = row2.SOB_Multi + row1.Total_Qty_Multi
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        row2 = TBL3.Newlaporan_sales_orderbalance_view9dRow
                        With row2
                            .Primary_SO = row1.Primary_SO
                            .Primary_Customer = row1.Primary_Customer
                            .Primary_Product = row1.Primary_Product
                            .Product_Code = row1.Product_Code
                            .SO_No = row1.So_No
                            .Cust_PO_No = row1.Cust_PO_No
                            .Order_Balance = row1.Total_Qty
                            .Primary_Units_Multi = row1.Primary_Units_Multi
                            .Ratio_Reference = row1.Ratio_Reference
                            .Ratio_Default = row1.Ratio_Default
                            .SOB_Multi = row1.Total_Qty_Multi
                            .Units_Code_Multi = row1.Units_Code_Multi
                            .Primary_Units_Default = row1.Primary_Units_Default
                            .Units_Code_Default = row1.Units_Code_Default
                        End With
                        Me.TBL3.Rows.Add(row2)
                        'Me.TBL3.Rows.Add(row1.Primary_SO, row1.Primary_Customer, row1.Primary_Product, _
                        'row1.Product_Code, row1.Total_Qty, row1.So_No, row1.Cust_PO_No)
                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsiTBLProduct(Me.txtPrimarySO.Text)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                ketemu = False
                For Each row2 In Me.TBL3.Rows
                    If row1.Primary_Customer = row2.Primary_Customer _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_SO = row1.Primary_SO Then
                        row2.Order_Balance = row2.Order_Balance - row1.Total_Qty
                        row2.SOB_Multi = row2.SOB_Multi - row1.Total_Qty_Multi
                        'If row2.Order_Balance = 0 Then
                        '    Me.TBL3.Rows.Remove(row2)
                        'End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL3.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    'MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4.Rows.Remove(row1)
            End While

            For i As Integer = TBL3.Rows.Count - 1 To 0 Step -1
                row2 = TBL3.Rows(i)
                If Me.ChkIgnoreTolerance.Checked = True Then
                    If row2.SOB_Multi + row2.Tolerance_Up <= 0 Then
                        Me.TBL3.Rows.Remove(row2)
                    End If
                Else
                    If row2.SOB_Multi - row2.Tolerance_Down <= 0 Then
                        Me.TBL3.Rows.Remove(row2)
                    End If
                End If
            Next

            Me.TBL3.DefaultView.Sort = "[Product Code] ASC"
            Me.cboCodeNo.DataSource = Me.TBL3.DefaultView
            Me.cboCodeNo.DisplayMember = "Product Code"
            Me.cboCodeNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTBLProduct(ByVal vPrimarySO As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiDO.DataTable2Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Primary_Customer = Me.txtPrimaryCustomer.Text And _
                        row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        row1.Primary_SO = row.Cells(Me.ColPrimarySO.Index).Value Then
                            'kalo ketemu ubah data order balance
                            row1.Total_Qty = row1.Total_Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            row1.Total_Qty_Multi = row1.Total_Qty_Multi + CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL4.Rows.Add(Me.TxtPrimaryDO.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimarySO.Index).Value, row.Cells(Me.ColTotalQty.Index).Value, _
                        Me.txtPrimaryCustomer.Text, row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColSONO.Index).Value, Me.ObjSO.GetCustPONo(row.Cells(Me.ColPrimarySO.Index).Value), _
                        row.Cells(Me.ColTotalQtyMulti.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value, row.Cells(Me.ColRatioReference.Index).Value, row.Cells(Me.ColRatioDefault.Index).Value, _
                        row.Cells(Me.ColUnitsCodeMulti.Index).Value, row.Cells(Me.ColPrimaryUnit.Index).Value, row.Cells(Me.ColUnit.Index).Value)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboBatchNo(ByVal vPrimaryProduct As Decimal, ByVal vPrimarySO As Decimal)
        Try
            Me.TBL5.Rows.Clear()
            Me.TBL6.Rows.Clear()
            Me.TBL7.Rows.Clear()
            Me.TBL8.Rows.Clear()
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNo, Me.TBL5)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiDO.tbltrans_do_item1Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjDOItem.IsiTBLBatchNo(Me.TxtPrimaryDO.Text, vPrimaryProduct, Me.TBL6)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL6.Rows.Count > 0
                    row1 = Me.TBL6.Rows(0)
                    For Each row2 In Me.TBL5.Rows
                        If row1.Lot_No = row2.Lot_No Then
                            row2.Qty = row2.Qty + row1.Total_Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL5.Rows.Add(row1.Lot_No, row1.Expiry_Date, row1.Total_Qty)
                    End If
                    Me.TBL6.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLBatchNo(vPrimaryProduct)
            While Me.TBL6.Rows.Count > 0
                row1 = Me.TBL6.Rows(0)
                For Each row2 In Me.TBL5.Rows
                    If row1.Lot_No = row2.Lot_No Then
                        row2.Qty = row2.Qty - row1.Total_Qty
                        If row2.Qty = 0 Then
                            Me.TBL5.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL5.Rows.Count > 0 And row1.Lot_No <> "" Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL6.Rows.Remove(row1)
            End While


            If Me.chkExpiryDateControlled.Checked = False Then
                Me.TBL5.DefaultView.Sort = "[Lot No] ASC"
            Else
                Me.TBL5.DefaultView.Sort = "[Expiry Date] ASC"
            End If

            Me.cboBatchNo.DataSource = Me.TBL5.DefaultView
            Me.cboBatchNo.ValueMember = "Lot No"
            Me.cboBatchNo.DisplayMember = "Lot No"
            Me.cboBatchNo.Refresh()
            'pilih salah satu
            'If Me.cboBatchNo.Items.Count > 0 Then
            'Me.cboBatchNo.SelectedIndex = -1
            'Me.cboBatchNo.SelectedIndex = 0
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiDO.tbltrans_do_item1Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL6.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColBatchNo.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Total_Qty = row1.Total_Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = True Then
                            ED = row.Cells(Me.ColExpiryDate.Index).Value
                        End If
                        Me.TBL6.Rows.Add(CStr(row.Cells(Me.ColBatchNo.Index).Value), ED, CDbl(row.Cells(Me.ColTotalQty.Index).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboSecondary(ByVal vPrimaryProduct As Decimal, ByVal vBatchNo As String)
        Try
            Me.TBL7.Rows.Clear()
            Me.TBL8.Rows.Clear()
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboSecondaryUnit(vPrimaryProduct, Me.cboSecondaryQty, _
            Me.TBL7, vBatchNo)
            Dim row1 As DseTransaksiDO.tbltrans_do_item2Row
            Dim row2 As dseMasterProduct.tblmaster_productqty3Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjDOItem.IsiSecondaryQty(Me.TxtPrimaryDO.Text, vPrimaryProduct, vBatchNo, _
                Me.TBL8)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL8.Rows.Count > 0
                    row1 = Me.TBL8.Rows(0)
                    ketemu = False
                    For Each row2 In Me.TBL7.Rows
                        If row1.Secondary_Qty = row2.Qty_Secondary Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL7.Rows.Add(row1.Secondary_Qty, row1.Qty)
                    End If
                    Me.TBL8.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLSecondary(vPrimaryProduct, vBatchNo)
            While Me.TBL8.Rows.Count > 0
                row1 = Me.TBL8.Rows(0)
                ketemu = False
                For Each row2 In Me.TBL7.Rows
                    If row1.Secondary_Qty = row2.Qty_Secondary Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL7.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL7.Rows.Count > 0 And row1.Secondary_Qty <> -1 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL8.Rows.Remove(row1)
            End While

            'set data source
            Me.TBL7.DefaultView.Sort = "[Qty Secondary] DESC"
            Me.cboSecondaryQty.DataSource = Me.TBL7.DefaultView
            Me.cboSecondaryQty.ValueMember = "Qty Secondary"
            Me.cboSecondaryQty.DisplayMember = "Qty Secondary"
            Me.cboSecondaryQty.Refresh()

            'pilih salah satu
            If Me.cboSecondaryQty.Items.Count > 0 Then
                Me.cboSecondaryQty.SelectedIndex = -1
                Me.cboSecondaryQty.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLSecondary(ByVal vPrimaryProduct As Decimal, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiDO.tbltrans_do_item2Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(17).Value = vPrimaryProduct And CStr(row.Cells(6).Value) = vBatchNo _
                And IsNumeric(row.Cells(14).Value) = True Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL8.Rows
                        If row1.Secondary_Qty = CDbl(row.Cells(14).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(12).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL8.Rows.Add(CDbl(row.Cells(14).Value), CDbl(row.Cells(12).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'perubahan ada pada satu fungsi SetGboItem
    Private Sub SetGboItem(Optional ByVal vIndexSO As Double = -1, Optional ByVal vIndexProduct As Double = -1)
        Try
            Me.cboBatchNo.DataSource = Nothing
            Me.cboSecondaryQty.DataSource = Nothing
            Me.cboSalesOrderNo.DataSource = Nothing
            Me.cboSecondaryQty.DataSource = Nothing
            Me.cboCustPONO.DataSource = Nothing

            'sales order & product code
            Me.IsiCboSalesOrder(Me.txtPrimaryCustomer.Text)

            If vIndexSO <> -1 And Me.TBL.DefaultView.Count > vIndexSO Then
                'geser po ke yg lama
                Me.cboSalesOrderNo.SelectedIndex = vIndexSO
                'geser product
                If vIndexProduct <> -1 And Me.TBL3.DefaultView.Count > vIndexProduct Then
                    Me.cboCodeNo.SelectedIndex = vIndexProduct
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function GetIndex2(ByVal vPrimaryProduct As Decimal, ByVal BatchNo As String, _
    ByVal SecondaryQty As Decimal, ByVal ExpiryDate As Date, ByVal vPrimarySO As Decimal) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty _
                        And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) Then
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

    Private Function GetIndex3(ByVal vPrimaryProduct As Decimal, ByVal BatchNo As String, _
  ByVal SecondaryQty As Decimal, ByVal ExpiryDate As Date, ByVal vPrimarySO As Decimal, ByVal vPrimaryLocation As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty _
                        And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
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

    ''' <summary>
    ''' mengembalikan index baris dalam dgv yang sesuai dengan parameter yang diberikan
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="BatchNo"></param>
    ''' <param name="SecondaryQty"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Decimal, ByVal BatchNo As String, _
    ByVal SecondaryQty As Decimal, ByVal ExpiryDate As Date) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty _
                        And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatchNo.Index).Value = BatchNo And row.Cells(Me.ColSecQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiryDate.Index).Value) Then
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

    Private Function GetDeliveryDate(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Date
        Dim RowDgv As DataGridViewRow
        Dim RetVal As Date = #1/1/1950#

        Try
            For Each RowDgv In Me.dgvItem.Rows
                If RowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And RowDgv.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO Then
                    If IsDate(RowDgv.Cells(Me.ColDeliveryDate.Index).Value) = False Then
                        Return #1/1/1950#
                    Else
                        Return RowDgv.Cells(Me.ColDeliveryDate.Index).Value
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return RetVal
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim RowDgv As DataGridViewRow
        Dim TxtSOLama As String = ""
        Dim frmoptionDO As New FrmTransaksiDO3_Option
        Dim DeliveryDate As Date = #1/1/1950#

        Try
            If Me.ApakahDataItemBenar = True Then
                Me.txtCustomerCode.Enabled = False
                Me.btnFindCustomer.Enabled = False

                Dim SecQty As Decimal = -1
                If Me.txtSecondaryQty.Text = "" Or Me.txtSecondaryQty.Text = "-1" Then
                    SecQty = -1
                Else
                    SecQty = Me.txtSecondaryQty.Text
                End If
                Dim Index As Integer = Me.GetIndex2(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text, _
                SecQty, Me.DtpExpired.Value, Me.txtPrimarySO.Text)

                'cari data yang sama persis
                Dim DefPrecMulti As Integer = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                Dim DefPrecDefault As Integer = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    'row.Cells(12).Value = CDbl(row.Cells(12).Value) + CDbl(Me.TxtQtyperBatchMulti.Text)
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatchMulti.Text), DefPrecMulti)
                    Me.SetQtyDefaultOnDgv(row.Index)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If

                    'TR/Tambahan Form option untuk pilihan delivery date
                    If Me.ShowDeliveryDate = True Then
                        DeliveryDate = Me.GetDeliveryDate(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)

                        If DeliveryDate = #1/1/1950# Then
                            frmoptionDO.PrimarySO = Me.txtPrimarySO.Text
                            frmoptionDO.PrimaryProd = Me.txtPrimaryCodeNo.Text
                            frmoptionDO.ShowDialog()
                            If frmoptionDO.DialogResult = Windows.Forms.DialogResult.Cancel Then
                                Exit Sub
                            Else
                                If frmoptionDO.cbo_DeliveryDate.Text = "" Then
                                    DeliveryDate = #1/1/1950#
                                Else
                                    DeliveryDate = frmoptionDO.cbo_DeliveryDate.Text
                                End If
                            End If
                        End If
                    End If

                    Me.dgvItem.Rows.Add()
                    RowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(Me.ColPrimaryLocation.Index).Value = -1
                        .Cells(Me.ColLocation.Index).Value = ""
                        .Cells(Me.ColWarehouse.Index).Value = ""

                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColCustCodeNo.Index).Value = Me.cboCustCodeNo.Text
                        .Cells(Me.ColCustDesc.Index).Value = Me.txtCustProductDesc.Text
                        .Cells(Me.ColSONO.Index).Value = Me.cboSalesOrderNo.Text
                        .Cells(Me.ColBatchNo.Index).Value = Me.cboBatchNo.Text
                        .Cells(Me.ColExpiryDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, DefPrecDefault)
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColSecQty.Index).Value = SecQty
                        .Cells(Me.ColSecUnit.Index).Value = Me.lblSecondaryUnits.Text
                        .Cells(Me.ColRemarks.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimarySO.Index).Value = Me.txtPrimarySO.Text
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryUnit.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecUnit.Index).Value = Me.txtPrimarySecondaryUnits.Text
                        .Cells(Me.ColTotalQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                        .Cells(Me.ColUnitsCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                        .Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, DefPrecDefault)
                        .Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(Me.TxtRatioReference.Text, DefPrecMulti)
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                        If Me.txtPrimaryUnitsDefault.Text <> Me.txtPrimaryUnitsMulti.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, DefPrecMulti) _
                            & " " & Me.lblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, DefPrecDefault) & " " & Me.lblUnitsDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        'TR/Tambahan isi kolom delivery date
                        If Me.ShowDeliveryDate = True Then
                            If DeliveryDate = #1/1/1950# Then
                                .Cells(Me.ColDeliveryDate.Index).Value = "-"
                            Else
                                .Cells(Me.ColDeliveryDate.Index).Value = Format(DeliveryDate, "dd MMM yyyy")
                            End If
                        Else
                            .Cells(Me.ColDeliveryDate.Index).Value = "-"
                        End If

                        Dim DtaSOAddress As New DseTransaksiSOTableAdapters.tbltrans_so3TableAdapter
                        Dim TblSOAddress As New DseTransaksiSO.tbltrans_so3DataTable
                        Dim RowSOAddress As DseTransaksiSO.tbltrans_so3Row
                        ''01 apakah default atau other customer
                        DtaSOAddress.Connection.ConnectionString = Core_Module.GetConstring
                        If DtaSOAddress.Fill(TblSOAddress, Me.txtPrimarySO.Text) > 0 Then
                            RowSOAddress = TblSOAddress.Rows(0)
                            ''01b update alamat bill to
                            Me.txtCustomerAddIdentifier.Text = RowSOAddress.Indentifier_Alamat_Customer
                            If Me.cboCustomerAddress.Items.Count > 0 Then
                                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                            End If
                            If RowSOAddress.primary_customer_ship_to = -1 Then
                                ''02 jika default, isi ship to address
                                Me.Panel1.Visible = True
                                Me.TxtIdentifierAddressShipTo.Text = RowSOAddress.identifier_address_ship_to
                                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                                If CboAddressShipTo.Items.Count > 0 Then
                                    Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressShipTo.Text)
                                End If
                                Me.Panel2.Visible = False
                                Me.TxtPrimaryCustomerShipTo.Text = "-1"
                                Me.TxtCustomerCodeShipTo.Text = ""
                                Me.TxtCustomerDescShipTo.Text = ""
                                TxtIdentifierAddressCustomerShipTo.Text = ""
                                CboAddressCustomerShipTo.Items.Clear()
                                Me.Label6.Location = New Point(Me.Label6.Location.X, Me.Label1.Location.Y)
                                Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.TxtCustomerName.Location.Y + 5)
                                Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, (Me.TxtCustomerName.Size.Height * 2) + (Me.cboCustomerAddress.Top - Me.TxtCustomerName.Bottom))
                            Else
                                ''03 jika other customer, isi customer ship to dan customer ship to address
                                Me.Panel2.Visible = True
                                Me.TxtPrimaryCustomerShipTo.Text = RowSOAddress.primary_customer_ship_to
                                Me.TxtCustomerCodeShipTo.Text = RowSOAddress.Customer_Code
                                Me.TxtCustomerDescShipTo.Text = RowSOAddress.Customer_Name
                                Me.TxtIdentifierAddressCustomerShipTo.Text = RowSOAddress.identifier_address_customer_ship_to
                                Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, RowSOAddress.primary_customer_ship_to)
                                If Me.CboAddressCustomerShipTo.Items.Count > 0 Then
                                    Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressCustomerShipTo.Text)
                                End If
                                Me.Panel1.Visible = False
                                TxtIdentifierAddressShipTo.Text = ""
                                CboAddressShipTo.Items.Clear()
                                Me.Label6.Location = New Point(Me.Label6.Location.X, Me.cboCustomerAddress.Location.Y)
                                Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, Me.TxtCustomerName.Size.Height)
                                Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.cboCustomerAddress.Location.Y)
                            End If
                        End If
                    End With

                    'Me.dgvItem.Rows.Add("", Me.cboCodeNo.Text, Me.txtDesc.Text, Me.cboCustCodeNo.Text, _
                    'Me.txtCustProductDesc.Text, Me.cboSalesOrderNo.Text, Me.cboBatchNo.Text, ExpiredDateStr, _
                    '0, 0, "", 0, Me.TxtQtyperBatchMulti.Text, Me.lblUnitsMulti.Text, SecQty, Me.lblSecondaryUnits.Text, _
                    'Me.TxtRemarkItem.Text, Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text, -1, Me.txtPrimaryUnitsDefault.Text, _
                    'Me.txtPrimarySecondaryUnits.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If Me.dgvItem.RowCount = 1 Then
                        'set nilai acuan untuk pengecekan SO
                        Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.txtPrimarySO.Text, Me.PrimaryCurrencySO, _
                        Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
                    End If
                End If

                'kendalikan focus item selanjutnya
                Dim JmlObjLama, JmlSOBBaru, JmlDGV, JmlSOB As Decimal
                JmlDGV = Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
                JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)
                JmlObjLama = Me.ObjDOItem.GetJml_Multi(Me.TxtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)

                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlSOBBaru = JmlSOB + JmlObjLama - JmlDGV
                Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(0, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping(0, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
                Me.TxtRemarkItem.Text = ""

                TxtSOLama = Me.cboSalesOrderNo.Text
                Me.SetGboItem(Me.cboSalesOrderNo.SelectedIndex, Me.cboCodeNo.SelectedIndex)

                Me.cboCodeNo.Focus()

                'format dgv
                Me.FormatDGV()

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

    Private Sub SetJmlBO_N_FreeQty()
        Dim QtyOnHand, JmlBookingAll, JmlBookingSO, JmlObjLama, JmlDGV, JmlPendingSendiri, JmlPendingOther As Decimal
        Dim BatchNo As String = ""
        Dim row1 As dseMasterProduct.tblmaster_productqty1Row
        Dim rowview As DataRowView
        Dim mDec As Integer = 0

        Try
            If Me.cboBatchNo.SelectedIndex = -1 Then
                BatchNo = ""
            Else
                rowview = Me.cboBatchNo.Items(Me.cboBatchNo.SelectedIndex)
                row1 = rowview.Row
                BatchNo = row1.Lot_No
            End If

            'jml pending do
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.txtPendingDO.Text = Me.ObjCD.DigitGrouping(Me.ObjDOItem.GetJmlPendingDO3_Multi(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text, -1) - _
                Me.ObjDOItem.GetJml4_Multi(Me.TxtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text) + _
                Me.GetJmlDGV2_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text, Me.cboBatchNo.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
            Else
                Me.txtPendingDO.Text = Me.ObjCD.DigitGrouping(Me.ObjDOItem.GetJmlPendingDO_Multi(Me.txtPrimaryCodeNo.Text) - _
                Me.ObjDOItem.GetJml_Multi(Me.TxtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text) + _
                Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text), _
                Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
            End If

            'jml sales order balance
            'Me.txtQtyOrderBalance.Text = Me.ObjCD.DigitGrouping( Me.ObjSOItem.GetSalesOrderBalance(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text) _
            '+ Me.ObjDOItem.GetJml(Me.TxtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text) _
            '- Me.GetJmlDGV(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text) , _   ' - _
            'Me.ObjDOItem.GetJmlPendingDO_Other(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryDO.Text), _
            'Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
            Me.txtQtyOrderBalance.Text = Me.ObjCD.DigitGrouping(Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text) _
            + Me.ObjDOItem.GetJml_Multi(Me.TxtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text) _
            - Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))

            'jml free dilihat dari sisi inventory, meperhitungkan juga jumlah dalam database
            'jml free = qty on hand - jml booking + jml booking untuk so ybs - 
            '(JmlPendingDOSendiri - (jml obj lama - jml dgv)) - JmlPendingOtherDO
            If Me.txtSecondaryQty.Text = "" Then
                Me.txtSecondaryQty.Text = "-1"
            End If
            QtyOnHand = Me.ObjProduct.GetJml4(Me.txtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, BatchNo)
            JmlBookingAll = Me.ObjBookingInv.GetJml5(Me.txtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, BatchNo)
            JmlBookingSO = Me.ObjBookingInv.GetJml4(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text, Me.txtSecondaryQty.Text, BatchNo)
            If Me.baru = False Then
                JmlObjLama = Me.ObjDOItem.GetJml2(Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryDO.Text, Me.txtSecondaryQty.Text, BatchNo)
            Else
                JmlObjLama = 0
            End If
            JmlDGV = Me.GetJmlDGV_Default(Me.txtPrimaryCodeNo.Text, BatchNo, Me.txtSecondaryQty.Text)

            'Me.txtQTyFree.Text = Me.ObjCD.DigitGrouping(QtyOnHand - JmlBookingAll + _
            'JmlBookingSO + JmlObjLama - JmlDGV)

            JmlPendingSendiri = Me.ObjDOItem.GetJml3(Me.TxtPrimaryDO.Text, Me.txtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, BatchNo)
            JmlPendingOther = Me.ObjDOItem.GetJmlPendingDO_Other(Me.txtPrimaryCodeNo.Text, BatchNo, Me.txtSecondaryQty.Text, Me.TxtPrimaryDO.Text)

            Me.txtAvailable.Text = Me.ObjCD.DigitGrouping(QtyOnHand - JmlBookingAll + _
            JmlBookingSO - (JmlPendingSendiri - (JmlObjLama - JmlDGV)) - JmlPendingOther, _
            Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))


            If Me.TxtRatioDefault.Text = "0" Or Me.TxtRatioDefault.Text = "" Then
                Me.TxtRatioDefault.Text = "1"
            End If
            If Me.TxtRatioReference.Text = "0" Or Me.TxtRatioReference.Text = "" Then
                Me.TxtRatioReference.Text = "1"
            End If
            Me.txtAvailableMulti.Text = ObjCD.DigitGrouping((CDec(Me.txtAvailable.Text) * CDec(Me.TxtRatioReference.Text)) / CDec(Me.TxtRatioDefault.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text), ClsCekData.EnumTipeRounding.Up)
            If Me.lblAvailableUnits.Text = Me.lblAvailableUnitsMulti.Text Then
                Me.txtAvailable.Visible = False
                Me.txtAvailable.TabStop = False
                Me.lblAvailableUnits.Visible = False
            Else
                Me.txtAvailable.TabStop = True
                Me.txtAvailable.Visible = True
                Me.lblAvailableUnits.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SetFieldToleranceUp()
        End Try
    End Sub

    Private Function GetJmlAvailable(ByVal vPrimaryProduct As Decimal, _
    ByVal vPrimarySO As Decimal, ByVal vBatchNo As String, ByVal vSecQty As Decimal) As Decimal
        Dim QtyOnHand, JmlBookingAll, JmlBookingSO, JmlObjLama, JmlDGV, _
        JmlPendingSendiri, JmlPendingOther, JmlAvailable As Decimal

        Try
            QtyOnHand = Me.ObjProduct.GetJml4(vPrimaryProduct, _
            vSecQty, vBatchNo)
            JmlBookingAll = Me.ObjBookingInv.GetJml5(vPrimaryProduct, _
            vSecQty, vBatchNo)
            JmlBookingSO = Me.ObjBookingInv.GetJml4(vPrimaryProduct, Me.txtPrimarySO.Text, _
            vSecQty, vBatchNo)
            If Me.baru = False Then
                JmlObjLama = Me.ObjDOItem.GetJml2(vPrimaryProduct, Me.TxtPrimaryDO.Text, _
                vSecQty, vBatchNo)
            Else
                JmlObjLama = 0
            End If
            JmlDGV = Me.GetJmlDGV_Default(vPrimaryProduct, vBatchNo, vSecQty)
            'Me.txtQTyFree.Text = Me.ObjCD.DigitGrouping(QtyOnHand - JmlBookingAll + _
            'JmlBookingSO + JmlObjLama - JmlDGV)
            JmlPendingSendiri = Me.ObjDOItem.GetJml3(Me.TxtPrimaryDO.Text, vPrimaryProduct, _
            vSecQty, vBatchNo)
            JmlPendingOther = Me.ObjDOItem.GetJmlPendingDO_Other(vPrimaryProduct, vBatchNo, _
            vSecQty, Me.TxtPrimaryDO.Text)
            JmlAvailable = QtyOnHand - JmlBookingAll + _
            JmlBookingSO - (JmlPendingSendiri - (JmlObjLama - JmlDGV)) - JmlPendingOther

            Return JmlAvailable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Dim IsMultiUnit As Boolean = False

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = False Then
                    row.Cells(Me.ColExpiryDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiryDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiryDate.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiryDate.Index).Value = Format(CDate(row.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy")
                    End If
                End If

                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                If row.Cells(Me.ColSecQty.Index).Value = "-1" Then
                    row.Cells(Me.ColSecQty.Index).Value = "-"
                    row.Cells(Me.ColSecUnit.Index).Value = "-"
                End If

                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatchNo.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If

                '1.e. tandai bila tidak ada lokasi yg kosong
                If row.Cells(Me.ColPrimaryLocation.Index).Value <> -1 Then
                    TidakAdaLokasi = False
                End If
            Next

            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = True
                Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = True
            End If
            '3. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColBatchNo.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColBatchNo.Index).Visible = True
            End If
            '4. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiryDate.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiryDate.Index).Visible = True
            End If
            '5. ada lokasi
            If TidakAdaLokasi = False Then
                Me.ColLocation.Visible = True
                Me.ColWarehouse.Visible = True
            Else
                Me.ColLocation.Visible = False
                Me.ColWarehouse.Visible = False
            End If

            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnit.Index).Value <> .Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
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

            'Me.ColTotalQty.Visible = True
            'Me.ColRatioDesc.Visible = True

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
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

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = ColTotalQtyMulti.Index Or e.ColumnIndex = ColSecQty.Index Or e.ColumnIndex = Me.ColBatchNo.Index _
                Or e.ColumnIndex = ColRemarks.Index Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = ColSecQty.Index Then
                        'cek secondary qty
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColBatchNo.Index Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                        'ahdar
                        If objCustomDO.ApakahEditDgvBenar_Custom(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, True) = False Then
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

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim frmRatio As New FrmTransaksiSO_SetQtyRatio
        Dim row As DataGridViewRow
        Dim BufferQtyDefault As String = ""
        Dim BufferRatioReference As String = ""
        Dim BufferRatioDefault As String = ""
        Dim BufferRatioDesc As String = ""
        Dim BufferLblQty As String = ""
        Dim DefPrec As Integer = 0
        Dim lanjut As Boolean = True
        Dim SecQty As Decimal = -1
        Dim ED As Date
        Dim JmlDGV As Decimal = 0

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                If Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value Or _
                Me.TxtRemark.Enabled = False Then
                    Exit Sub
                End If

                'ahdar 
                If objCustomDO.ApakahEditDgvBenar_Custom(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, True) = False Then
                    Exit Sub
                End If

                'BufferQtyDefault = Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value
                BufferRatioReference = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                BufferRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                BufferRatioDesc = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                BufferLblQty = Me.LlbQty.Text

                frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                frmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value
                frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value

                If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                    Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value & _
                                                                            " " & Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value & " = " & _
                                                                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value & " " & _
                                                                            Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    If Me.txtPrimarySO.Text = Me.dgvItem.Item(Me.ColPrimarySO.Index, e.RowIndex).Value.ToString And Me.txtPrimaryCodeNo.Text = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value.ToString Then
                        Me.TxtRatioReference.Text = frmRatio.TxtRatioMulti.Text
                        Me.TxtRatioDefault.Text = frmRatio.TxtRatioDefault.Text
                        Me.LlbQty.Text = "Qty (" & Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value & ")"
                    End If
                    Me.SetQtyDefaultOnDgv(e.RowIndex)
                    If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Ratio restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = BufferRatioDesc
                        Me.LlbQty.Text = BufferLblQty
                        Me.SetQtyDefaultOnDgv(e.RowIndex)
                        Exit Sub
                    End If


                    ''cek batasan kebenaran data qty default
                    row = Me.dgvItem.Rows(e.RowIndex)
                    If IsDate(row.Cells(Me.ColExpiryDate.Index).Value) = True Then
                        ED = row.Cells(Me.ColExpiryDate.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    JmlDGV = row.Cells(Me.ColTotalQty.Index).Value
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                        SecQty = -1
                    Else
                        SecQty = row.Cells(Me.ColSecQty.Index).Value
                    End If
                    If Me.ApakahEditSah(row, row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColBatchNo.Index).Value, ED, SecQty, JmlDGV, True) = False Then
                        lanjut = False
                    End If

                    If lanjut = False Then
                        Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value = BufferQtyDefault
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioDesc
                        Me.LlbQty.Text = BufferLblQty
                    End If
                    SetGboItem()
                    Me.SetQtyDefaultOnDgv(e.RowIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim SecondaryQty, JmlDGV As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> Me.ColExpiryDate.Index Then
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        Else
                            If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True And _
                            IsNumeric(Me.buffer) = True Then
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = CDec(Me.buffer) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColRemarks.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(Me.ColExpiryDate.Index).Value) = True Then
                            ExpiryDate = CDate(.Cells(Me.ColExpiryDate.Index).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(Me.ColSecQty.Index).Value) = True Then
                            SecondaryQty = .Cells(Me.ColSecQty.Index).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                            Me.EndEditQty(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColSecQty.Index Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColBatchNo.Index Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        End If
                    End If
                Else
                    'If e.ColumnIndex = Me.ColRemarks.Index Or e.ColumnIndex = Me.ColBatchNo.Index Then
                    '    Dim BufferExpiryDate As String = ""
                    '    If Me.buffer <> "" Then
                    '        BufferExpiryDate = Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value
                    '    End If
                    '    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    '    Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value = CDate(#1/1/1950#)
                    '    If e.ColumnIndex = Me.ColBatchNo.Index Then
                    '        Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    '        If AdaSama = True Then
                    '            If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                    '            "Confirm") = MsgBoxResult.No Then
                    '                'kembalikan data batch no dan expiry date
                    '                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    '                Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value = BufferExpiryDate
                    '                Exit Sub
                    '            End If
                    '            Me.GabungkanBaris()
                    '        End If
                    '    End If
                    '    Me.FormatDGV()
                    'Else
                    '    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    'End If

                    If e.ColumnIndex = Me.ColRemarks.Index Then
                        If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    ElseIf e.ColumnIndex = Me.ColBatchNo.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
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
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim DefPrecMulti As Integer = 0
        Dim vPrimaryCodeNo As Decimal = row.Cells(Me.ColPrimaryCodeNo.Index).Value
        Dim vPrimarySO As Decimal = row.Cells(Me.ColPrimarySO.Index).Value
        Dim SecQty As Decimal = -1
        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
            SecQty = row.Cells(Me.ColSecQty.Index).Value
        Else
            SecQty = -1
        End If
        BatchNo = row.Cells(Me.ColBatchNo.Index).Value

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
                DefPrecMulti = ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefPrecMulti)

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                Dim ED As Date
                If IsDate(.Cells(Me.ColExpiryDate.Index).Value) = True Then
                    ED = .Cells(Me.ColExpiryDate.Index).Value
                Else
                    ED = #1/1/1950#
                End If

                Me.SetQtyDefaultOnDgv(e.RowIndex)
                If CDec(row.Cells(Me.ColTotalQty.Index).Value) <= 0 Then
                    MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(e.RowIndex)
                    Exit Sub
                End If

                JmlDGV = .Cells(Me.ColTotalQty.Index).Value
                If IsNumeric(.Cells(Me.ColSecQty.Index).Value) = False Then
                    SecQty = -1
                Else
                    SecQty = .Cells(Me.ColSecQty.Index).Value
                End If
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatchNo.Index).Value, ED, SecQty, JmlDGV, True) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(e.RowIndex)
                End If

                'jika jumlah menurun maka harus load ulang data
                Me.SetGboItem()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> -1 Then
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                'msgbox("belum")
                Dim ED As Date
                If IsDate(.Cells(Me.ColExpiryDate.Index).Value) = True Then
                    ED = .Cells(Me.ColExpiryDate.Index).Value
                Else
                    ED = #1/1/1950#
                End If
                JmlDGV = .Cells(Me.ColTotalQty.Index).Value
                Dim SecQty As Decimal
                If IsNumeric(.Cells(Me.ColSecQty.Index).Value) = False Then
                    SecQty = -1
                Else
                    SecQty = .Cells(Me.ColSecQty.Index).Value
                End If

                SecondaryQty = Me.buffer
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatchNo.Index).Value, ED, _
                SecQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'cek apakah ada kemungkinan merger
                Me.FormatDGV()
                Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                If AdaSama = True Then
                    If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                    "Confirm") = MsgBoxResult.No Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    Me.GabungkanBaris()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNo(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim BufferExpiryDate As String = row.Cells(Me.ColExpiryDate.Index).Value

        Try
            With row
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper.Trim
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If


                'tidak perlu cek berkaitan dengan jumlah order balance, karena tidak ada pengaruhnya pada 
                'order balance

                'cek berkaitan dengan jumlah digudang
                Dim ED As Date
                If IsDate(.Cells(Me.ColExpiryDate.Index).Value) = True Then
                    ED = .Cells(Me.ColExpiryDate.Index).Value
                Else
                    ED = #1/1/1950#
                End If
                JmlDGV = .Cells(Me.ColTotalQtyMulti.Index).Value
                Dim SecQty As Decimal
                If IsNumeric(.Cells(Me.ColSecQty.Index).Value) = False Then
                    SecQty = -1
                Else
                    SecQty = .Cells(Me.ColSecQty.Index).Value
                End If
                'hanya ketika ada batch no dan sec qty diisikan
                Dim BatchNoCont As Boolean = Me.ObjProduct.ApakahBatchNoControlled(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                Dim UseSecondary As Boolean = Me.ObjProduct.ApakahUseSecondary(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                Dim CekQty As Boolean = True
                If BatchNoCont = True And .Cells(Me.ColBatchNo.Index).Value = "" Then
                    CekQty = False
                End If
                If CekQty = True Then
                    If UseSecondary = True Then
                        If IsNumeric(.Cells(Me.ColSecQty.Index).Value) = True Then
                            If UseSecondary = True And .Cells(Me.ColSecQty.Index).Value = -1 Then
                                CekQty = False
                            End If
                        Else
                            CekQty = False
                        End If
                    End If
                End If
                If CekQty = True Then
                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatchNo.Index).Value, ED, SecQty, JmlDGV) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyDefaultOnDgv(row.Index)
                    End If
                End If
                'khusus bila product yang bersangkutan expiry date controlled juga maka
                'cek ke dalam basis data:
                '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                If IsDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value) = True Then
                    Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> CDate(Me.dgvItem(Me.ColExpiryDate.Index, e.RowIndex).Value) Then
                            If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy") + _
                            ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            Else
                                'update data expiry date juga
                                BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                            End If
                        End If
                    End If

                    'data grid view
                    Dim row1 As DataGridViewRow
                    For Each row1 In Me.dgvItem.Rows
                        If row1.Cells(Me.ColBatchNo.Index).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value _
                        And row1.Index <> e.RowIndex And row1.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value Then
                            If CDate(row1.Cells(Me.ColExpiryDate.Index).Value) <> CDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatchNo.Index, e.RowIndex).Value) + " is " + _
                                  Format(CDate(row1.Cells(Me.ColExpiryDate.Index).Value), "dd MMM yyyy") + _
                                   ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value = row1.Cells(Me.ColExpiryDate.Index).Value
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
                        Me.dgvItem.Item(Me.ColExpiryDate.Index, e.RowIndex).Value = Format(CDate(BufferExpiryDate), "dd MMM yyyy")
                        Exit Sub
                    End If
                    Me.GabungkanBaris()
                End If
            End With
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

                    If row.Cells(Me.ColCodeNo.Index).Value = row1.Cells(Me.ColCodeNo.Index).Value And row.Cells(Me.ColDescription.Index).Value = row1.Cells(Me.ColDescription.Index).Value And _
                    row.Cells(Me.ColSONO.Index).Value = row1.Cells(Me.ColSONO.Index).Value And row.Cells(Me.ColBatchNo.Index).Value = row1.Cells(Me.ColBatchNo.Index).Value And _
                    CStr(row.Cells(Me.ColQtyPerMaster.Index).Value) = CStr(row1.Cells(Me.ColQtyPerMaster.Index).Value) And row.Cells(Me.ColQty.Index).Value = _
                    row1.Cells(Me.ColQty.Index).Value Then
                        If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = True Then
                            If row.Cells(Me.ColSecQty.Index).Value = row1.Cells(Me.ColSecQty.Index).Value Then
                                Return True
                            Else
                                Return False
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False And IsNumeric(row1.Cells(Me.ColSecQty.Index).Value) = False Then
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
                    If CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
                    CStr(row.Cells(Me.ColCustCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCustCodeNo.Index).Value) And CStr(row.Cells(Me.ColCustDesc.Index).Value) = CStr(row1.Cells(Me.ColCustDesc.Index).Value) And _
                    CStr(row.Cells(Me.ColQtyPerMaster.Index).Value) = CStr(row1.Cells(Me.ColQtyPerMaster.Index).Value) And _
                    CStr(row.Cells(Me.ColQty.Index).Value) = CStr(row1.Cells(Me.ColQty.Index).Value) Then
                        If IsNumeric(row.Cells(Me.ColExpiryDate.Index).Value) = True And IsNumeric(row1.Cells(Me.ColExpiryDate.Index).Value) = True Then
                            If CDbl(row.Cells(Me.ColExpiryDate.Index).Value) = CDbl(row1.Cells(Me.ColExpiryDate.Index).Value) Then
                                Me.GabungkanBaris(row, row1)
                                Me.ObjCD.SetNumber(Me.dgvItem)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(Me.ColExpiryDate.Index).Value) = False And IsNumeric(row1.Cells(Me.ColExpiryDate.Index).Value) = False Then
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
            row.Cells(Me.ColTotalQtyMulti.Index).Value = CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDbl(row1.Cells(Me.ColTotalQtyMulti.Index).Value)
            'gabungkan data remarks
            If CStr(row.Cells(Me.ColRemarks.Index).Value).EndsWith(".") = False Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(Me.ColRemarks.Index).Value = CStr(row.Cells(Me.ColRemarks.Index).Value) + ". "
            End If
            row.Cells(Me.ColRemarks.Index).Value = CStr(row.Cells(Me.ColRemarks.Index).Value) + CStr(row1.Cells(Me.ColRemarks.Index).Value)
            'hapus baris ybs
            Me.dgvItem.Rows.RemoveAt(row1.Index)
            Me.SetQtyDefaultOnDgv(row.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, _
    ByVal PrimaryProduct As Decimal, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlObjLama, QtyOnHand, JmlSOB, JmlAcuan, _
        JmlSendiri, JmlBookingAll, JmlBookingSO, JmlPendingSendiri, _
        JmlPendingOther, JmlAvailable As Decimal
        JmlAcuan = Decimal.MaxValue

        Try
            JmlSendiri = row.Cells(Me.ColTotalQtyMulti.Index).Value
            'edit sah harus dilihat dari 2 sisi:
            '1. sisi jumlah sales order balance
            'jml sales order balance
            JmlSOB = _
            Me.ObjCD.DigitGrouping(Me.ObjSOItem.GetSalesOrderBalance_Multi(row.Cells(Me.ColPrimarySO.Index).Value, PrimaryProduct) _
            + Me.ObjDOItem.GetJml_Multi(Me.TxtPrimaryDO.Text, row.Cells(Me.ColPrimarySO.Index).Value, PrimaryProduct) _
            - Me.GetJmlDGV_Multi(PrimaryProduct, row.Cells(Me.ColPrimarySO.Index).Value))

            Dim ToleranceUp As Double = ObjSO.GetToleranceUp(row.Cells(ColPrimarySO.Index).Value, row.Cells(ColPrimaryCodeNo.Index).Value)
            JmlSOB += ToleranceUp

            If JmlSOB < 0 Then
                JmlAcuan = JmlSendiri + JmlSOB
            End If

            '2. sisi jumlah available
            'jml available dilihat dari sisi inventory, meperhitungkan juga jumlah dalam database
            If Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                If Me.ApakahBolehDOTanpaStock = False Then
                    QtyOnHand = Me.ObjProduct.GetJml4(PrimaryProduct, SecondaryQty, BatchNo)
                    JmlBookingAll = Me.ObjBookingInv.GetJml5(PrimaryProduct, SecondaryQty, BatchNo)
                    JmlBookingSO = Me.ObjBookingInv.GetJml4(PrimaryProduct, Me.txtPrimarySO.Text, SecondaryQty, BatchNo)
                    If Me.baru = False Then
                        JmlObjLama = Me.ObjDOItem.GetJml2(PrimaryProduct, Me.TxtPrimaryDO.Text, SecondaryQty, BatchNo)
                    Else
                        JmlObjLama = 0
                    End If
                    JmlDGV = Me.GetJmlDGV_Default(PrimaryProduct, BatchNo, SecondaryQty)
                    'Me.txtQTyFree.Text = Me.ObjCD.DigitGrouping(QtyOnHand - JmlBookingAll + _
                    'JmlBookingSO + JmlObjLama - JmlDGV)
                    JmlPendingSendiri = Me.ObjDOItem.GetJml3(Me.TxtPrimaryDO.Text, PrimaryProduct, SecondaryQty, BatchNo)
                    JmlPendingOther = Me.ObjDOItem.GetJmlPendingDO_Other(PrimaryProduct, BatchNo, SecondaryQty, Me.TxtPrimaryDO.Text)
                    JmlAvailable = Me.ObjCD.DigitGrouping(QtyOnHand - JmlBookingAll + _
                    JmlBookingSO - (JmlPendingSendiri - (JmlObjLama - JmlDGV)) - JmlPendingOther)

                    ''terjemahkan kedalam satuan multi
                    If CDec(row.Cells(Me.ColRatioDefault.Index).Value) = 0 Then
                        MsgBox("Logic error, please contact your software provider.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    Dim Ratio_ As Decimal = CDec(row.Cells(Me.ColRatioReference.Index).Value) / CDec(row.Cells(Me.ColRatioDefault.Index).Value)
                    JmlAvailable = JmlAvailable * Ratio_
                    JmlSendiri = JmlSendiri * Ratio_

                    If JmlAvailable < 0 Then
                        If JmlAcuan = 0 Then
                            JmlAcuan = JmlSendiri + JmlAvailable
                        Else
                            If JmlAcuan > JmlSendiri + JmlAvailable Then
                                JmlAcuan = JmlSendiri + JmlAvailable
                            End If
                        End If
                    End If
                End If
            End If


            'sampe sini 
            If JmlAcuan = Decimal.MaxValue Then
                Return True
            Else
                If MsgBox("The maximum value you can enter is " + Me.ObjCD.DigitGrouping(JmlAcuan, Me.ObjProduct.GetAngkaDecimal(PrimaryProduct, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)) + _
                ".", MsgBoxStyle.Information, "Information") Then
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'ObjSC = Nothing
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
                Me.txtCustomerCode.Enabled = True
                Me.btnFindCustomer.Enabled = True
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
            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)

            'load ulang data cbo product
            Me.SetGboItem()

            'format tampilan
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'tidak perlu cek akibatnya pada sales order balance
            'ketika user akan menghapus data, tidak perlu cek apakah akibatnya akan mengakibatkan ada stock yang minus, karena akan mengakibatkan
            'stock menjadi positif
            'tetapi untuk transaksi yang berkaitan dengan transaksi lainnya (DO, Invoice, GRN atau PO) harus cek validitas jumlah berkaitan dengan mereka
            'tandai apakah pada baris yang sama no SO-nya sama
            If Me.txtPrimarySO.Text = CStr(Me.dgvItem.Item(Me.ColPrimarySO.Index, e.Row.Index).Value) Then
                Me.SamaSO = True
            Else
                Me.SamaSO = False
            End If
            If Me.txtPrimaryCodeNo.Text = CStr(Me.dgvItem(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) Then
                Me.SamaCodeNo = True
            Else
                Me.SamaCodeNo = False
            End If
            Me.BufferSO = CStr(Me.dgvItem.Item(Me.ColPrimarySO.Index, e.Row.Index).Value)
            Me.bufferProduct = CStr(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Try
            Me.TxtRemark.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Try
            Me.TxtRemark.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.GotFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        Try
            Me.TxtQtyperBatchMulti.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboSecondaryQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSecondaryQty.GotFocus
        Try
            Me.cboSecondaryQty.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboSecondaryQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSecondaryQty.LostFocus
        Try
            Me.cboSecondaryQty.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        Try
            Me.TxtRemarkItem.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        Try
            Me.TxtRemarkItem.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNo.SelectedIndexChanged
        Try
            'isikan data expiry date 
            If Me.chkExpiryDateControlled.Checked = True Then
                If Me.cboBatchNo.Text <> "" Then
                    Me.DtpExpired.Value = Me.ObjProduct.GetExpiryDate(Me.txtPrimaryCodeNo.Text, _
                    Me.cboBatchNo.Text)
                    Me.DtpExpired.Visible = True
                    Me.lblExpiryDate.Visible = False
                Else
                    Me.DtpExpired.Visible = False
                    Me.lblExpiryDate.Visible = True
                    Me.DtpExpired.Value = #1/1/1950#
                End If
            Else
                Me.DtpExpired.Value = #1/1/1950#
            End If
            'isikan secondary unit
            If Me.chkUseSecondaryUnit.Checked = True And Me.cboBatchNo.SelectedIndex <> -1 Then
                Me.IsiCboSecondary(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                If Me.txtSecondaryQty.Text = "-1" Then
                    Me.SetJmlBO_N_FreeQty()
                End If
            Else
                Me.SetJmlBO_N_FreeQty()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboSecondaryQty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSecondaryQty.SelectedIndexChanged
        Try
            If Me.cboSecondaryQty.SelectedIndex <> -1 Then
                Me.txtSecondaryQty.Text = Me.cboSecondaryQty.Text
                Me.SetJmlBO_N_FreeQty()
            Else
                Me.txtSecondaryQty.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.TxtDONo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtDONo.Text <> "" Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_Cetak") = False Then
                    FrmO.AccessName = "FrmTransaksiDO_Cetak"
                    If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If

                Dim frmCetak As New FrmCetakDO
                frmCetak.Primary = Me.TxtPrimaryDO.Text
                frmCetak.UserCode = Me.UserCode
                frmCetak.ShowDialog()
                frmCetak = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO.Dispose()
        End Try
    End Sub

    Private Sub cboCustomerAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustomerAddress.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.cboCustomerAddress.SelectedIndex <> -1 Then
                    Me.txtCustomerAddIdentifier.Text = Me.ObjCustomer.GetIdentifierAdd(Me.txtPrimaryCustomer.Text, _
                    Me.cboCustomerAddress.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As _
    System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
            If Me.TxtRemark.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryDO.Text) AndAlso CDec(Me.TxtPrimaryDO.Text) > 0) Then
                If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                    MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.txtCustomerAddIdentifier.Text
                ij.PrimaryCust = Me.txtPrimaryCustomer.Text
                'If Me.TxtRemark.Enabled = False Then
                ij.cboCustAddCode.Enabled = False
                'End If
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False

                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.DefaultFillCustomer(False)
                    Me.cboCustomerAddress.SelectedIndex = 0
                    Me.chkKwitansi.Focus()
                    Me.SetGboItem()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.DefaultFillCustomer(False)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.chkKwitansi.Focus()
                        Me.SetGboItem()
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.DefaultFillCustomer(True)
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtCustomerCode.Focus()
                        Me.KosongkanItem()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Public Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    If PrimaryBaru <> Me.txtPrimaryCustomer.Text Then
                        Me.txtPrimaryCustomer.Text = PrimaryBaru
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.DefaultFillCustomer(False)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.chkKwitansi.Focus()
                        Me.SetGboItem()
                    End If
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.DefaultFillCustomer(False)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.chkKwitansi.Focus()
                        Me.SetGboItem()
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.DefaultFillCustomer(True)
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtCustomerCode.Focus()
                        Me.KosongkanItem()
                    End If
                End If
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.DefaultFillCustomer(True)
                Me.txtPrimaryCustomer.Text = "-1"
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.DefaultFillCustomer(False)
                Me.cboCustomerAddress.SelectedIndex = 0
                Me.chkKwitansi.Focus()
                Me.SetGboItem()
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.DefaultFillCustomer(True)
                Me.CboAddressShipTo.Items.Clear()
                Me.TxtIdentifierAddressShipTo.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtCustomerCode.Focus()
                Me.KosongkanItem()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Public Sub DefaultFillCustomer(ByVal vClear As Boolean)
        Try
            Me.Panel1.Visible = True
            Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
            If Me.CboAddressShipTo.Items.Count > 0 Then Me.CboAddressShipTo.SelectedIndex = 0
            Me.Panel2.Visible = False
            If vClear Then
                TxtIdentifierAddressShipTo.Text = ""
                CboAddressShipTo.Items.Clear()
                TxtPrimaryCustomerShipTo.Text = "-1"
                TxtCustomerCodeShipTo.Text = ""
                TxtCustomerDescShipTo.Text = ""
                TxtIdentifierAddressCustomerShipTo.Text = ""
                CboAddressCustomerShipTo.Items.Clear()
            End If

            Me.Label6.Location = New Point(Me.Label6.Location.X, Me.Label1.Location.Y)
            Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.TxtCustomerName.Location.Y + 5)
            Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, (Me.TxtCustomerName.Size.Height * 2) + (Me.cboCustomerAddress.Top - Me.TxtCustomerName.Bottom))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub cboSalesOrderNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSalesOrderNo.SelectedValueChanged
        Try
            'hanya jalan, jika perubahan bukan diinitiate dari cbo po
            If Me.SelectedCBOSOChange = True Then
                Me.SelectedCBOPOChange = False

                If Me.cboSalesOrderNo.SelectedIndex <> -1 Then
                    Dim RowView As DataRowView = Me.cboSalesOrderNo.Items(Me.cboSalesOrderNo.SelectedIndex)
                    Dim row2 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow = RowView.Row
                    Dim PrimarySOBaru As Decimal = row2.Primary_SO
                    If Me.txtPrimarySO.Text <> PrimarySOBaru Then
                        Me.txtPrimarySO.Text = PrimarySOBaru
                        Me.IsiCboProduct(Me.txtPrimarySO.Text)
                    End If

                    Me.cboCustPONO.DataSource = Me.TBL.DefaultView
                    Me.cboCustPONO.ValueMember = "Cust PO No"
                    Me.cboCustPONO.DisplayMember = "Cust PO No"
                    Me.cboSalesOrderNo.ValueMember = "SO No"
                    Me.cboSalesOrderNo.DisplayMember = "SO No"
                    Me.cboCustPONO.SelectedIndex = Me.cboSalesOrderNo.SelectedIndex

                    If IsNumeric(Me.txtPrimaryCodeNo.Text) = True Then
                        Me.TxtRemarkItem.Text = Me.ObjSOItem.GetRemarks(PrimarySOBaru, Me.txtPrimaryCodeNo.Text)
                    End If
                Else
                    Me.txtPrimarySO.Text = "-1"
                    Me.cboCustPONO.DataSource = Nothing
                    Me.cboCustPONO.Items.Clear()
                    Me.TxtRemarkItem.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SelectedCBOPOChange = True
        End Try
    End Sub

    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Try
            If Me.cboCodeNo.SelectedIndex <> -1 Then
                Dim RowView As DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
                Dim row2 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow = RowView.Row
                Dim PrimaryCodeNoBaru As Double = row2.Primary_Product

                If Me.txtPrimaryCodeNo.Text <> PrimaryCodeNoBaru Then
                    Me.txtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.TxtRemarkItem.Text = Me.ObjSOItem.GetRemarks(Me.txtPrimarySO.Text, PrimaryCodeNoBaru)

                    'cek fitur apa saja yang dikontrol
                    Me.ObjProduct.IsiCheckBoxControll(Me.txtPrimaryCodeNo.Text, Me.chkLocationControlled, _
                    Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
                    Me.chkPdoControlled, Me.txtPrimaryUnitsDefault, Me.txtPrimarySecondaryUnits)

                    'location
                    If Me.chkLocationControlled.Checked = True Then
                        'If Me.baru = True And Me.GetJml_Default(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text) = 0 Then
                        If Me.ApakahBolehDOTanpaStock = False Then
                            'jika ada 1 baris pada dgv yang menggunakan primary product yg sama tapi primary locationnya -1, maka tidak
                            'bisa menggunakan fitur ini, karena di fungsi fill item, ketika akan kurangi dengan dgv, bingung
                            If Me.GetJml_PrimaryLocation_Minus(Me.txtPrimaryCodeNo.Text) = 0 Then
                                If baru = True Then
                                    Me.llbAddByLocation.Visible = True
                                    Me.txtDesc.Width = Me.llbAddByLocation.Left - 5 - Me.txtDesc.Left
                                Else
                                    'pastikan item product pada do tsb primary locationnya <> -1
                                    If Me.ObjDOItem.GetJmlWithOutLocation(Me.TxtPrimaryDO.Text, Me.txtPrimaryCodeNo.Text) = 0 Then
                                        Me.llbAddByLocation.Visible = True
                                        Me.txtDesc.Width = Me.llbAddByLocation.Left - 5 - Me.txtDesc.Left
                                    Else
                                        Me.llbAddByLocation.Visible = False
                                        Me.txtDesc.Width = Me.SplitContainer2.Panel1.Width - Me.txtDesc.Left
                                    End If
                                End If
                            Else
                                Me.llbAddByLocation.Visible = False
                                Me.txtDesc.Width = Me.SplitContainer2.Panel1.Width - Me.txtDesc.Left
                            End If
                        Else
                            Me.llbAddByLocation.Visible = False
                            Me.txtDesc.Width = Me.SplitContainer2.Panel1.Width - Me.txtDesc.Left
                        End If
                        'Else
                        '    Me.llbAddByLocation.Visible = False
                        '    Me.txtDesc.Width = Me.SplitContainer2.Panel1.Width - Me.txtDesc.Left
                        'End If
                    Else
                        Me.llbAddByLocation.Visible = False
                        Me.txtDesc.Width = Me.SplitContainer2.Panel1.Width - Me.txtDesc.Left
                    End If

                    'batch no
                    If Me.chkBatchNoControlled.Checked = True Then
                        Me.cboBatchNo.TabStop = True
                        Me.lblBatchNo.Visible = False
                        Me.cboBatchNo.Visible = True
                        'isikan data batch no dari basis data
                        Me.IsiCboBatchNo(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
                    Else
                        Me.cboBatchNo.TabStop = False
                        Me.lblBatchNo.Visible = True
                        Me.cboBatchNo.Visible = False
                        Me.cboBatchNo.DataSource = Nothing
                        Me.cboBatchNo.Items.Clear()
                    End If

                    'expiry date
                    If Me.chkExpiryDateControlled.Checked = True Then
                        Me.lblExpiryDate.Visible = False
                        Me.DtpExpired.Visible = True
                    Else
                        Me.lblExpiryDate.Visible = True
                        Me.DtpExpired.Visible = False
                        Me.DtpExpired.Value = #1/1/1950#
                    End If

                    'isikan secondary unit
                    If Me.chkUseSecondaryUnit.Checked = False Then
                        Me.lblSecondaryUnits.Visible = False
                        'Me.lblUnitsMulti.Left = Me.lblSecondaryUnits.Left
                        'Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
                        Me.cboSecondaryQty.Text = "-1"
                        Me.lblSecQty.Visible = False
                        Me.cboSecondaryQty.Visible = False
                        Me.cboSecondaryQty.DataSource = Nothing
                        Me.cboSecondaryQty.Items.Clear()
                    Else
                        Me.lblSecondaryUnits.Visible = True
                        Me.lblSecQty.Visible = True
                        'Me.lblUnitsMulti.Left = Me.lblSecQty.Left - 37
                        'Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
                        Me.cboSecondaryQty.Visible = True
                        If Me.chkBatchNoControlled.Checked = False Then
                            'isikan data secondary jika perlu
                            Me.IsiCboSecondary(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
                        End If
                    End If

                    '----------------------------------------------------------------------------------------------------------------------------------------
                    Me.txtPrimaryUnitsMulti.Text = row2.Primary_Units_Multi
                    Me.lblUnitsMulti.Text = row2.Units_Code_Multi
                    Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                    Dim DefPrecDefault As Integer = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
                    Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping("0", DefPrecMulti)

                    Me.txtPrimaryUnitsDefault.Text = row2.Primary_Units_Default
                    Me.lblUnitsDefault.Text = row2.Units_Code_Default
                    Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping("0", DefPrecDefault)

                    Me.TxtRatioReference.Text = Me.ObjCD.DigitGrouping(row2.Ratio_Reference, DefPrecMulti)
                    Me.TxtRatioDefault.Text = Me.ObjCD.DigitGrouping(row2.Ratio_Default, DefPrecDefault)

                    Dim idx As Integer = -1
                    idx = Me.GetIndexRatio(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)
                    If Me.txtPrimaryUnitsMulti.Text <> Me.txtPrimaryUnitsDefault.Text Then
                        If idx <> -1 Then
                            Me.LlbQty.Text = "Qty (" & Me.dgvItem.Item(Me.ColRatioDesc.Index, idx).Value & ")"
                            Me.TxtRatioReference.Text = Me.dgvItem.Item(Me.ColRatioReference.Index, idx).Value
                            Me.TxtRatioDefault.Text = Me.dgvItem.Item(Me.ColRatioDefault.Index, idx).Value
                        Else
                            Me.LlbQty.Text = "Qty (" & Me.ObjCD.DigitGrouping(row2.Ratio_Reference, DefPrecMulti) & _
                            " " & row2.Units_Code_Multi & " = " & Me.ObjCD.DigitGrouping(row2.Ratio_Default, DefPrecDefault) & _
                            " " & row2.Units_Code_Default & ")"
                        End If
                        Me.TxtQtyperBatchDefault.Visible = True
                        Me.lblUnitsDefault.Visible = True
                    Else
                        Me.LlbQty.Text = "Qty"
                        Me.TxtQtyperBatchDefault.Visible = False
                        Me.lblUnitsDefault.Visible = False
                    End If

                    If Me.LlbQty.Left + Me.LlbQty.Width + 2 > Me.TxtRemarkItem.Left Then
                        Me.TxtQtyperBatchMulti.Left = Me.LlbQty.Left + Me.LlbQty.Width + 2
                    Else
                        Me.TxtQtyperBatchMulti.Left = Me.TxtRemarkItem.Left
                    End If
                    '----------------------------------------------------------------------------------------------------------------------------------------

                End If
            Else
                'reset tampilan
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Me.txtPrimaryUnitsDefault.Text = "-1"
                Me.txtPrimaryUnitsMulti.Text = "-1"
                Me.txtPrimarySecondaryUnits.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.TxtQtyperBatchMulti.Text = "0"
                Me.TxtQtyperBatchDefault.Text = "0"
                Me.txtAvailable.Text = "0"
                Me.txtAvailableMulti.Text = "0"
                Me.cboSecondaryQty.Text = "-1"
                Me.txtQTyFree.Text = "0"
                Me.txtQtyOrderBalance.Text = "0"
                Me.chkBatchNoControlled.Checked = False
                Me.chkExpiryDateControlled.Checked = False
                Me.chkUseSecondaryUnit.Checked = False
                Me.lblSecondaryUnits.Visible = False
                'Me.lblUnitsMulti.Left = Me.lblSecondaryUnits.Left
                'Me.TxtQtyperBatchMulti.Width = Me.lblUnitsMulti.Left - Me.TxtQtyperBatchMulti.Left - 6
                Me.lblSecQty.Visible = False
                Me.lblExpiryDate.Visible = False
                Me.DtpExpired.Visible = True
                Me.DtpExpired.Value = Today
                Me.lblBatchNo.Visible = False
                Me.cboBatchNo.Visible = True
                Me.cboBatchNo.TabStop = True
                Me.cboBatchNo.DataSource = Nothing
                Me.cboBatchNo.Items.Clear()
                Me.cboSecondaryQty.Visible = False
                Me.TxtRemarkItem.Text = ""

                Me.TxtRatioReference.Text = "1"
                Me.TxtRatioDefault.Text = "1"
                Me.TxtQtyperBatchMulti.Left = Me.TxtRemarkItem.Left
            End If

            'format tampilan
            If Me.txtPrimaryCodeNo.Text <> "-1" Then
                'Me.lblUnitsMulti.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.txtPrimaryCodeNo.Text)
                End If
                Me.lblUnitFree.Text = Me.lblUnitsMulti.Text
                Me.LblUnitOB.Text = Me.lblUnitsMulti.Text
                Me.lblUnitsPendingDO.Text = Me.lblUnitsMulti.Text
                Me.lblAvailableUnits.Text = Me.lblUnitsDefault.Text
                Me.lblAvailableUnitsMulti.Text = Me.lblUnitsMulti.Text
                Me.SetJmlBO_N_FreeQty()
            Else
                Me.txtQtyOrderBalance.Text = "0"
                Me.TxtToleranceUp.Text = "0"
                Me.LblUnitOB.Text = "Units"

                Me.txtPendingDO.Text = "0"
                Me.lblUnitsPendingDO.Text = "Units"

                Me.txtAvailableMulti.Text = "0"
                Me.lblAvailableUnitsMulti.Text = "Units"

                Me.txtAvailableMulti.Text = "0"
                Me.lblAvailableUnits.Text = "Units"

                Me.LlbQty.Text = "Qty"

                Me.TxtQtyperBatchMulti.Text = "0"
                Me.lblUnitsMulti.Text = "Units"

                Me.TxtQtyperBatchDefault.Text = "0"
                Me.lblUnitsDefault.Text = "Units"

            End If
            Me.TxtQtyperBatchMulti.Width = Me.txtAvailableMulti.Right - Me.TxtQtyperBatchMulti.Left
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbDetailStock_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbDetailStock.LinkClicked
        Dim i As New Inventory_Module.FrmMasterKondisiStock

        Try
            If Me.cboCodeNo.Text <> "" Then
                i.txtCodeNo.Text = Me.cboCodeNo.Text
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Function GetJmlDO(ByVal TBL As DseTransaksiDO.tbltrans_do_itemDataTable, _
    ByVal vPrimaryProduct As Decimal, ByVal vPrimarySO As Decimal) As Decimal
        Dim cmd As String = "[Primary Product] = " + CStr(vPrimaryProduct) + " AND [Primary SO] =" + _
        CStr(vPrimarySO)
        Dim row As DseTransaksiDO.tbltrans_do_itemRow

        Try
            TBL.DefaultView.RowFilter = cmd
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView.Item(0).Row
                Return row.Total_Qty
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' menambahkan langsung dari booking, hanya bisa untuk item baru
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub llbAddBooking_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddBooking.LinkClicked
        Dim DTABooking As New DseTransaksiBookingInvTableAdapters.tbltrans_booking_invitem_view2TableAdapter
        Dim TBLBooking As New DseTransaksiBookingInv.tbltrans_booking_invitem_view2DataTable
        Dim row As DseTransaksiBookingInv.tbltrans_booking_invitem_view2Row
        Dim rowDGV As DataGridViewRow
        Dim JmlDGV, JmlQuickAdd, JmlAvailable As Decimal
        Dim Index As Integer = 0
        Dim mPrimaryCurrencySO, mPersenPajak As Decimal
        Dim mJenisFakturPajak As ClsTransaksiSalesOrder.EnumJenisFakturPajak = _
        ClsTransaksiSalesOrder.EnumJenisFakturPajak.FakturPajakSederhana

        Try
            'isikan data dari database
            DTABooking.FillBy_PrimarySO(TBLBooking, Me.txtPrimarySO.Text)
            If TBLBooking.Rows.Count > 0 Then
                If Me.dgvItem.RowCount > 0 Then
                    'samakan dengan referensinya, yang sudah dimasukkan ke dalam form
                    Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.txtPrimarySO.Text, mPrimaryCurrencySO, mJenisFakturPajak, mPersenPajak)
                    '   a. % pajak yang digunakan sama
                    If mPersenPajak <> Me.PersenPajakAcuan Then
                        MsgBox("This sales order cannot be in one delivery order with the existing items. " + _
                        "For several sales order item written in one delivery order their currency, tax, and tax invoice must be identical", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    '   b. jenis faktur pajak yang digunakan sama
                    If mJenisFakturPajak <> Me.JenisFakturPajakAcuan Then
                        MsgBox("This sales order cannot be in one delivery order with the existing items. " + _
                        "For several sales order item written in one delivery order their currency, tax, and tax invoice must be identical", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    '   c. primary currency sama
                    If mPrimaryCurrencySO <> Me.PrimaryCurrencySO Then
                        MsgBox("This sales order cannot be in one delivery order with the existing items. " + _
                        "For several sales order item written in one delivery order their currency, tax, and tax invoice must be identical", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            End If
            'masukan ke dalam dgv
            For Each row In TBLBooking.Rows
                'quick add ini hanya bisa digunakan untuk item yang bukan location controlled
                If row.Location_Controlled = 0 Then
                    'hitung jumlah yang sudah ada di dalam dgv untuk booking ybs
                    JmlDGV = Me.GetJmlDGV_Multi(row.Primary_Product, row.Batch_No, row.Sec_Qty, row.Primary_SO)
                    'cek batasan jumlah berdasarkan jumlah barang yang sudah dibooking
                    If JmlDGV < row.Qty Then
                        'hitung jumlah yang bisa dimasukkan
                        JmlQuickAdd = row.Qty - JmlDGV
                        'hitung ketersedian dalam inventory
                        JmlAvailable = Me.GetJmlAvailable(row.Primary_Product, row.Primary_SO, _
                        row.Batch_No, row.Sec_Qty)
                        If JmlAvailable > 0 Then
                            'set nilai acuan untuk pengecekan SO
                            Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.txtPrimarySO.Text, Me.PrimaryCurrencySO, _
                            Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
                            If Me.dgvItem.RowCount = 0 Then
                                'set nilai acuan untuk pengecekan SO
                                Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(row.Primary_SO, Me.PrimaryCurrencySO, _
                                Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
                            End If
                            'jika yang ada digudang < yang dibooking maka yang dianggap ada adalah yang digudang
                            If JmlAvailable < JmlQuickAdd Then
                                JmlQuickAdd = JmlAvailable
                            End If
                            Index = Me.GetIndex2(row.Primary_Product, row.Batch_No, _
                            row.Sec_Qty, row.Expiry_Date, row.Primary_SO)
                            If Index <> -1 Then
                                'tambhakan qty dalam dgv sejumlah jml quick add
                                rowDGV = Me.dgvItem.Rows(Index)
                                rowDGV.Cells(Me.ColTotalQty.Index).Value = _
                                Me.ObjCD.DigitGrouping(CDbl(rowDGV.Cells(Me.ColTotalQty.Index).Value) _
                                + row.Qty)
                            Else
                                'tambahkan baris baru
                                Me.dgvItem.Rows.Add()
                                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                                rowDGV.Cells(Me.ColNo.Index).Value = ""
                                rowDGV.Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                                rowDGV.Cells(Me.ColDescription.Index).Value = row.Description
                                rowDGV.Cells(Me.ColCustCodeNo.Index).Value = ""
                                rowDGV.Cells(Me.ColCustDesc.Index).Value = ""
                                rowDGV.Cells(Me.ColSONO.Index).Value = row.SO_Number
                                rowDGV.Cells(Me.ColBatchNo.Index).Value = row.Batch_No
                                rowDGV.Cells(Me.ColExpiryDate.Index).Value = row.Expiry_Date
                                rowDGV.Cells(Me.ColSecQty.Index).Value = row.Sec_Qty
                                rowDGV.Cells(Me.ColSecUnit.Index).Value = row.Sec_Unit
                                rowDGV.Cells(Me.ColQty.Index).Value = row.Qty
                                rowDGV.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(row.Qty, row._Decimal)
                                rowDGV.Cells(Me.ColUnit.Index).Value = row.Unit_Code
                                rowDGV.Cells(Me.ColMasterUnits.Index).Value = ""
                                rowDGV.Cells(Me.ColQtyMaster.Index).Value = 0
                                rowDGV.Cells(Me.ColQtyPerMaster.Index).Value = 1
                                rowDGV.Cells(Me.ColRemarks.Index).Value = ""
                                rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                                rowDGV.Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                                rowDGV.Cells(Me.ColPrimarySecUnit.Index).Value = row.Primary_Sec_Unit
                                rowDGV.Cells(Me.ColPrimarySO.Index).Value = row.Primary_SO
                                rowDGV.Cells(Me.ColPrimaryUnit.Index).Value = row.Primary_Unit
                            End If
                        End If
                    End If
                End If
            Next

            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.FormatDGV()
            'load ulang data
            Me.cboSalesOrderNo.DataSource = Nothing
            Me.IsiCboSalesOrder(Me.txtPrimaryCustomer.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTABooking.Dispose()
            TBLBooking.Dispose()
        End Try
    End Sub

    Private Sub txtSecondaryQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.TextChanged
        Try
            If Me.txtSecondaryQty.Text = "0" Then
                MsgBox("")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = False Then
                    Me.txtCustomerCode.BackColor = Me.BackColorAll
                Else
                    Me.txtCustomerCode.BackColor = Me.TxtCustomerName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = False Then
                    Me.txtCustomerCode.BackColor = Me.txtPrimaryCategory.BackColor
                Else
                    Me.txtCustomerCode.BackColor = Me.TxtCustomerName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then
                    Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.SetQtyDefault()
                    If CDec(Me.TxtQtyperBatchDefault.Text) = 0 And CDec(Me.TxtQtyperBatchMulti.Text) > 0 Then
                        Me.TxtQtyperBatchMulti.Text = Me.buffer
                        MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.SetQtyDefault()
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboSalesOrderNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSalesOrderNo.SelectedIndexChanged

    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                    If Me.baru = False Then
                        If Me.objDO.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objDO.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                            'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                            Me.TxtDONo.Text = Me.objDO.NoTransaksi
                        Else
                            'jika pindah bulan maka akan ganti no
                            Me.TxtDONo.Text = Me.objDO.GetNoBaru1(Me.dtpDate.Value)
                        End If
                    Else
                        Me.TxtDONo.Text = Me.objDO.GetNoBaru1(Me.dtpDate.Value)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDONo.GotFocus
        Try
            If Me.TxtRemark.Enabled = True And Me.TxtDONo.Enabled = True Then
                If Me.TxtDONo.ReadOnly = True Then
                    Me.TxtDONo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDONo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDONo.LostFocus
        Try
            If Me.TxtRemark.Enabled = True And Me.TxtDONo.Enabled = True Then
                If Me.TxtDONo.ReadOnly = True Then
                    Me.TxtDONo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDONo.BackColor = Me.TxtPrimaryDO.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCustPONO_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCustPONO.SelectedValueChanged
        Try
            'hanya jalan, jika perubahan bukan diinitiate dari cbo po
            If Me.SelectedCBOPOChange = True Then
                Me.SelectedCBOSOChange = False

                If Me.cboCustPONO.SelectedIndex <> -1 Then
                    Dim RowView As DataRowView = Me.cboSalesOrderNo.Items(Me.cboCustPONO.SelectedIndex)
                    Dim row2 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow = RowView.Row
                    Dim PrimarySOBaru As Decimal = row2.Primary_SO
                    If Me.txtPrimarySO.Text <> PrimarySOBaru Then
                        Me.txtPrimarySO.Text = PrimarySOBaru
                        Me.IsiCboProduct(Me.txtPrimarySO.Text)
                    End If

                    Me.cboSalesOrderNo.DataSource = Me.TBL.DefaultView
                    Me.cboSalesOrderNo.ValueMember = "SO No"
                    Me.cboSalesOrderNo.DisplayMember = "SO No"
                    Me.cboCustPONO.ValueMember = "SO No"
                    Me.cboCustPONO.DisplayMember = "Cust PO No"

                    Me.cboSalesOrderNo.SelectedIndex = Me.cboCustPONO.SelectedIndex

                    If IsNumeric(Me.txtPrimaryCodeNo.Text) = True Then
                        Me.TxtRemarkItem.Text = Me.ObjSOItem.GetRemarks(PrimarySOBaru, Me.txtPrimaryCodeNo.Text)
                    End If
                Else
                    Me.txtPrimarySO.Text = "-1"
                    Me.cboSalesOrderNo.DataSource = Nothing
                    Me.cboSalesOrderNo.Items.Clear()
                    Me.TxtRemarkItem.Text = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SelectedCBOSOChange = True
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.Timer1.Enabled = False
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_New") = False Then
                Me.btnNew.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexRatio(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Integer
        Try
            For Each rowDgv As DataGridViewRow In Me.dgvItem.Rows
                With rowDgv
                    If .Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And .Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                        .Cells(Me.ColRatioDesc.Index).Value <> "-" Then
                        Return rowDgv.Index
                    End If
                End With
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub txtQtyOrderBalance_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyOrderBalance.TextChanged
        Try
            If Me.txtQtyOrderBalance.Text = "0" Then
                Me.TxtToleranceUp.Text = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkIgnoreTolerance_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkIgnoreTolerance.CheckedChanged
        Try
            If Me.txtPrimaryCustomer.Text <> "" And Me.txtPrimaryCustomer.Text <> "-1" And Me.TxtRemark.Enabled = True Then
                Me.SetGboItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFieldToleranceUp()
        Dim DefPrec As Integer
        Try
            DefPrec = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            Me.TxtToleranceUp.Text = ObjCD.DigitGrouping(Me.ObjSO.GetToleranceUp(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text), DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbQty_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbQty.LinkClicked
        Dim frmRatio As New FrmTransaksiSO_SetQtyRatio
        Try
            If Me.TxtRemark.Enabled = False Or Me.LlbQty.Text = "Qty" Then
                Exit Sub
            End If
            If Me.GetIndexRatio(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text) <> -1 Then
                MsgBox("There is already ratio defined below. You cannot set different ratio for exist product.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            frmRatio.vRatioMulti = Me.TxtRatioReference.Text
            frmRatio.vUnitMulti = Me.lblUnitsMulti.Text
            frmRatio.vRatioDefault = Me.TxtRatioDefault.Text
            frmRatio.vUnitDefault = Me.lblUnitsDefault.Text
            If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtRatioReference.Text = frmRatio.TxtRatioMulti.Text
                Me.TxtRatioDefault.Text = frmRatio.TxtRatioDefault.Text
                Me.LlbQty.Text = "Qty (" & frmRatio.TxtRatioMulti.Text & " " & frmRatio.LlbUnitMulti.Text & " = " & frmRatio.TxtRatioDefault.Text & " " & frmRatio.LlbUnitDefault.Text & ")"
                SetQtyDefault()
                If Me.LlbQty.Left + Me.LlbQty.Width + 2 > Me.TxtRemarkItem.Left Then
                    Me.TxtQtyperBatchMulti.Left = Me.LlbQty.Left + Me.LlbQty.Width + 2
                Else
                    Me.TxtQtyperBatchMulti.Left = Me.TxtRemarkItem.Left
                End If
                Me.TxtQtyperBatchMulti.Width = Me.txtAvailableMulti.Right - Me.TxtQtyperBatchMulti.Left
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.buffer = Me.TxtQtyperBatchMulti.Text
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Enter
        Try
            Me.buffer = Me.TxtQtyperBatchDefault.Text
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.GotFocus
        Try
            Me.TxtQtyperBatchDefault.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub TxtQtyperBatchDefault_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub TxtQtyperBatchDefault_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Leave
        Try
            Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
            If Me.txtPrimaryCodeNo.Text <> "-1" And Me.TxtRatioDefault.Text <> "" And Me.TxtRatioReference.Text <> "" And CDec(Me.buffer) <> CDec(Me.TxtQtyperBatchDefault.Text) Then
                SetQtyMulti()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyMulti()
        Try
            If IsNumeric(Me.TxtQtyperBatchDefault.Text) = True And IsNumeric(Me.TxtRatioDefault.Text) = True Then
                Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchDefault.Text) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text), ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchDefault_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
        Try
            Me.TxtQtyperBatchDefault.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgv(ByVal vRowIdx As Integer)
        Try
            Me.dgvItem.Item(Me.ColTotalQty.Index, vRowIdx).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColTotalQtyMulti.Index, vRowIdx).Value) * _
            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value / _
            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value)), _
            Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_DeliveryOrder_Create(ByVal vPrimaryTransaksi As Double, ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core, _
    ByVal vTblPrimary As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, ByVal vTransItemDO As Boolean)
        Dim TblTemp As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.objDO.ResetObject(Me.dgvItem)
            Me.objDO.CariDetail(vPrimaryTransaksi, Me.dgvItem)

            'Update Primary SO
            If vTransItemDO = True Then
                For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                    TblTemp.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value)
                    If TblTemp.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), vTblPrimary)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), PrimaryUpdate, TblTemp)
                        rowD.Cells(Me.ColPrimarySO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimarySO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), TblTemp)
                    End If
                Next
            Else
                Me.dgvItem.Rows.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_DeliveryOrder_Save() As Double
        Try
            Me.baru = True
            Me.objDO.SimpanBaru1(Me.dgvItem, Me.UserCode, Me, False, False, False, True)
            Return Me.objDO.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function


    Private Sub llbAddByLocation_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddByLocation.LinkClicked
        Dim RowDgv, RowDgvDelivery As DataGridViewRow
        Dim ExpiredDateStr As String = ""
        Dim IdxTotal, IdxInput As Integer
        Dim ij As New FrmTransaksiDO3_PickQty
        Dim JmlInputMulti As Decimal = 0
        Dim JmlTotalInputMulti As Decimal = 0
        Dim DeliveryDate As Date = #1/1/1950#
        Dim frmoptionDO As New FrmTransaksiDO3_Option
        Dim TxtSOLama As String = ""
        Dim IdxAcuan As Integer = -1

        Try
            IdxTotal = 0
            IdxInput = 0

            ij.PrimarySO = Me.txtPrimarySO.Text
            ij.PrimaryProduct = Me.txtPrimaryCodeNo.Text
            ij.PrimaryUnitDefault = Me.txtPrimaryUnitsDefault.Text
            ij.PrimaryUnitMulti = Me.txtPrimaryUnitsMulti.Text
            ij.DefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            ij.DefPrecMulti = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            ij.FrmRef = Me
            ij.vDGVDOItem = Me.dgvItem
            ij.UserCode = Me.UserCode
            ij.PrimaryDO = Me.TxtPrimaryDO.Text

            ij.UnitDefault = Me.lblUnitsDefault.Text
            ij.UnitMulti = Me.lblUnitsMulti.Text
            ij.SONumber = Me.cboSalesOrderNo.Text
            ij.CodeNo = Me.cboCodeNo.Text
            ij.Desc = Me.txtDesc.Text
            ij.BatchNoCOntrol = Me.chkBatchNoControlled.Checked
            ij.ExpiryControl = Me.chkExpiryDateControlled.Checked

            ij.RatioMulti = Me.TxtRatioReference.Text
            ij.RatioDefault = Me.TxtRatioDefault.Text
            ij.RatioDesc = Me.ObjCD.DigitGrouping(ij.RatioMulti, ij.DefPrecMulti) + " " + ij.UnitMulti + " = " + Me.ObjCD.DigitGrouping(ij.RatioDefault, ij.DefPrecDefault) + " " + ij.UnitDefault
            ij.QtyOBMulti = Me.txtQtyOrderBalance.Text
            If ij.PrimaryUnitMulti = ij.PrimaryUnitDefault Then
                ij.QtyOBDefault = ij.QtyOBMulti
            Else
                ij.QtyOBDefault = Math.Round((CDec(Me.txtQtyOrderBalance.Text) * ij.RatioDefault) / ij.RatioMulti, ij.DefPrecDefault, MidpointRounding.AwayFromZero)
            End If
            ij.QtyAvailableDefault_Ref = Me.txtAvailable.Text
            ij.QtyAvailableMulti_Ref = Me.txtAvailableMulti.Text

            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtQtyperBatchDefault.Text = ij.TxtQtyDeliveryDefault.Text
                Me.TxtQtyperBatchMulti.Text = ij.TxtQtyDeliveryMulti.Text
                If Me.ApakahDataItemBenar(False) = True Then
                    If Me.ShowDeliveryDate = True Then
                        DeliveryDate = Me.GetDeliveryDate(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)

                        If DeliveryDate = #1/1/1950# Then
                            frmoptionDO.PrimarySO = Me.txtPrimarySO.Text
                            frmoptionDO.PrimaryProd = Me.txtPrimaryCodeNo.Text
                            frmoptionDO.ShowDialog()
                            If frmoptionDO.DialogResult = Windows.Forms.DialogResult.Cancel Then
                                Exit Sub
                            Else
                                If frmoptionDO.cbo_DeliveryDate.Text = "" Then
                                    DeliveryDate = #1/1/1950#
                                Else
                                    DeliveryDate = frmoptionDO.cbo_DeliveryDate.Text
                                End If
                            End If
                        End If
                    End If

                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                    For Each RowDgvDelivery In ij.dgvItemDetail.Rows
                        If CDec(RowDgvDelivery.Cells(ij.ColQtyDelivery.Index).Value) > 0 Then
                            IdxTotal += 1
                        End If
                    Next
                    JmlTotalInputMulti = CDec(ij.TxtQtyDeliveryMulti.Text)

                    For Each RowDgvDelivery In ij.dgvItemDetail.Rows
                        If RowDgvDelivery.Cells(ij.ColQtyDelivery.Index).Value > 0 Then
                            IdxInput += 1

                            IdxAcuan = Me.GetIndex3(Me.txtPrimaryCodeNo.Text, RowDgvDelivery.Cells(ij.ColBatchNo.Index).Value, -1, _
                            CDate(RowDgvDelivery.Cells(ij.ColExpiry.Index).Value), Me.txtPrimarySO.Text, RowDgvDelivery.Cells(ij.ColPrimaryLocation.Index).Value)
                            If IdxAcuan = -1 Then
                                Me.dgvItem.Rows.Add()
                                RowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                                With RowDgv
                                    .Cells(Me.ColPrimaryLocation.Index).Value = RowDgvDelivery.Cells(ij.ColPrimaryLocation.Index).Value
                                    .Cells(Me.ColLocation.Index).Value = RowDgvDelivery.Cells(ij.ColLocation.Index).Value
                                    .Cells(Me.ColWarehouse.Index).Value = RowDgvDelivery.Cells(ij.ColWarehouse.Index).Value

                                    .Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                                    .Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                                    .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                                    .Cells(Me.ColCustCodeNo.Index).Value = Me.cboCustCodeNo.Text
                                    .Cells(Me.ColCustDesc.Index).Value = Me.txtCustProductDesc.Text
                                    .Cells(Me.ColSONO.Index).Value = Me.cboSalesOrderNo.Text
                                    .Cells(Me.ColBatchNo.Index).Value = RowDgvDelivery.Cells(ij.ColBatchNo.Index).Value
                                    If RowDgvDelivery.Cells(ij.ColExpiry.Index).Value = #1/1/1950# Then
                                        ExpiredDateStr = "-"
                                    Else
                                        ExpiredDateStr = Format(CDate(RowDgvDelivery.Cells(ij.ColExpiry.Index).Value), "dd MMM yyyy")
                                    End If
                                    .Cells(Me.ColExpiryDate.Index).Value = ExpiredDateStr
                                    .Cells(Me.ColQtyMaster.Index).Value = 0
                                    .Cells(Me.ColQtyPerMaster.Index).Value = 0
                                    .Cells(Me.ColMasterUnits.Index).Value = ""
                                    .Cells(Me.ColQty.Index).Value = 0
                                    .Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(RowDgvDelivery.Cells(ij.ColQtyDelivery.Index).Value, ij.DefPrecDefault)
                                    .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                                    .Cells(Me.ColSecQty.Index).Value = -1
                                    .Cells(Me.ColSecUnit.Index).Value = Me.lblSecondaryUnits.Text
                                    .Cells(Me.ColRemarks.Index).Value = ""
                                    .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                                    .Cells(Me.ColPrimarySO.Index).Value = Me.txtPrimarySO.Text
                                    .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                                    .Cells(Me.ColPrimaryUnit.Index).Value = Me.txtPrimaryUnitsDefault.Text
                                    .Cells(Me.ColPrimarySecUnit.Index).Value = Me.txtPrimarySecondaryUnits.Text
                                    If IdxInput = IdxTotal Then
                                        .Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(JmlTotalInputMulti - JmlInputMulti, ij.DefPrecMulti)
                                    Else
                                        .Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(Math.Round((CDec(RowDgvDelivery.Cells(ij.ColQtyDelivery.Index).Value) * ij.RatioMulti) / _
                                        ij.RatioDefault, ij.DefPrecMulti, MidpointRounding.AwayFromZero), ij.DefPrecMulti)
                                    End If
                                    JmlInputMulti += CDec(.Cells(Me.ColTotalQtyMulti.Index).Value)
                                    .Cells(Me.ColUnitsCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                                    .Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, ij.DefPrecDefault)
                                    .Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(Me.TxtRatioReference.Text, ij.DefPrecMulti)
                                    .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                                    If Me.txtPrimaryUnitsDefault.Text <> Me.txtPrimaryUnitsMulti.Text Then
                                        .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, ij.DefPrecMulti) _
                                        & " " & Me.lblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, ij.DefPrecDefault) & " " & Me.lblUnitsDefault.Text
                                    Else
                                        .Cells(Me.ColRatioDesc.Index).Value = "-"
                                    End If
                                    'TR/Tambahan isi kolom delivery date
                                    If Me.ShowDeliveryDate = True Then
                                        If DeliveryDate = #1/1/1950# Then
                                            .Cells(Me.ColDeliveryDate.Index).Value = "-"
                                        Else
                                            .Cells(Me.ColDeliveryDate.Index).Value = Format(DeliveryDate, "dd MMM yyyy")
                                        End If
                                    Else
                                        .Cells(Me.ColDeliveryDate.Index).Value = "-"
                                    End If

                                    Dim DtaSOAddress As New DseTransaksiSOTableAdapters.tbltrans_so3TableAdapter
                                    Dim TblSOAddress As New DseTransaksiSO.tbltrans_so3DataTable
                                    Dim RowSOAddress As DseTransaksiSO.tbltrans_so3Row
                                    ''01 apakah default atau other customer
                                    DtaSOAddress.Connection.ConnectionString = Core_Module.GetConstring
                                    If DtaSOAddress.Fill(TblSOAddress, Me.txtPrimarySO.Text) > 0 Then
                                        RowSOAddress = TblSOAddress.Rows(0)
                                        ''01b update alamat bill to
                                        Me.txtCustomerAddIdentifier.Text = RowSOAddress.Indentifier_Alamat_Customer
                                        If Me.cboCustomerAddress.Items.Count > 0 Then
                                            Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                                        End If
                                        If RowSOAddress.primary_customer_ship_to = -1 Then
                                            ''02 jika default, isi ship to address
                                            Me.Panel1.Visible = True
                                            Me.TxtIdentifierAddressShipTo.Text = RowSOAddress.identifier_address_ship_to
                                            Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                                            If CboAddressShipTo.Items.Count > 0 Then
                                                Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressShipTo.Text)
                                            End If
                                            Me.Panel2.Visible = False
                                            Me.TxtPrimaryCustomerShipTo.Text = "-1"
                                            Me.TxtCustomerCodeShipTo.Text = ""
                                            Me.TxtCustomerDescShipTo.Text = ""
                                            TxtIdentifierAddressCustomerShipTo.Text = ""
                                            CboAddressCustomerShipTo.Items.Clear()
                                            Me.Label6.Location = New Point(Me.Label6.Location.X, Me.Label1.Location.Y)
                                            Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.TxtCustomerName.Location.Y + 5)
                                            Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, (Me.TxtCustomerName.Size.Height * 2) + (Me.cboCustomerAddress.Top - Me.TxtCustomerName.Bottom))
                                        Else
                                            ''03 jika other customer, isi customer ship to dan customer ship to address
                                            Me.Panel2.Visible = True
                                            Me.TxtPrimaryCustomerShipTo.Text = RowSOAddress.primary_customer_ship_to
                                            Me.TxtCustomerCodeShipTo.Text = RowSOAddress.Customer_Code
                                            Me.TxtCustomerDescShipTo.Text = RowSOAddress.Customer_Name
                                            Me.TxtIdentifierAddressCustomerShipTo.Text = RowSOAddress.identifier_address_customer_ship_to
                                            Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, RowSOAddress.primary_customer_ship_to)
                                            If Me.CboAddressCustomerShipTo.Items.Count > 0 Then
                                                Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressCustomerShipTo.Text)
                                            End If
                                            Me.Panel1.Visible = False
                                            TxtIdentifierAddressShipTo.Text = ""
                                            CboAddressShipTo.Items.Clear()
                                            Me.Label6.Location = New Point(Me.Label6.Location.X, Me.cboCustomerAddress.Location.Y)
                                            Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, Me.TxtCustomerName.Size.Height)
                                            Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.cboCustomerAddress.Location.Y)
                                        End If
                                    End If

                                End With

                                If Me.dgvItem.RowCount = 1 Then
                                    'set nilai acuan untuk pengecekan SO
                                    Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.txtPrimarySO.Text, Me.PrimaryCurrencySO, _
                                    Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
                                End If
                            Else
                                'tambahkan qty default + hitung ulang qty multi
                                RowDgv = Me.dgvItem.Rows(IdxAcuan)
                                RowDgv.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(RowDgvDelivery.Cells(ij.ColQtyDelivery.Index).Value) + _
                                CDec(RowDgv.Cells(Me.ColTotalQty.Index).Value), ij.DefPrecDefault)
                                RowDgv.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping((CDec(RowDgv.Cells(Me.ColTotalQty.Index).Value) * ij.RatioMulti) / _
                                ij.RatioDefault, ij.DefPrecMulti)
                            End If
                        End If
                    Next

                    'kendalikan focus item selanjutnya
                    Dim JmlObjLama, JmlSOBBaru, JmlDGV, JmlSOB As Decimal
                    JmlDGV = Me.GetJmlDGV_Multi(Me.txtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
                    JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)
                    JmlObjLama = Me.ObjDOItem.GetJml_Multi(Me.TxtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.txtPrimaryCodeNo.Text)

                    'perhitungan menggunakan nilai jumlah dgv yang baru
                    JmlSOBBaru = JmlSOB + JmlObjLama - JmlDGV
                    Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(0, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    Me.TxtQtyperBatchDefault.Text = ObjCD.DigitGrouping(0, Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
                    Me.TxtRemarkItem.Text = ""

                    TxtSOLama = Me.cboSalesOrderNo.Text
                    Me.SetGboItem(Me.cboSalesOrderNo.SelectedIndex, Me.cboCodeNo.SelectedIndex)

                    Me.cboCodeNo.Focus()

                    'format dgv
                    Me.FormatDGV()

                    'fokuskan ke baris terakhir
                    While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                    Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboAddressShipTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboAddressShipTo.SelectedIndexChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.CboAddressShipTo.SelectedIndex <> -1 Then
                    Me.TxtIdentifierAddressShipTo.Text = Me.ObjCustomer.GetIdentifierAdd(Me.txtPrimaryCustomer.Text, Me.CboAddressShipTo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbViewAddressShipTo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressShipTo.LinkClicked
        Try
            If Me.TxtRemark.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryDO.Text) AndAlso CDec(Me.TxtPrimaryDO.Text) > 0) Then
                If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                    MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.TxtIdentifierAddressShipTo.Text
                ij.PrimaryCust = Me.txtPrimaryCustomer.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                ij.cboCustAddCode.Enabled = False
                ij.ShowDialog()
                'If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                '    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                '    Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                'End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbViewAddressCustomerShipTo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressCustomerShipTo.LinkClicked
        Try
            If Me.TxtRemark.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryDO.Text) AndAlso CDec(Me.TxtPrimaryDO.Text) > 0) Then
                If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                    MsgBox("Please specify ship to customer first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.TxtIdentifierAddressCustomerShipTo.Text
                ij.PrimaryCust = Me.TxtPrimaryCustomerShipTo.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                ij.cboCustAddCode.Enabled = False
                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    'Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    'Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLCustom()
        Dim rowCustom As CustomSMESS.DseTransaksiDO_AddByCustomProduct.tbltrans_do_itemRow
        Dim rowDGV As DataGridViewRow

        Try
            Me.TBLCustomItem.Clear()

            For Each rowDGV In Me.dgvItem.Rows
                rowCustom = Me.TBLCustomItem.Newtbltrans_do_itemRow

                With rowDGV
                    rowCustom.primary_location = -1
                    rowCustom.Product_Code = .Cells(Me.ColCodeNo.Index).Value
                    rowCustom.Product_Desc = .Cells(Me.ColDescription.Index).Value
                    rowCustom.SO_Number = .Cells(Me.ColSONO.Index).Value
                    rowCustom.Batch_No = .Cells(Me.ColBatchNo.Index).Value

                    If .Cells(Me.ColExpiryDate.Index).Value = "-" Then
                        rowCustom.Expiry_Date = #1/1/1950#
                    Else
                        rowCustom.Expiry_Date = .Cells(Me.ColExpiryDate.Index).Value
                    End If
                    rowCustom.Qty_Master = 0
                    rowCustom.Qty_Per_Master = 0
                    rowCustom.Primary_Master_Units = -1
                    rowCustom.Qty = 0
                    rowCustom.Total_Qty = .Cells(Me.ColTotalQty.Index).Value
                    rowCustom.Unit_Code = .Cells(Me.ColUnit.Index).Value
                    rowCustom.Secondary_Qty = -1
                    rowCustom.Primary_Secondary_Units = -1
                    rowCustom.Remarks = .Cells(Me.ColRemarks.Index).Value
                    rowCustom.Primary_Product = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    rowCustom.Primary_SO = .Cells(Me.ColPrimarySO.Index).Value
                    rowCustom.Primary_Units = .Cells(Me.ColPrimaryUnit.Index).Value
                    rowCustom.Total_Qty_Multi = .Cells(Me.ColTotalQtyMulti.Index).Value
                    rowCustom.Unit_Code_Multi = .Cells(Me.ColUnitsCodeMulti.Index).Value
                    rowCustom.Ratio_Default = .Cells(Me.ColRatioDefault.Index).Value
                    rowCustom.Ratio_Reference = .Cells(Me.ColRatioReference.Index).Value
                    rowCustom.Primary_Units_Multi = .Cells(Me.ColPrimaryUnitsMulti.Index).Value
                    rowCustom.Primary_DO = Me.TxtPrimaryDO.Text
                    rowCustom.Primary = -1
                    rowCustom.Def_Prec_Mullti = Me.ObjCD.GetJmlPrecision1(.Cells(Me.ColTotalQtyMulti.Index).Value)
                    rowCustom.Def_Prec_Default = Me.ObjCD.GetJmlPrecision1(.Cells(Me.ColTotalQty.Index).Value)
                    rowCustom.Tolerance_Down = 1
                    rowCustom.Tolerance_Up = 1
                    rowCustom.Remarks = .Cells(Me.ColRemarks.Index).Value
                    'bagian kolom delivery date dianggap kosong dulu. Nanti jika ada perubahan, maka perubahan ini harus
                    'diupdate di software ABT 
                    rowCustom.Delivery_Date = #1/1/1950#

                    If Me.TxtPrimaryCustomerShipTo.Text = "" Then
                        Me.TxtPrimaryCustomerShipTo.Text = -1
                    End If

                    rowCustom.primary_customer_ship_to = Me.TxtPrimaryCustomerShipTo.Text
                    rowCustom.cust_ship_to_code = Me.TxtCustomerCodeShipTo.Text
                    rowCustom.cust_ship_to_name = Me.TxtCustomerDescShipTo.Text
                End With

                Me.TBLCustomItem.Rows.Add(rowCustom)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAdvance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAdvance.LinkClicked
        Dim rowCustomItem_Loc As CustomSMESS.DseTransaksiDO_Location.tbltrans_custom_doRow
        Dim vFrm As New CustomSMESS.FrmTransaksiDO_AddItemCustom
        Dim rowCustom As CustomSMESS.DseTransaksiDO_AddByCustomProduct.tbltrans_do_itemRow
        Dim rowDGV As DataGridViewRow
        Dim ExpiredDateStr As String

        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                    MsgBox("Please specify customer.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
                'modif 09/07/2020
                If Me.baru = True Then
                    vFrm.PrimaryCustLama = Me.txtPrimaryCustomer.Text
                Else
                    vFrm.PrimaryCustLama = Me.objDO.PrimaryCustomer
                End If
                vFrm.UserCode = Me.UserCode
                vFrm.PrimaryCust = Me.txtPrimaryCustomer.Text
                vFrm.vPrimaryDO = Me.TxtPrimaryDO.Text
                vFrm.BackColorAll = Me.BackColorAll
                vFrm.TBLDORef = Me.TBLCustomItem
                vFrm.TBLCustomItem_Loc = Me.TBLCustomItem_Loc

                'isikan table custom
                Me.IsikanTBLCustom()

                If vFrm.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.TBLCustomItem_Loc.Rows.Clear()
                    rowCustomItem_Loc = Me.TBLCustomItem_Loc.Newtbltrans_custom_doRow
                    rowCustomItem_Loc.primary = -1
                    rowCustomItem_Loc.primary_do = -1
                    rowCustomItem_Loc.primary_customer = -1
                    rowCustomItem_Loc.primary_location = vFrm.PrimaryLoc
                    Me.TBLCustomItem_Loc.Rows.Add(rowCustomItem_Loc)

                    'hapuskan dgv
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.dgvItem.Rows.Clear()

                    'isikan ke dalam dgv
                    For Each rowCustom In Me.TBLCustomItem.Rows
                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)

                        With rowDGV
                            .Cells(Me.ColPrimaryLocation.Index).Value = -1
                            .Cells(Me.ColLocation.Index).Value = ""
                            .Cells(Me.ColWarehouse.Index).Value = ""

                            .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                            .Cells(Me.ColCodeNo.Index).Value = rowCustom.Product_Code
                            .Cells(Me.ColDescription.Index).Value = rowCustom.Product_Desc
                            .Cells(Me.ColCustCodeNo.Index).Value = ""
                            .Cells(Me.ColCustDesc.Index).Value = ""
                            .Cells(Me.ColSONO.Index).Value = rowCustom.SO_Number
                            .Cells(Me.ColBatchNo.Index).Value = rowCustom.Batch_No
                            If rowCustom.Expiry_Date.Date = #1/1/1950# Then
                                ExpiredDateStr = "-"
                            Else
                                ExpiredDateStr = Format(rowCustom.Expiry_Date.Date, "dd MMM yyyy")
                            End If
                            .Cells(Me.ColExpiryDate.Index).Value = ExpiredDateStr
                            .Cells(Me.ColQtyMaster.Index).Value = 0
                            .Cells(Me.ColQtyPerMaster.Index).Value = 0
                            .Cells(Me.ColMasterUnits.Index).Value = ""
                            .Cells(Me.ColQty.Index).Value = 0
                            .Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(rowCustom.Total_Qty, rowCustom.Def_Prec_Default)
                            .Cells(Me.ColUnit.Index).Value = rowCustom.Unit_Code
                            .Cells(Me.ColSecQty.Index).Value = -1
                            .Cells(Me.ColSecUnit.Index).Value = ""
                            .Cells(Me.ColRemarks.Index).Value = rowCustom.Remarks
                            .Cells(Me.ColPrimaryCodeNo.Index).Value = rowCustom.Primary_Product
                            .Cells(Me.ColPrimarySO.Index).Value = rowCustom.Primary_SO
                            .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                            .Cells(Me.ColPrimaryUnit.Index).Value = rowCustom.Primary_Units
                            .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                            .Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(rowCustom.Total_Qty_Multi, rowCustom.Def_Prec_Mullti)
                            .Cells(Me.ColUnitsCodeMulti.Index).Value = rowCustom.Unit_Code_Multi
                            .Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(rowCustom.Ratio_Default, rowCustom.Def_Prec_Default)
                            .Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(rowCustom.Ratio_Reference, rowCustom.Def_Prec_Mullti)
                            .Cells(Me.ColPrimaryUnitsMulti.Index).Value = rowCustom.Primary_Units_Multi
                            If rowCustom.Primary_Units <> rowCustom.Primary_Units_Multi Then
                                .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(rowCustom.Ratio_Reference, rowCustom.Def_Prec_Mullti) _
                                & " " & rowCustom.Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(rowCustom.Ratio_Default, rowCustom.Def_Prec_Default) & " " & rowCustom.Unit_Code
                            Else
                                .Cells(Me.ColRatioDesc.Index).Value = "-"
                            End If

                            'bagian kolom delivery date dianggap kosong dulu. Nanti jika ada perubahan, maka perubahan ini harus
                            'diupdate di software ABT 
                            .Cells(Me.ColDeliveryDate.Index).Value = "-"
                            'If Me.ShowDeliveryDate = True Then
                            '  If DeliveryDate = #1/1/1950# Then
                            '    .Cells(Me.ColDeliveryDate.Index).Value = "-"
                            '  Else
                            '    .Cells(Me.ColDeliveryDate.Index).Value = Format(DeliveryDate, "dd MMM yyyy")
                            '  End If
                            'Else
                            '  .Cells(Me.ColDeliveryDate.Index).Value = "-"
                            'End If 
                        End With
                    Next

                    If Me.TBLCustomItem.Rows.Count > 0 Then
                        rowCustom = Me.TBLCustomItem.Rows(0)

                        Dim DtaSOAddress As New DseTransaksiSOTableAdapters.tbltrans_so3TableAdapter
                        Dim TblSOAddress As New DseTransaksiSO.tbltrans_so3DataTable
                        Dim RowSOAddress As DseTransaksiSO.tbltrans_so3Row
                        ''01 apakah default atau other customer
                        DtaSOAddress.Connection.ConnectionString = Core_Module.GetConstring
                        If DtaSOAddress.Fill(TblSOAddress, rowCustom.Primary_SO) > 0 Then
                            RowSOAddress = TblSOAddress.Rows(0)
                            ''01b update alamat bill to
                            Me.txtCustomerAddIdentifier.Text = RowSOAddress.Indentifier_Alamat_Customer
                            If Me.cboCustomerAddress.Items.Count > 0 Then
                                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                            End If
                            If RowSOAddress.primary_customer_ship_to = -1 Then
                                ''02 jika default, isi ship to address
                                Me.Panel1.Visible = True
                                Me.TxtIdentifierAddressShipTo.Text = RowSOAddress.identifier_address_ship_to
                                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                                If CboAddressShipTo.Items.Count > 0 Then
                                    Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressShipTo.Text)
                                End If
                                Me.Panel2.Visible = False
                                Me.TxtPrimaryCustomerShipTo.Text = "-1"
                                Me.TxtCustomerCodeShipTo.Text = ""
                                Me.TxtCustomerDescShipTo.Text = ""
                                TxtIdentifierAddressCustomerShipTo.Text = ""
                                CboAddressCustomerShipTo.Items.Clear()
                                Me.Label6.Location = New Point(Me.Label6.Location.X, Me.Label1.Location.Y)
                                Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.TxtCustomerName.Location.Y + 5)
                                Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, (Me.TxtCustomerName.Size.Height * 2) + (Me.cboCustomerAddress.Top - Me.TxtCustomerName.Bottom))
                            Else
                                ''03 jika other customer, isi customer ship to dan customer ship to address
                                Me.Panel2.Visible = True
                                Me.TxtPrimaryCustomerShipTo.Text = RowSOAddress.primary_customer_ship_to
                                Me.TxtCustomerCodeShipTo.Text = RowSOAddress.Customer_Code
                                Me.TxtCustomerDescShipTo.Text = RowSOAddress.Customer_Name
                                Me.TxtIdentifierAddressCustomerShipTo.Text = RowSOAddress.identifier_address_customer_ship_to
                                Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, RowSOAddress.primary_customer_ship_to)
                                If Me.CboAddressCustomerShipTo.Items.Count > 0 Then
                                    Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressCustomerShipTo.Text)
                                End If
                                Me.Panel1.Visible = False
                                TxtIdentifierAddressShipTo.Text = ""
                                CboAddressShipTo.Items.Clear()
                                Me.Label6.Location = New Point(Me.Label6.Location.X, Me.cboCustomerAddress.Location.Y)
                                Me.TxtRemark.Size = New Size(Me.TxtRemark.Size.Width, Me.TxtCustomerName.Size.Height)
                                Me.TxtRemark.Location = New Point(Me.TxtRemark.Location.X, Me.cboCustomerAddress.Location.Y)
                            End If
                        End If

                        DtaSOAddress.Dispose()
                        TblSOAddress.Dispose()

                        'set nilai acuan untuk pengecekan SO
                        Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(rowCustom.Primary_SO, Me.PrimaryCurrencySO, _
                        Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
                    End If

                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

                    'hitung ulang sob
                    Me.ChkIgnoreTolerance.Checked = True
                    Me.SetGboItem(Me.cboSalesOrderNo.SelectedIndex, Me.cboCodeNo.SelectedIndex)
                    Me.cboCodeNo.Focus()

                    'format dgv
                    Me.FormatDGV()

                    If Me.TxtRemark.Enabled = True Then
                        If Me.dgvItem.Rows.Count > 0 Then
                            Me.txtCustomerCode.Enabled = False
                            Me.btnFindCustomer.Enabled = False
                        Else
                            Me.txtCustomerCode.Enabled = True
                            Me.btnFindCustomer.Enabled = True
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vFrm = Nothing
        End Try
    End Sub


#Region "Fungsi untuk quick DO"
    Public Sub SetBaru()
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.objDO.ResetObject(Me.dgvItem)
            'Me.objDO.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            Me.TxtDONo.Text = Me.objDO.GetNoBaru1(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanLangsung()
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.objDO.SimpanBaru1(Me.dgvItem, Me.UserCode, Me) = True Then
                        Me.vPrimaryDO = Me.objDO.Primary
                        Me.vDONo = Me.objDO.NoTransaksi
                        Me.KosongkanItem()
                        Me.SalinKeField()
                        Me.NonAktifkanDetail()
                        Me.btnFind.Focus()

                        If vFromQuick = True Then
                            Me.Close()
                        End If
                    End If
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Me.NonAktifkanDetail()
                    Me.vStopCancellation = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnStopProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStopProcess.Click
        Try
            If MsgBox("Stop cancellation process?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.vStopCancellation = True
                Me.NonAktifkanDetail()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    ''Update SMESS 2026
    'Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
    '    Try
    '        If vFromQuick Then
    '            If Me.dtpDate.Value < vMinDate Then
    '                Me.dtpDate.Value = vMinDate
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
End Class