Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockOutAdj
    Public UserCode As String = "USER"
    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vFromPickUpOrder As Boolean = False
    Public vPrimaryPickUpOrder As Double = -1
    Public vPrimaryLoadStockOut As Double = -1

    Dim baru As Boolean
    Dim objStockOut As New clsTransaksiStockInOut
    Dim ObjCD As New ClsCekData
    Dim ObjUser As New ClsUser
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim objStockOutOutItem As New clsTransaksiStockInOutItem

    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL3 As New dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item1DataTable
    Dim TBL5 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item2DataTable
    Dim TBL6 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item3DataTable
    Dim tbl_Dummy As New dseTransaksiStockInOut.tblmaster_productqtyDataTable
    Dim row_Dummy As dseTransaksiStockInOut.tblmaster_productqtyRow

    Dim ObjLicense As New ClsLicense
    Dim buffer As String = ""
    Dim objItem As New ClsTransaksiMaterialPlanningGroupItem
    Dim BufferCboBatch As Integer = -1
    Dim BufferLocation As Integer = -1
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjLoc As New Core_Module.clsMasterStorageArea
    Dim QuickAddByRefDoc As Boolean = False
    Dim QuickAddByBarcode As Boolean = False
    Dim QuickAddByProdLevel As Boolean = False
    Dim QuickAddByBatchNo As Boolean = False
    Public TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
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
            Me.llbQuickClear.Enabled = True
            Me.llbQuickInput.Enabled = True
            Me.lblQuickInput3.Enabled = True
            Me.llbAdvance.Enabled = True

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutAdj_QuickInputProductLevel") = True Then
                Me.QuickAddByProdLevel = True
            Else
                Me.QuickAddByProdLevel = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutAdj_QuickInputDocRef") = True Then
                Me.QuickAddByRefDoc = True
            Else
                Me.QuickAddByRefDoc = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutAdj_QuickInputBarcode") = True Then
                Me.QuickAddByBarcode = True
            Else
                Me.QuickAddByBarcode = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutAdj_QuickInputBatchNo") = True Then
                Me.QuickAddByBatchNo = True
            Else
                Me.QuickAddByBatchNo = False
            End If
            If Me.QuickAddByBarcode = True Or Me.QuickAddByProdLevel = True Or Me.QuickAddByRefDoc = True Or Me.QuickAddByBatchNo = True Then
                Me.llbAdvance.Enabled = True
            Else
                Me.llbAdvance.Enabled = False
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
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
            Me.llbQuickClear.Enabled = False
            Me.llbQuickInput.Enabled = False
            Me.lblQuickInput3.Enabled = False
            Me.llbAdvance.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.TxtPrimaryAdjType.Text = -1
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.CboAdjType.SelectedIndex = -1
            Me.TblItemCustom.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnits.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyperBatch.Text = "0"
            Me.txtSecondaryQty.Text = "-1"
            Me.txtQtyOnHand.Text = "0"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnits.Left = Me.lblSecondaryUnits.Left
            'Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
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
            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNo.Text = ""
                Me.TxtCodeNo.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatch.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatch.Focus()
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
                    If row.Cells(9).Value = Me.cboBatchNo.Text And Me.cboBatchNo.Text <> "" And row.Cells(15).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(10).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(10).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
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
                JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.txtPrimaryPDO.Text, _
                Me.DtpExpired.Value)
                JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
                Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, _
                Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatch.Text)
                If JmlSetelahTambah < 0 Then
                    If Me.chkUseSecondaryUnit.Checked = True Then
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + _
                        Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, _
                        "Information")
                    Else
                        MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + _
                        CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                        "Information")
                    End If
                    Me.TxtQtyperBatch.Focus()
                    Return False
                End If

                '9.b. Cek stock pada tanggal yang bersangkutan
                JmlGudang = ObjSC.GetStock(Me.TxtPrimaryCodeNo.Text, Me.dtpDate.Value)
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
                            Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.txtPrimaryPDO.Text, _
                            Me.DtpExpired.Value)
                JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
                Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, _
                Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatch.Text)
                If JmlSetelahTambah < 0 Then
                    If JmlGudang + JmlObjLama - JmlDGV > 0 Then
                        If Me.chkUseSecondaryUnit.Checked = True Then
                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + _
                            Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                            Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, _
                            "Information")
                        Else
                            MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping((JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) + " " + _
                            CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                            "Information")
                        End If
                    Else
                        MsgBox("Stock level on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + _
                        " is 0. Entering this data will cause negative value in your stock.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Me.TxtQtyperBatch.Focus()
                    Return False
                End If
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

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal BatchNoControlled As Boolean, ByVal LocationControlled As Boolean, ByVal UseSecondaryUnit As Boolean) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(15).Value = vPrimaryCodeNo Then
                    OK = True
                    'cek secondary qty
                    If UseSecondaryUnit = True And OK = True Then
                        If row.Cells(7).Value <> Secondary_Qty Then
                            OK = False
                        End If
                    End If
                    'cek location
                    If LocationControlled = True And OK = True Then
                        If vPrimaryLocation <> row.Cells(16).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If BatchNoControlled = True Then
                        If vBatchNo <> row.Cells(9).Value Then
                            OK = False
                        End If
                    End If
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    If OK = True Then
                        Jml = Jml + CDec(row.Cells(5).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahTransferOtomatis(ByVal vTanggal As Date) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            'Transfer otomatis jika:
            '1. Tanggal invoice >= convertion date
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA.ScalarQuery_GetConvertionDate.Value.Date <= vTanggal.Date Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#

        Try
            'If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
            '    MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If
            If Me.ApakahTransferOtomatis(Me.dtpDate.Value) = False Then
                MsgBox("Please specify date after convertion date", MsgBoxStyle.Information, "Information")
                Me.dtpDate.Focus()
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("stock out number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'berkaitan dengan tanggal closing max
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

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_OUT_ADJUST, Me.TBLDetailLocation) = False Then
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
                    Dim PrimaryPOSO As Double = -1
                    Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                    Dim Batch As String = .Cells(Me.ColBatch.Index).Value
                    Dim QtyDef As Decimal = .Cells(Me.ColQty.Index).Value
                    Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                    If ObjCust.CompareQtySMESS_Custom(PrimaryProduct, PrimaryPOSO, PrimaryLocation, Batch, QtyDef, 0, TblDummy) = False Then
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
                            .Primary_Warehouse = Me.ObjLoc.GetPrimaryWarehouse(.Primary)
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
                .PrimaryAdjType = Me.TxtPrimaryAdjType.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim AdjCode As String = ""

        Try
            With Me.objStockOut
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.TxtPrimaryAdjType.Text = .PrimaryAdjType
                AdjCode = Me.objStockOut.GetCodeSettingByPrimary(.PrimaryAdjType)
                If AdjCode <> "" Then
                    Me.CboAdjType.Text = AdjCode
                Else
                    Me.CboAdjType.SelectedIndex = 0
                End If
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
                If IsNumeric(row.Cells(7).Value) = True Then
                    row.Cells(7).Value = Me.ObjCD.DigitGrouping(row.Cells(7).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.AturHakAksesFormCustom()
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
            Me.KosongkanField()
            Me.objStockOut.IsiCboSetting(Me.CboAdjType)
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.UserCode = "USER" Then
                Me.llbQuickClear.Visible = True
                Me.lblQuickInput3.Visible = True
            Else
                Me.llbQuickClear.Visible = False
                Me.lblQuickInput3.Visible = False
            End If
            If Me.vFromPickUpOrder = True Then
                Me.LoadPickUpOrder()
            End If
            If vPrimaryLoadStockOut <> -1 Then
                Me.ShowData(Me.vPrimaryLoadStockOut)
            End If
            Me.llbQuickInput.Visible = False
            Me.ShowIcon = False
            Me.CboAdjType.Anchor = Me.TxtStockInNo.Anchor
            Me.TxtAdjustmentTypeDesc.Anchor = Me.TxtRemark.Anchor
            Me.TxtAdjustmentTypeDesc.Left += 50
            Me.TxtAdjustmentTypeDesc.Width -= 50
            Me.CboAdjType.Width += 50
            '
            Me.llbAdvance.Parent = Me.gboItem
            Me.llbAdvance.BringToFront()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLDummy(ByVal TBLItem As DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_order_item1DataTable, _
    ByVal ListPrimary As List(Of Double), ByVal TBLLocation As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable, _
    ByVal TBLDummy As DseTransaksiPickUpOrder.TblDummyDataTable)
        Dim DTAQty As New DseTransaksiPickUpOrderTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLQty As New DseTransaksiPickUpOrder.tblmaster_productqtyDataTable
        Dim rowQty As DseTransaksiPickUpOrder.tblmaster_productqtyRow
        Dim rowItem As DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_order_item1Row
        Dim rowLocation As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row
        Dim rowDummy As DseTransaksiPickUpOrder.TblDummyRow

        Try
            DTAQty.Connection.ConnectionString = Core_Module.GetConstring

            For Each PrimaryCodeNo As Double In ListPrimary
                TBLItem.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                TBLItem.DefaultView.Sort = "primary"
                DTAQty.Fill(TBLQty, PrimaryCodeNo)

                For Each DV As Data.DataRowView In TBLItem.DefaultView
                    rowItem = DV.Row
                    If rowItem.Location_Controlled = 1 Then
                        'ambil dari location priority
                        For Each rowLocation In TBLLocation.Rows
                            TBLQty.DefaultView.RowFilter = "[Primary Storage Area] = " + rowLocation.primary_location.ToString
                            TBLQty.DefaultView.Sort = "[Expiry Date]"

                            While TBLQty.DefaultView.Count > 0 And rowItem.qty > 0
                                rowQty = TBLQty.DefaultView(0).Row
                                If rowQty.Qty <= rowItem.qty Then
                                    'qty di gudang lebih kecil atau sama dengan qty pick up
                                    'pindahkan ke tabel dummy
                                    rowDummy = TBLDummy.NewTblDummyRow
                                    rowDummy.expiry = rowQty.Expiry_Date
                                    rowDummy.lot_no = rowQty.Lot_No
                                    rowDummy.Qty = rowQty.Qty
                                    rowDummy.primary = rowItem.primary
                                    rowDummy.primary_storage_area = rowLocation.primary_location
                                    rowDummy.storage_area_code = rowLocation.Storage_Area_Code
                                    rowDummy.warehouse_code = rowLocation.Warehouse_Code
                                    rowDummy.urutan = TBLDummy.Rows.Count + 1
                                    TBLDummy.Rows.Add(rowDummy)

                                    rowItem.qty -= rowQty.Qty
                                    TBLQty.Rows.Remove(rowQty)
                                Else
                                    'qty pick up lebih kecil dr qty d gudang
                                    rowDummy = TBLDummy.NewTblDummyRow
                                    rowDummy.expiry = rowQty.Expiry_Date
                                    rowDummy.lot_no = rowQty.Lot_No
                                    rowDummy.Qty = rowItem.qty
                                    rowDummy.primary = rowItem.primary
                                    rowDummy.primary_storage_area = rowLocation.primary_location
                                    rowDummy.storage_area_code = rowLocation.Storage_Area_Code
                                    rowDummy.warehouse_code = rowLocation.Warehouse_Code
                                    rowDummy.urutan = TBLDummy.Rows.Count + 1
                                    TBLDummy.Rows.Add(rowDummy)
                                    rowQty.Qty -= rowItem.qty
                                    rowItem.qty = 0

                                End If
                            End While

                            If rowItem.qty = 0 Then
                                Exit For
                            End If
                        Next

                        'sisanya dari urutan berdasarkan kode lokasi dan expiry saja
                        If rowItem.qty > 0 Then
                            TBLQty.DefaultView.RowFilter = ""
                            TBLQty.DefaultView.Sort = "[Storage Area Code], [Expiry Date]"

                            While TBLQty.DefaultView.Count > 0 And rowItem.qty > 0
                                rowQty = TBLQty.DefaultView(0).Row
                                If rowQty.Qty <= rowItem.qty Then
                                    'qty di gudang lebih kecil atau sama dengan qty pick up
                                    'pindahkan ke tabel dummy
                                    rowDummy = TBLDummy.NewTblDummyRow
                                    rowDummy.expiry = rowQty.Expiry_Date
                                    rowDummy.lot_no = rowQty.Lot_No
                                    rowDummy.Qty = rowQty.Qty
                                    rowDummy.primary = rowItem.primary
                                    rowDummy.primary_storage_area = rowQty.Primary_Storage_Area
                                    rowDummy.storage_area_code = rowQty.Storage_Area_Code
                                    rowDummy.warehouse_code = rowQty.Warehouse_Code
                                    rowDummy.urutan = TBLDummy.Rows.Count + 1
                                    TBLDummy.Rows.Add(rowDummy)

                                    rowItem.qty -= rowQty.Qty
                                    TBLQty.Rows.Remove(rowQty)
                                Else
                                    'qty pick up lebih kecil dr qty d gudang
                                    rowDummy = TBLDummy.NewTblDummyRow
                                    rowDummy.expiry = rowQty.Expiry_Date
                                    rowDummy.lot_no = rowQty.Lot_No
                                    rowDummy.Qty = rowItem.qty
                                    rowDummy.primary = rowItem.primary
                                    rowDummy.primary_storage_area = rowQty.Primary_Storage_Area
                                    rowDummy.storage_area_code = rowQty.Storage_Area_Code
                                    rowDummy.warehouse_code = rowQty.Warehouse_Code
                                    rowDummy.urutan = TBLDummy.Rows.Count + 1
                                    TBLDummy.Rows.Add(rowDummy)
                                    rowQty.Qty -= rowItem.qty
                                    rowItem.qty = 0
                                End If
                            End While
                        End If
                    Else
                        TBLQty.DefaultView.RowFilter = ""
                        TBLQty.DefaultView.Sort = "[Storage Area Code], [Expiry Date]"

                        While TBLQty.DefaultView.Count > 0 And rowItem.qty > 0
                            rowQty = TBLQty.DefaultView(0).Row
                            If rowQty.Qty <= rowItem.qty Then
                                'qty di gudang lebih kecil atau sama dengan qty pick up
                                'pindahkan ke tabel dummy
                                rowDummy = TBLDummy.NewTblDummyRow
                                rowDummy.expiry = rowQty.Expiry_Date
                                rowDummy.lot_no = rowQty.Lot_No
                                rowDummy.Qty = rowQty.Qty
                                rowDummy.primary = rowItem.primary
                                rowDummy.primary_storage_area = rowQty.Primary_Storage_Area
                                rowDummy.storage_area_code = rowQty.Storage_Area_Code
                                rowDummy.warehouse_code = rowQty.Warehouse_Code
                                rowDummy.urutan = TBLDummy.Rows.Count + 1
                                TBLDummy.Rows.Add(rowDummy)

                                rowItem.qty -= rowQty.Qty
                                TBLQty.Rows.Remove(rowQty)
                            Else
                                'qty pick up lebih kecil dr qty d gudang
                                rowDummy = TBLDummy.NewTblDummyRow
                                rowDummy.expiry = rowQty.Expiry_Date
                                rowDummy.lot_no = rowQty.Lot_No
                                rowDummy.Qty = rowItem.qty
                                rowDummy.primary = rowItem.primary
                                rowDummy.primary_storage_area = rowQty.Primary_Storage_Area
                                rowDummy.storage_area_code = rowQty.Storage_Area_Code
                                rowDummy.warehouse_code = rowQty.Warehouse_Code
                                rowDummy.urutan = TBLDummy.Rows.Count + 1
                                TBLDummy.Rows.Add(rowDummy)
                                rowQty.Qty -= rowItem.qty
                                rowItem.qty = 0
                            End If
                        End While
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAQty.Dispose()
            TBLQty.Dispose()
        End Try
    End Sub

    Private Sub LoadPickUpOrder()
        Dim DTAHeader As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_orderTableAdapter
        Dim TBLHeader As New DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_orderDataTable
        Dim rowHeader As DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_orderRow
        Dim DTAItem As New DseTransaksiPickUpOrderTableAdapters.tbltrans_purchase_req_pickup_order_item1TableAdapter
        Dim TBLItem As New DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_order_item1DataTable
        Dim rowItem As DseTransaksiPickUpOrder.tbltrans_purchase_req_pickup_order_item1Row
        Dim DTALocation As New Core_Module.DseMasterLocationPriorityTableAdapters.tblmaster_location_priority1_TableAdapter
        Dim TBLLocation As New Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable
        Dim TipeProductGoods As Core_Module.clsMasterProduct.EnumTipeProduct = clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeLocationPriority As Core_Module.ClsMasterLocationPriority.EnumTipeTransaksi = ClsMasterLocationPriority.EnumTipeTransaksi.StockOutAdjust
        Dim ListPrimary As New List(Of Double)
        Dim TBLDummy As New DseTransaksiPickUpOrder.TblDummyDataTable
        Dim AdalocationControl As Boolean
        Dim ijk As New Core_Module.FrmMasterLocationPriority
        Dim rowDummy As DseTransaksiPickUpOrder.TblDummyRow
        Dim DummyRemarks As String = ""

        Try
            'new
            Me.btnNew.PerformClick()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'set tanggal dan remarks
            DTAHeader.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Connection.ConnectionString = DTAHeader.Connection.ConnectionString
            Core_Module.SetConstring_DTALocationPriority(DTALocation)

            DTAHeader.Fill(TBLHeader, Me.vPrimaryPickUpOrder)
            DTAItem.Fill(TBLItem, Me.vPrimaryPickUpOrder)
            If TBLHeader.Rows.Count = 0 Or TBLItem.Rows.Count = 0 Then
                MsgBox("Pick up order not found, process cancelled.", MsgBoxStyle.Information, "Information")
                Me.KosongkanField()
                Me.NonAktifkanDetail()
                Me.Close()
            Else
                rowHeader = TBLHeader.Rows(0)
                Me.dtpDate.Value = Today
                If dtpDate.Value.Date < rowHeader.tanggal.Date Then
                    dtpDate.Value = rowHeader.tanggal.Date
                End If
                Me.TxtRemark.Text = "Pick up order: " + rowHeader.no_transaksi + " - " + rowHeader.ref_party_desc
                'cek location priority jika ada location control
                For Each rowItem In TBLItem.Rows
                    If rowItem.Location_Controlled = 1 And rowItem.Type = TipeProductGoods Then
                        AdalocationControl = True
                        Exit For
                    End If
                Next
                If AdalocationControl = True Then
                    ijk.BackColorAll = Me.BackColorAll
                    ijk.UserCode = Me.UserCode
                    ijk.vTipeTransaksi = ClsMasterLocationPriority.EnumTipeTransaksi.StockOutAdjust
                    ijk.ShowDialog()
                End If

                'tambahkan ke dalam dgv berdasarkan location priority, expiry date (jika ada)
                DTALocation.Fill(TBLLocation, Me.UserCode, TipeLocationPriority)
                'untuk setiap code no (hanya yg goods aja), tentukan qty yang di allocate ke dalam stock out adjust -> berdasarkan prioritas lokasi, batch no, dan expiry
                For Each rowItem In TBLItem.Rows
                    If rowItem.Type = TipeProductGoods Then
                        If ListPrimary.Contains(rowItem.primary_code_no) = False Then
                            ListPrimary.Add(rowItem.primary_code_no)
                        End If
                    End If
                Next

                'isikan tbl dummy, untuk kemudian dipindahkan ke dalam dgv
                If ListPrimary.Count > 0 Then
                    Me.IsikanTBLDummy(TBLItem, ListPrimary, TBLLocation, TBLDummy)
                End If

                'tambahkan setiap item pick order ke dalam dgv
                For Each rowItem In TBLItem.Rows
                    If rowItem.Type = TipeProductGoods Then
                        'kalo goods tambahkan berdasarkan alokasi stock
                        TBLDummy.DefaultView.RowFilter = "primary = " + rowItem.primary.ToString
                        TBLDummy.DefaultView.Sort = "urutan"
                        If TBLDummy.DefaultView.Count = 0 Then
                            'kemungkinan tidak ada stock sama sekali
                        Else
                            For Each DV As Data.DataRowView In TBLDummy.DefaultView
                                rowDummy = DV.Row
                                DummyRemarks = rowItem.purch__req__no
                                If rowItem.remarks.Trim <> "" Then
                                    DummyRemarks = DummyRemarks + "-" + rowItem.remarks
                                End If

                                DummyRemarks = DummyRemarks + "(" + Me.ObjCD.DigitGrouping(rowDummy.Qty, rowItem._Decimal) + " " + rowItem.Unit_Code + ")"
                                Me.TambahkanItemDGV(rowItem.primary_code_no, rowDummy.primary_storage_area, rowDummy.lot_no, rowDummy.expiry, rowDummy.Qty, _
                                rowItem.primary_unit, rowItem.Product_Code, rowItem.Description, rowDummy.storage_area_code, rowDummy.warehouse_code, _
                                DummyRemarks, rowItem.Unit_Code, rowItem._Decimal)
                            Next
                        End If
                    Else
                        'kalo jasa langsung tambahkan
                        DummyRemarks = rowItem.purch__req__no
                        If rowItem.remarks.Trim <> "" Then
                            DummyRemarks = DummyRemarks + "-" + rowItem.remarks
                        End If
                        DummyRemarks = DummyRemarks + "(" + Me.ObjCD.DigitGrouping(rowItem.qty, rowItem._Decimal) + " " + rowItem.Unit_Code + ")"
                        Me.TambahkanItemDGV(rowItem.primary_code_no, -1, "", #1/1/1950#, rowItem.qty, _
                        rowItem.primary_unit, rowItem.Product_Code, rowItem.Description, "", "", DummyRemarks, rowItem.Unit_Code, rowItem._Decimal)
                    End If
                Next

                'kosongkan gbo item
                Me.KosongkanItem()
                'format dgv
                Me.FormatDGV()
                'fokuskan ke baris terakhir
                If Me.dgvItem.RowCount > 0 Then
                    While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                                   Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            End If

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.CboAdjType.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAHeader.Dispose()
            TBLHeader.Dispose()
            ijk.Dispose()
            DTALocation.Dispose()
            TBLLocation.Dispose()
            ListPrimary = Nothing
            TBLDummy.Dispose()
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
                        If Me.objStockOut.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, Me.vPrimaryPickUpOrder) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            If Me.vFromPickUpOrder = True Then
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockOut.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom, Me.vPrimaryPickUpOrder) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()

                            If Me.vFromPickUpOrder = True Then
                                Me.Close()
                            End If
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


                If Me.vFromPickUpOrder = True Then
                    Me.Close()
                End If
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
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

                If Me.vFromPickUpOrder = True Then
                    Me.Close()
                End If
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
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
            Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value.Date)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiStockInOut
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
                If Me.objStockOut.ApakahBolehEdit(Me.UserCode) = True Then
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
                            If Me.objStockOut.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, Me.vPrimaryPickUpOrder) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()


                                If Me.vFromPickUpOrder = True Then
                                    Me.Close()
                                End If
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objStockOut.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom, Me.vPrimaryPickUpOrder) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()


                                If Me.vFromPickUpOrder = True Then
                                    Me.Close()
                                End If
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
                If row.Cells(15).Value = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Lot_No = CStr(row.Cells(9).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(5).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(10).Value) = True Then
                            ED = row.Cells(10).Value
                        End If
                        Me.TBL4.Rows.Add(CStr(row.Cells(9).Value), CDbl(row.Cells(5).Value), ED)
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
                If row.Cells(15).Value = vPrimaryProduct And CStr(row.Cells(9).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL5.Rows
                        If row1.Lot_No = CStr(row.Cells(9).Value) And row1.Location = row.Cells(11).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(5).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL5.Rows.Add(CStr(row.Cells(9).Value), row.Cells(11).Value, CDbl(row.Cells(5).Value))
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
                If row.Cells(15).Value = vPrimaryProduct And CStr(row.Cells(9).Value) = vBatchNo And _
                row.Cells(16).Value = vPrimaryLocation And IsNumeric(row.Cells(7).Value) = True Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL6.Rows
                        If row1.Secondary_Qty = CDbl(row.Cells(7).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(5).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(10).Value) = True Then
                            ED = row.Cells(10).Value
                        End If
                        Me.TBL6.Rows.Add(CDbl(row.Cells(7).Value), CStr(row.Cells(5).Value), _
                        CStr(row.Cells(9).Value), CStr(row.Cells(11).Value))
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
            Me.chkPdoControlled, Me.txtPrimaryUnits, Me.txtPrimarySecondaryUnits)

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
                'Me.lblUnits.Left = Me.lblSecondaryUnits.Left
                'Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
                Me.cboSecondaryUnits.Visible = False
                Me.cboSecondaryUnits.DataSource = Nothing
                Me.cboSecondaryUnits.Items.Clear()
            Else
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                Me.lblUnits.Left = Me.lblSecQty.Left - 37
                Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
                Me.cboSecondaryUnits.Visible = True
                If Me.chkLocationControlled.Checked = False Then
                    'isikan data secondary jika perlu
                    Me.IsiCboSecondary(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryStockIn.Text, _
                    Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                End If
            End If

            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.lblUnits.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            Else
                Me.lblUnits.Text = "Units"
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
                        Me.TxtQtyperBatch.Focus()
                    End If

                    GetQtyAvailable()
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
                            Me.TxtQtyperBatch.Focus()
                        End If

                        GetQtyAvailable()
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
                        Me.TxtQtyperBatch.Focus()
                    End If
                    GetQtyAvailable()
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
                            Me.TxtQtyperBatch.Focus()
                        End If

                        GetQtyAvailable()
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
                    Me.TxtQtyperBatch.Focus()
                End If

                GetQtyAvailable()
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
    ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(10).Value) = False Then
                    If IsNumeric(row.Cells(7).Value) = False Then
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And row.Cells(19).Value = PrimaryPdO And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And row.Cells(7).Value = SecondaryQty And row.Cells(19).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(7).Value) = False Then
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And SecondaryQty = -1 And row.Cells(19).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(10).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And row.Cells(7).Value = SecondaryQty And row.Cells(19).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(10).Value) Then
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
                Me.TambahkanItemDGV(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, _
                Me.DtpExpired.Value, Me.TxtQtyperBatch.Text, Me.txtPrimaryUnits.Text, Me.TxtCodeNo.Text, Me.txtDesc.Text, _
                Me.cboLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, Me.lblUnits.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))

                'kosongkan gbo item
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
                'format dgv
                Me.FormatDGV()
                'fokuskan ke baris terakhir
                GetQtyAvailable()

                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
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
        Dim SettingAutosize As Boolean = False

        Try
            If Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells Then
                SettingAutosize = True
            End If
            If SettingAutosize = True Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            End If
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(10).Value) = False Then
                    row.Cells(10).Value = "-"
                Else
                    If CDate(row.Cells(10).Value).Date = #1/1/1950# Then
                        row.Cells(10).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(10).Value = Format(CDate(row.Cells(10).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(7).Value) = True Then
                    If row.Cells(7).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(7).Value = "-1" Then
                    row.Cells(7).Value = "-"
                    row.Cells(8).Value = "-"
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(16).Value = "-1" Then
                    row.Cells(11).Value = "-"
                    row.Cells(12).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(9).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItem.Columns(7).Visible = False
                Me.dgvItem.Columns(8).Visible = False
            Else
                Me.dgvItem.Columns(7).Visible = True
                Me.dgvItem.Columns(8).Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(11).Visible = False
                Me.dgvItem.Columns(12).Visible = False
            Else
                Me.dgvItem.Columns(11).Visible = True
                Me.dgvItem.Columns(12).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(9).Visible = False
            Else
                Me.dgvItem.Columns(9).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(10).Visible = False
            Else
                Me.dgvItem.Columns(10).Visible = True
            End If
            If SettingAutosize = True Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatch, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatch, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatch.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatch, sender, e)
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
                If e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 Or e.ColumnIndex = 11 _
                Or e.ColumnIndex = 13 Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = 7 Then
                        'cek secondary qty
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 9 Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItem.Item(15, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 10 Then
                        'cek expiry date
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 11 Then
                        'cek location controlled
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItem.Item(15, e.RowIndex).Value) = False Then
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
                    If e.ColumnIndex <> 10 Then
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

                    If e.ColumnIndex <> 13 Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(10).Value) = True Then
                            ExpiryDate = CDate(.Cells(10).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(7).Value) = True Then
                            SecondaryQty = .Cells(7).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = 5 Then
                            Me.EndEditQty(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                            GetQtyAvailable()
                        ElseIf e.ColumnIndex = 7 Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 9 Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                            GetQtyAvailable()
                        ElseIf e.ColumnIndex = 11 Then
                            Me.EndEditLocation(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                            GetQtyAvailable()
                        End If
                    End If
                Else
                    If e.ColumnIndex = 13 Then
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
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.ObjProduct.ApakahGoods(.Cells(15).Value) = True Then
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

    Private Sub EndEditSecondaryUnit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem.Item(10, e.RowIndex).Value = BufferExpiryDate
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
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
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
                    If CStr(row.Cells(1).Value) = CStr(row1.Cells(1).Value) And CStr(row.Cells(2).Value) = CStr(row1.Cells(2).Value) And _
                    CStr(row.Cells(3).Value) = CStr(row1.Cells(3).Value) And CStr(row.Cells(4).Value) = CStr(row1.Cells(4).Value) And _
                    CStr(row.Cells(9).Value) = CStr(row1.Cells(9).Value) And _
                    CStr(row.Cells(11).Value) = CStr(row1.Cells(11).Value) Then
                        If IsNumeric(row.Cells(7).Value) = True And IsNumeric(row1.Cells(7).Value) = True Then
                            If CDbl(row.Cells(7).Value) = CDbl(row1.Cells(7).Value) Then
                                Me.GabungkanBaris(row, row1)
                                Me.ObjCD.SetNumber(Me.dgvItem)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(7).Value) = False And IsNumeric(row1.Cells(7).Value) = False Then
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

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
    ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal

            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGVLain), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
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
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + "  minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGVLain), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
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
            GetQtyAvailable()
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
            Me.TxtRemark.BackColor = Me.BackColorAll
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
            Me.TxtCodeNo.BackColor = Me.BackColorAll
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

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.GotFocus
        Try
            Me.TxtQtyperBatch.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.LostFocus
        Try
            Me.TxtQtyperBatch.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQty.GotFocus
        Try
            Me.txtSecondaryQty.BackColor = Me.BackColorAll
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
            Me.TxtRemarkItem.BackColor = Me.BackColorAll
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

            If Me.cboBatchNo.SelectedIndex <> -1 Then
                If BufferCboBatch <> Me.cboBatchNo.SelectedIndex Then
                    GetQtyAvailable()
                    BufferCboBatch = cboBatchNo.SelectedIndex
                End If
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

                If Me.cboLocation.SelectedIndex <> -1 Then
                    If Me.BufferLocation <> Me.cboLocation.SelectedIndex Then
                        GetQtyAvailable()
                        BufferLocation = cboLocation.SelectedIndex
                    End If
                End If

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

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                    Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboAdjType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboAdjType.SelectedValueChanged
        Try
            Me.TxtPrimaryAdjType.Text = Me.objStockOut.GetPrimarySettingByCode(Me.CboAdjType.Text)
            Me.TxtAdjustmentTypeDesc.Text = Me.objStockOut.GetDescriptionSettingAdjustment(Me.TxtPrimaryAdjType.Text)
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
                i.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutAdjust
                i.ShowDialog()

                i = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetItemStart(ByVal TBL As dseLaporanStockTaking.tblmaster_productqtyDataTable, _
    ByVal StartDate As Date, ByVal vTBLRef As dseTransaksiStockInOut.tblmaster_product1DataTable)
        Dim DTA As New dseLaporanStockTakingTableAdapters.tbltrans_stock_taking_itemstart1TableAdapter

        Dim DTAInOut As New dseTransaksiStockInOutTableAdapters.tbltrans_stockinout_item_view0TableAdapter
        Dim TBLInOut As New dseTransaksiStockInOut.tbltrans_stockinout_item_view0DataTable
        Dim RowInOut As dseTransaksiStockInOut.tbltrans_stockinout_item_view0Row
        Dim TipeInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust

        Dim DTABefore As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBLBefore As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeDataTable
        Dim rowBefore As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeRow
        Dim DTAAfter As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBLAfter As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowAfter As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterRow

        Dim DTATransfer As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Dim TBLTransfer As New dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemDataTable
        Dim rowTransfer As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemRow

        Dim TipeProductGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            'start date termasuk, jadi semua transaksi stock pada tanggal stock taking dianggap dilakukan
            'setelah proses stock taking
            'hanya memperhitungkan item2 goods
            'stock in out
            DTAInOut.FillBy_TglDariSampai(TBLInOut, #1/1/2100#, StartDate, CInt(TipeProductGoods))
            For Each RowInOut In TBLInOut.Rows
                vTBLRef.DefaultView.RowFilter = "Primary =" + RowInOut.Primary_Product.ToString
                If vTBLRef.DefaultView.Count > 0 Then
                    TipeInOut = RowInOut.Tipe_Transaksi
                    With RowInOut
                        If TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                        TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                            'kurangkan item
                            Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                            .Expiry_Date, .Primary_Storage_Area)
                        ElseIf TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Or _
                            TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process Or _
                            TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Or _
                            TipeInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                            'tambahkan item
                            Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, .Batch_No, _
                            .Expiry_Date, .Primary_Storage_Area)
                        End If
                    End With
                End If
            Next

            'stock bundling -> pasti goods
            DTABefore.FillBy_TanggalSampaiDari2(TBLBefore, #1/1/2100#, StartDate, CInt(TipeProductGoods))
            For Each rowBefore In TBLBefore.Rows
                vTBLRef.DefaultView.RowFilter = "Primary =" + rowBefore.Primary_Product.ToString
                If vTBLRef.DefaultView.Count > 0 Then
                    With rowBefore
                        Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                        .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                    End With
                End If
            Next
            DTAAfter.FillBy_TanggalSampaiDari2(TBLAfter, #1/1/2100#, StartDate, CInt(TipeProductGoods))
            For Each rowAfter In TBLAfter.Rows
                vTBLRef.DefaultView.RowFilter = "Primary =" + rowAfter.Primary_Product.ToString
                If vTBLRef.DefaultView.Count > 0 Then
                    With rowAfter
                        Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                        .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                    End With
                End If
            Next

            'stock transfer-pasti goods, jadi tidak dibatasi lagi.
            DTATransfer.FillBy_TanggalSampaiDari(TBLTransfer, #1/1/2100#, StartDate)
            For Each rowTransfer In TBLTransfer.Rows
                With rowTransfer
                    vTBLRef.DefaultView.RowFilter = "Primary =" + rowTransfer.Primary_Product.ToString
                    If vTBLRef.DefaultView.Count > 0 Then
                        'kurangkan jumlah dari lokasi tujuan
                        Me.KurangkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                        .Batch_No, .Expiry_Date, .primary_destination_area)
                        'tambahkan jumlah dari lokasi asal
                        Me.TambahkanItem(TBL, .Primary_Product, .Qty, .Secondary_Qty, _
                        .Batch_No, .Expiry_Date, .Primary_Storage_Area)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAInOut.Dispose()
            TBLInOut.Dispose()
            DTABefore.Dispose()
            TBLBefore.Dispose()
            DTAAfter.Dispose()
            TBLAfter.Dispose()
            DTATransfer.Dispose()
            TBLTransfer.Dispose()
        End Try
    End Sub

    Private Sub KurangkanItem(ByVal TBL As dseLaporanStockTaking.tblmaster_productqtyDataTable, _
   ByVal vPrimaryProduct As Double, ByVal vQty As Double, ByVal vQtySecondary As Double, _
   ByVal vLotNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryStorageArea As Double)
        Dim filter As String = ""
        Dim rowInOut As dseLaporanStockTaking.tblmaster_productqtyRow

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + _
            " AND [Qty Secondary] = " + vQtySecondary.ToString + " AND [Primary Storage Area] = " + _
            vPrimaryStorageArea.ToString + " AND [Expiry Date] = #" + vExpiryDate.ToShortDateString + _
            "# AND [Lot No] = '" + vLotNo + "'"

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count = 0 Then
                'tambahkan baris baru
                rowInOut = TBL.Rows(TBL.Rows.Count - 1)
                filter = rowInOut.Primary
                rowInOut = TBL.Newtblmaster_productqtyRow
                With rowInOut
                    .Expiry_Date = vExpiryDate
                    If vPrimaryStorageArea <> -1 Then
                        .Location_Controlled = 1
                    Else
                        .Location_Controlled = 0
                    End If
                    .Lot_No = vLotNo
                    .Primary = CDbl(filter) + 1
                    .Primary_Product = vPrimaryProduct
                    .Primary_Storage_Area = vPrimaryStorageArea
                    .Qty = vQty * (-1)
                    .Qty_Secondary = vQtySecondary
                End With
                TBL.Rows.Add(rowInOut)
            Else
                'kurangkan nilai qty
                rowInOut = TBL.DefaultView(0).Row
                rowInOut.Qty -= vQty
                If rowInOut.Qty = 0 Then
                    TBL.Rows.Remove(rowInOut)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanItem(ByVal TBL As dseLaporanStockTaking.tblmaster_productqtyDataTable, _
  ByVal vPrimaryProduct As Double, ByVal vQty As Double, ByVal vQtySecondary As Double, _
  ByVal vLotNo As String, ByVal vExpiryDate As Date, ByVal vPrimaryStorageArea As Double)
        Dim filter As String = ""
        Dim rowInOut As dseLaporanStockTaking.tblmaster_productqtyRow

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + _
            " AND [Qty Secondary] = " + vQtySecondary.ToString + " AND [Primary Storage Area] = " + _
            vPrimaryStorageArea.ToString + " AND [Expiry Date] = #" + vExpiryDate.ToShortDateString + _
            "# AND [Lot No] = '" + vLotNo + "'"

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count = 0 Then
                'tambahkan baris baru
                If TBL.Rows.Count > 0 Then
                    rowInOut = TBL.Rows(TBL.Rows.Count - 1)
                    filter = rowInOut.Primary
                Else
                    filter = "0"
                End If

                rowInOut = TBL.Newtblmaster_productqtyRow
                With rowInOut
                    .Expiry_Date = vExpiryDate
                    If vPrimaryStorageArea <> -1 Then
                        .Location_Controlled = 1
                    Else
                        .Location_Controlled = 0
                    End If
                    .Lot_No = vLotNo
                    .Primary = CDbl(filter) + 1
                    .Primary_Product = vPrimaryProduct
                    .Primary_Storage_Area = vPrimaryStorageArea
                    .Qty = vQty
                    .Qty_Secondary = vQtySecondary
                End With
                TBL.Rows.Add(rowInOut)
            Else
                'tambahkan nilai qty
                rowInOut = TBL.DefaultView(0).Row
                rowInOut.Qty += vQty
                If rowInOut.Qty = 0 Then
                    TBL.Rows.Remove(rowInOut)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbQuickClear_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbQuickClear.LinkClicked
        Me.QuickClear()
    End Sub

    Private Sub QuickClear()
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tblmaster_product1TableAdapter
        Dim TBL_A As New dseTransaksiStockInOut.tblmaster_product1DataTable
        Dim TBLDummy As New dseTransaksiStockInOut.tblmaster_product1DataTable
        Dim row, rowDummy As dseTransaksiStockInOut.tblmaster_product1Row
        Dim Tipe As clsMasterProduct.EnumTipeProduct = clsMasterProduct.EnumTipeProduct.Goods
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim ExpiryRef As Date = #1/1/2500#
        Dim ij As New FrmTransaksiStockOutAdj_QuickInput

        Dim BatchNumber As String = ""
        Dim PrimaryLocation As Double = -1
        Dim Expiry As Date = #1/1/1950#
        Dim WarehouseCode, LocationCode As String

        Dim obj As New clsMasterStorageArea
        Dim Level As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim DTAQty As New dseLaporanStockTakingTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLQty As New dseLaporanStockTaking.tblmaster_productqtyDataTable
        Dim TBLQtyDummy As New dseLaporanStockTaking.tblmaster_productqtyDataTable
        Dim rowQty, rowQtyDummy As dseLaporanStockTaking.tblmaster_productqtyRow
        Dim DTALocation As New dseTransaksiStockInOutTableAdapters.TBLMasterLocationTableAdapter
        Dim TBLLocation As New dseTransaksiStockInOut.TBLMasterLocationDataTable
        Dim rowLocatio As dseTransaksiStockInOut.TBLMasterLocationRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAQty.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTALocation.Connection.ConnectionString = DTA.Connection.ConnectionString

            TBL_A.Rows.Clear()
            TBLQtyDummy.Rows.Clear()
            TBLQty.Rows.Clear()

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If Me.dgvItem.Rows.Count > 0 Then
                MsgBox("Please make sure dgv is empty.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Start quick clear?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            Else
                If MsgBox("Please make sure the date is the correct that you want to clear the data. Proceed?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
            End If

            If ij.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            ij.dgvItemProductGroup.AllowUserToAddRows = False

            'load data
            DTALocation.Fill(TBLLocation)
            For Each rowDGV1 In ij.dgvItemProductGroup.Rows
                Level = Me.objItem.GetProductLevel(rowDGV1.Cells(ij.ColProductGroup_Level.Index).Value)
                If Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                    DTAQty.FillBy_ProductGroup(TBLQtyDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                    DTA.FillBy_PrimaryProductGroup(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                    DTAQty.FillBy_Brand(TBLQtyDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                    DTA.FillBy_PrimaryBrand(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                    DTAQty.FillBy_SubBrand(TBLQtyDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                    DTA.FillBy_PrimarySubBrand(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                    DTAQty.FillBy_Category(TBLQtyDummy, CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value), CInt(Tipe))
                    DTA.FillBy_PrimaryCategory(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                Else
                    DTAQty.FillBy_PrimaryProduct(TBLQtyDummy, CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value), CInt(Tipe))
                    DTA.FillBy_PrimaryProduct(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                End If

                'pindahkan ke table, dipastikan 1 baris tidak dipindahkan 2x
                For Each rowQtyDummy In TBLQtyDummy.Rows
                    TBLQty.DefaultView.RowFilter = "[Primary Product] = " + rowQtyDummy.Primary_Product.ToString + _
                    " AND [Lot No] = '" + rowQtyDummy.Lot_No + "' AND [Primary Storage Area] = " + rowQtyDummy.Primary_Storage_Area.ToString
                    If TBLQty.DefaultView.Count = 0 Then
                        rowQty = TBLQty.Newtblmaster_productqtyRow
                        For i As Integer = 0 To TBLQty.Columns.Count - 1
                            rowQty.Item(i) = rowQtyDummy.Item(i)
                        Next
                        TBLQty.Rows.Add(rowQty)
                    End If
                Next
                For Each rowDummy In TBLDummy.Rows
                    TBL_A.DefaultView.RowFilter = "Primary = " + rowDummy.Primary.ToString
                    If TBL_A.DefaultView.Count = 0 Then
                        row = TBL_A.Newtblmaster_product1Row
                        For i As Integer = 0 To TBL_A.Columns.Count - 1
                            row.Item(i) = rowDummy.Item(i)
                        Next
                        TBL_A.Rows.Add(row)
                    End If
                Next
                TBLQtyDummy.Rows.Clear()
                TBLDummy.Rows.Clear()
            Next

            'filter tabel dummy berdasarkan tanggal
            Me.GetItemStart(TBLQty, Me.dtpDate.Value + TimeSpan.FromDays(1), TBL_A)

            'masukkan ke dalam dgv
            For Each rowQty In TBLQty.Rows
                If rowQty.Lot_No = "" Then
                    BatchNumber = ""
                    Expiry = #1/1/1950#
                Else
                    BatchNumber = rowQty.Lot_No
                    Expiry = rowQty.Expiry_Date
                End If
                If rowQty.Primary_Storage_Area <> -1 Then
                    PrimaryLocation = rowQty.Primary_Storage_Area
                    TBLLocation.DefaultView.RowFilter = "[Primary Location] = " + PrimaryLocation.ToString
                    If TBLLocation.DefaultView.Count = 0 Then
                        MsgBox("Logic error, please contact your vendor.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        rowLocatio = TBLLocation.DefaultView(0).Row
                        LocationCode = rowLocatio.Storage_Area_Code
                        WarehouseCode = rowLocatio.Warehouse_Code
                    End If
                Else
                    PrimaryLocation = -1
                    LocationCode = ""
                    WarehouseCode = ""
                End If
                TBL_A.DefaultView.RowFilter = "Primary = " + rowQty.Primary_Product.ToString
                If TBL_A.DefaultView.Count = 0 Then
                    MsgBox("Logic error, please contact your vendor.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    row = TBL_A.DefaultView(0).Row
                End If

                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColBatch.Index).Value = BatchNumber
                    .Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                    .Cells(Me.ColDescription.Index).Value = row.Description
                    If Expiry = #1/1/1950# Then
                        .Cells(Me.ColExpiredDate.Index).Value = "-"
                    Else
                        .Cells(Me.ColExpiredDate.Index).Value = Format(Expiry, "dd MMM yyyy")
                    End If
                    .Cells(Me.ColLocation.Index).Value = LocationCode
                    .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColPONoDONo.Index).Value = ""
                    .Cells(Me.ColPrimaryCodeNo.Index).Value = rowQty.Primary_Product
                    .Cells(Me.ColPrimaryLocation.Index).Value = PrimaryLocation
                    .Cells(Me.ColPrimaryPdO.Index).Value = -1
                    .Cells(Me.ColPrimaryPOnoDoNo.Index).Value = -1
                    .Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                    .Cells(Me.ColPrimaryUnits.Index).Value = row.Primary_Unit
                    .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowQty.Qty, row._Decimal)
                    .Cells(Me.ColRemark.Index).Value = ""
                    .Cells(Me.ColSecondaryQty.Index).Value = -1
                    .Cells(Me.ColSecondaryUnit.Index).Value = ""
                    .Cells(Me.ColSuppDONo.Index).Value = ""
                    .Cells(Me.ColUnit.Index).Value = row.Unit_Code
                    .Cells(Me.ColWareHouse.Index).Value = WarehouseCode
                End With
            Next

            FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL_A.Dispose()
            obj = Nothing
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            ij = Nothing
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTAQty.Dispose()
            TBLQty.Dispose()
            DTALocation.Dispose()
            TBLLocation.Dispose()
        End Try
    End Sub

    Private Sub TambahkanItemDGV(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal vExpiryDate As Date, ByVal vQty As Decimal, ByVal vPrimaryUnit As Double, ByVal vCodeNo As String, ByVal vDesc As String, ByVal vLocationCode As String, _
    ByVal vWarehouseCode As String, ByVal vRemarksItem As String, ByVal vUnitCode As String, ByVal vDefPrecUnit As Integer)
        Dim Index As Integer = -1
        Dim DummyRemarks As String

        Try
            Index = Me.GetIndex(vPrimaryCodeNo, vPrimaryLocation, vBatchNo, _
            -1, -1, vExpiryDate)

            'cari data yang sama persis
            If Index <> -1 Then
                'tambahkan nilai qty
                Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                If row.Cells(Me.ColRemark.Index).Value.ToString.Contains(vRemarksItem) = False And vRemarksItem <> "" Then
                    DummyRemarks = row.Cells(Me.ColRemark.Index).Value + ";" + vRemarksItem
                    If DummyRemarks.Length > 150 Then
                        DummyRemarks = DummyRemarks.Substring(0, 150)
                    End If
                    row.Cells(Me.ColRemark.Index).Value = DummyRemarks
                End If
                row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + CDbl(vQty), vDefPrecUnit)
            Else
                'tambahkan baris baru
                Dim ExpiredDateStr As String
                If vExpiryDate.Date = #1/1/1950# Then
                    ExpiredDateStr = "-"
                Else
                    ExpiredDateStr = Format(vExpiryDate.Date, "dd MMM yyyy")
                End If

                Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", vCodeNo, vDesc, Me.ObjCD.DigitGrouping(vQty, vDefPrecUnit), _
                vUnitCode, -1, "", vBatchNo, _
                ExpiredDateStr, vLocationCode, vWarehouseCode, vRemarksItem, _
                -1, vPrimaryCodeNo, vPrimaryLocation, vPrimaryUnit, _
                -1, -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowData(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Not vPrimary = -1 Then
                    If Me.objStockOut.CariDetail(vPrimary, Me.dgvItem, Me.TblItemCustom) = True Then
                        Me.SalinKeField()
                        Me.btnEdit.Focus()
                    Else
                        Me.btnFind.Focus()
                    End If
                    Me.btnFind.Focus()
                End If

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
                    Me.vPrimaryPickUpOrder = Me.objStockOut.GetPrimaryPickUpOrder(vPrimary)
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
                If Me.objStockOut.ApakahBolehEdit(Me.UserCode) = True Then
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
                        If Me.objStockOut.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom, Me.vPrimaryPickUpOrder) Then
                            'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()

                            If Me.vFromPickUpOrder = True Then
                                Me.Close()
                            End If
                        End If
                        'End If
                    End If
                    'End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbQuickInput_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbQuickInput.LinkClicked

        Try
            IsiDGVQuickIput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDGVQuickIput()
        Dim DSEReset As New dseTransaksiStockInOutQuickInputTableAdapters.februariTableAdapter
        Dim tblReset As New dseTransaksiStockInOutQuickInput.februariDataTable
        Dim rowReset As dseTransaksiStockInOutQuickInput.februariRow

        Dim DSEStok As New dseTransaksiStockInOutQuickInputTableAdapters.tblStokTableAdapter
        Dim tblStok As New dseTransaksiStockInOutQuickInput.tblStokDataTable
        Dim rowStok As dseTransaksiStockInOutQuickInput.tblStokRow

        Dim DSESC As New dseTransaksiStockInOutQuickInputTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim tblSC As New dseTransaksiStockInOutQuickInput.tbltransaksi_stock_cardDataTable
        Dim rowSC As dseTransaksiStockInOutQuickInput.tbltransaksi_stock_cardRow

        Dim DSEP As New dseTransaksiStockInOutQuickInputTableAdapters.tblmaster_productTableAdapter
        Dim tblP As New dseTransaksiStockInOutQuickInput.tblmaster_productDataTable
        Dim rowP As dseTransaksiStockInOutQuickInput.tblmaster_productRow
        Dim dtDate As Date = CDate("2016/02/29")
        Dim tblDGV As New dseTransaksiStockInOutQuickInput.isiDGVDataTable
        Dim rowDGV As dseTransaksiStockInOutQuickInput.isiDGVRow

        Dim objunit As New Core_Module.ClsMasterUnit
        Dim objSC As New ClsTransaksiStockCard
        Dim ListProduct As New List(Of String)
        Dim vListPrimaryCategory As New List(Of Double)
        Dim TBLDummy_P As New dseTransaksiStockInOutQuickInput.tblmaster_productDataTable
        Dim decS As Decimal = 0

        Try
            DSEReset.Fill(tblReset)
            DSEStok.Fill(tblStok)
            DSEP.Fill(tblP)
            DSESC.Fill(tblSC, dtDate)

            'tambahkan tabel reset dengan semua data product yang ada dimaster 6 category
            vListPrimaryCategory.Clear()
            vListPrimaryCategory.Add(3) 'slat chemical
            vListPrimaryCategory.Add(4) 'slat packaging
            vListPrimaryCategory.Add(106) 'pencil packaging
            vListPrimaryCategory.Add(105) 'pencil lead
            vListPrimaryCategory.Add(107) 'pencil indirect
            vListPrimaryCategory.Add(55) 'sparepart

            For Each PrimaryCategory As Double In vListPrimaryCategory
                TBLDummy_P.Rows.Clear()
                DSEP.FillBy_primary_category(TBLDummy_P, PrimaryCategory)
                For Each rowP1 As dseTransaksiStockInOutQuickInput.tblmaster_productRow In TBLDummy_P.Rows
                    tblReset.DefaultView.RowFilter = "[code no smess] = '" + rowP1.Product_Code + "'"
                    If tblReset.DefaultView.Count = 0 Then
                        tblSC.DefaultView.RowFilter = Nothing
                        tblSC.DefaultView.RowFilter = "[Primary Product] = " & rowP1.Primary
                        If tblSC.DefaultView.Count > 0 Then
                            rowSC = tblSC.DefaultView(0).Row
                            decS = rowSC.saldo
                            If decS > 0 Then
                                With rowP1
                                    rowDGV = tblDGV.NewisiDGVRow
                                    rowDGV.ColPONoDONo = ""
                                    rowDGV.ColSuppDONo = ""
                                    rowDGV.ColCodeNo = .Product_Code
                                    rowDGV.ColDescription = .Description
                                    rowDGV.ColUnit = objunit.GetUnitCode(.Primary_Unit)
                                    rowDGV.ColSecondaryQty = 0
                                    rowDGV.ColSecondaryUnit = ""
                                    rowDGV.ColRemark = ""
                                    rowDGV.ColPrimaryPOnoDoNo = -1
                                    rowDGV.ColPrimaryCodeNo = .Primary
                                    rowDGV.ColPrimaryUnits = .Primary_Unit
                                    rowDGV.ColPrimarySecondaryUnits = -1
                                    rowDGV.ColPrimaryPdO = -1
                                    rowDGV.DefPrec = ._Decimal

                                    rowDGV.ColSelisih = decS
                                    rowDGV.ColQty = 0

                                    tblDGV.Rows.Add(rowDGV)
                                End With
                            End If

                        End If

                    End If
                Next
            Next

            If tblReset.Rows.Count > 0 Then
                For Each rowReset In tblReset.Rows
                    With rowReset
                        tblP.DefaultView.RowFilter = Nothing
                        tblP.DefaultView.RowFilter = "[Product Code] = '" & .code_no_smess & "'"

                        If tblP.DefaultView.Count > 0 Then
                            rowP = tblP.DefaultView(0).Row
                            tblSC.DefaultView.RowFilter = Nothing
                            tblSC.DefaultView.RowFilter = "[Primary Product] = " & rowP.Primary
                            If tblSC.DefaultView.Count > 0 Then
                                decS = 0
                                rowSC = tblSC.DefaultView(0).Row
                                decS = rowSC.saldo

                                If decS > .qty Then
                                    rowDGV = tblDGV.NewisiDGVRow
                                    rowDGV.ColPONoDONo = ""
                                    rowDGV.ColSuppDONo = ""
                                    rowDGV.ColCodeNo = .code_no_smess
                                    rowDGV.ColDescription = rowP.Description
                                    rowDGV.ColUnit = objunit.GetUnitCode(rowP.Primary_Unit)
                                    rowDGV.ColSecondaryQty = 0
                                    rowDGV.ColSecondaryUnit = ""
                                    rowDGV.ColRemark = ""
                                    rowDGV.ColPrimaryPOnoDoNo = -1
                                    rowDGV.ColPrimaryCodeNo = rowP.Primary
                                    rowDGV.ColPrimaryUnits = rowP.Primary_Unit
                                    rowDGV.ColPrimarySecondaryUnits = -1
                                    rowDGV.ColPrimaryPdO = -1
                                    rowDGV.DefPrec = rowP._Decimal
                                    rowDGV.ColSelisih = decS - CDec(.qty)
                                    rowDGV.ColQty = CDec(.qty)
                                    tblDGV.Rows.Add(rowDGV)
                                End If
                            End If

                            'decS = objSC.GetStock(rowP.Primary, dtDate)
                            'If decS > .qty Then
                            '    rowDGV = tblDGV.NewisiDGVRow
                            '    rowDGV.ColPONoDONo = ""
                            '    rowDGV.ColSuppDONo = ""
                            '    rowDGV.ColCodeNo = .code_no_smess
                            '    rowDGV.ColDescription = rowP.Description
                            '    rowDGV.ColUnit = objunit.GetUnitCode(rowP.Primary_Unit)
                            '    rowDGV.ColSecondaryQty = 0
                            '    rowDGV.ColSecondaryUnit = ""
                            '    rowDGV.ColRemark = ""
                            '    rowDGV.ColPrimaryPOnoDoNo = -1
                            '    rowDGV.ColPrimaryCodeNo = rowP.Primary
                            '    rowDGV.ColPrimaryUnits = rowP.Primary_Unit
                            '    rowDGV.ColPrimarySecondaryUnits = -1
                            '    rowDGV.ColPrimaryPdO = -1
                            '    rowDGV.ColSelisih = decS - CDec(.qty)
                            '    tblDGV.Rows.Add(rowDGV)
                            'End If
                        Else
                            ListProduct.Add(.code_no_smess)
                        End If
                    End With
                Next
            End If

            'tambahkan semua item pada category ybs yang belum ada di dalam selisih


            If ListProduct.Count > 0 Then
                Dim msg As String = ""
                For intA As Integer = 0 To ListProduct.Count - 1
                    If msg = "" Then
                        msg &= ListProduct.Item(intA)
                    Else
                        msg &= ", " & ListProduct.Item(intA)
                    End If
                Next
                If msg <> "" Then
                    MsgBox("There are list of code no smess were not found on tblmaster product : " & vbNewLine & msg, MsgBoxStyle.Information, "Information")
                End If
            End If


            Dim ListX As New List(Of String)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each rowDGV In tblDGV.Rows
                With rowDGV
                    tblStok.DefaultView.RowFilter = Nothing
                    tblStok.DefaultView.RowFilter = "[Primary Product] = " & .ColPrimaryCodeNo
                    tblStok.DefaultView.Sort = "[Qty] Asc "
                    Dim totalStok As Decimal = 0
                    Dim decSelisih As Decimal = 0
                    Dim decDetStok As Decimal = 0
                    decSelisih = .ColSelisih
                    totalStok = 0
                    For Each dv1 As Data.DataRowView In tblStok.DefaultView
                        rowStok = dv1.Row
                        totalStok += rowStok.Qty
                    Next
                    'totalStok = .ColQty
                    If decSelisih > totalStok Then
                        ListX.Add(.ColCodeNo)
                        ' decDetStok = decSelisih
                    Else
                        Dim intA As Integer = 1
                        If tblStok.DefaultView.Count > 0 Then
                            While decSelisih <> decDetStok
                                rowStok = tblStok.DefaultView(tblStok.DefaultView.Count - intA).Row
                                If (decSelisih - decDetStok) <= CDec(rowStok.Qty) Then
                                    Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                                    .ColDescription, Me.ObjCD.DigitGrouping((decSelisih - decDetStok), .DefPrec), _
                                     .ColUnit, -1, "", rowStok.Lot_No, _
                                     rowStok.Expiry_Date, rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                                     -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                                     -1, -1)
                                    decDetStok += (decSelisih - decDetStok)
                                ElseIf (decSelisih - decDetStok) > CDec(rowStok.Qty) Then
                                    Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                                    .ColDescription, Me.ObjCD.DigitGrouping(CDec(rowStok.Qty), .DefPrec), _
                                    .ColUnit, -1, "", rowStok.Lot_No, _
                                    rowStok.Expiry_Date, rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                                    -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                                    -1, -1)
                                    decDetStok += CDec(rowStok.Qty)
                                End If
                                If tblStok.DefaultView.Count - intA > 0 Then
                                    intA += 1
                                End If
                            End While
                        End If


                        'If tblStok.DefaultView.Count > 0 Then
                        '    While decSelisih > 0
                        '        rowStok = tblStok.DefaultView(tblStok.DefaultView.Count - 1).Row
                        '        If decSelisih <= CDec(rowStok.Qty) Then
                        '            Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                        '            .ColDescription, Me.ObjCD.DigitGrouping(decSelisih, .DefPrec), _
                        '             .ColUnit, -1, "", rowStok.Lot_No, _
                        '             rowStok.Expiry_Date, rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                        '             -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                        '             -1, -1)
                        '            decSelisih -= decSelisih

                        '            ' tblStok.Rows.Remove(rowStok)
                        '        Else
                        '            Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                        '            .ColDescription, Me.ObjCD.DigitGrouping(CDec(rowStok.Qty), .DefPrec), _
                        '            .ColUnit, -1, "", rowStok.Lot_No, _
                        '            rowStok.Expiry_Date, rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                        '            -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                        '            -1, -1)
                        '            decSelisih -= CDec(rowStok.Qty)
                        '            'tblStok.Rows.Remove(rowStok)
                        '        End If
                        '    End While
                        'End If
                    End If
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If ListX.Count > 0 Then
                Dim msg As String = ""
                For intA As Integer = 0 To ListX.Count - 1
                    If msg = "" Then
                        msg &= ListX.Item(intA)
                    Else
                        msg &= ", " & ListX.Item(intA)
                    End If
                Next
                If msg <> "" Then
                    MsgBox("There are list of code no smess with condition stok for now > stok 29 feb 2016 : " & vbNewLine & msg, MsgBoxStyle.Information, "Information")
                End If
            End If

            Me.FormatDGV()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '# Quick Input 3 - From Batch No
    Private Sub LlbQuikInput3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblQuickInput3.LinkClicked
        Me.AddByBatchNo()
    End Sub

    Private Sub AddByBatchNo()
        Dim hj As New Core_Module.FrmSetting_QuickInputAdj
        Try

            hj.BackColorAll = Me.BackColorAll
            hj.Text = "Quick Input"
            If hj.ShowDialog = Windows.Forms.DialogResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.objStockOut.Isitbl_ByBatch(Me.tbl_Dummy, hj.txtBatchNo.Text.Trim)
                Me.GetDatagrid_FromBatchNo(hj.txtBatchNo.Text.Trim, hj.txtLimit.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Public Function GetDatagrid_FromBatchNo(ByVal vBatch As String, ByVal vLimit As Decimal) As Boolean
        Dim count As Integer = 1
        Dim count2 As Integer = 5
        Dim rowDgv As DataGridViewRow
        Dim List As New List(Of String)
        Dim dummy As String = ""
        Try

            If tbl_Dummy.Rows.Count > 0 Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.Rows.Clear()

                For Each row_Dummy In tbl_Dummy.Rows
                    If count <= vLimit Then
                        If row_Dummy.Qty > 0 Then
                            Me.dgvItem.Rows.Add()
                            rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                            With rowDgv
                                .Cells(Me.ColPrimaryCodeNo.Index).Value = row_Dummy.Primary_Product
                                .Cells(Me.ColCodeNo.Index).Value = row_Dummy.Product_Code
                                .Cells(Me.ColDescription.Index).Value = row_Dummy.Product_Desc
                                .Cells(Me.ColUnit.Index).Value = row_Dummy.Unit_Code
                                .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row_Dummy.Qty, row_Dummy._Decimal)
                                .Cells(Me.ColPrimaryUnits.Index).Value = row_Dummy.Primary_Unit
                                .Cells(Me.ColSecondaryQty.Index).Value = row_Dummy.Qty_Sec

                                If row_Dummy.Primary_Secondary_Unit <> -1 Then
                                    .Cells(Me.ColPrimarySecondaryUnits.Index).Value = row_Dummy.Primary_Secondary_Unit
                                    .Cells(Me.ColSecondaryUnit.Index).Value = ""
                                Else
                                    .Cells(Me.ColPrimarySecondaryUnits.Index).Value = "-1"
                                    .Cells(Me.ColSecondaryUnit.Index).Value = ""
                                End If
                                .Cells(Me.ColBatch.Index).Value = row_Dummy.Lot_No
                                .Cells(Me.ColPrimaryLocation.Index).Value = row_Dummy.Primary_Storage_Area
                                If row_Dummy.Primary_Storage_Area = -1 Then
                                    .Cells(Me.ColLocation.Index).Value = ""
                                    .Cells(Me.ColWareHouse.Index).Value = ""
                                Else
                                    .Cells(Me.ColLocation.Index).Value = row_Dummy.Storage_Area_Code
                                    .Cells(Me.ColWareHouse.Index).Value = row_Dummy.Warehouse_Code
                                End If
                                If row_Dummy.Expiry_Date.Date = #1/1/1950# Then
                                    .Cells(Me.ColExpiredDate.Index).Value = "-"
                                Else
                                    .Cells(Me.ColExpiredDate.Index).Value = Format(row_Dummy.Expiry_Date, "dd MMM yyyy")
                                End If
                                .Cells(Me.ColPrimaryPOnoDoNo.Index).Value = "-1"
                                .Cells(Me.ColPONoDONo.Index).Value = ""
                                .Cells(Me.ColPrimaryPdO.Index).Value = "-1"
                                .Cells(Me.ColSuppDONo.Index).Value = ""
                                .Cells(Me.ColRemark.Index).Value = ""
                            End With
                            count += 1
                        Else
                            If List.Contains(row_Dummy.Product_Code) = False Then
                                List.Add(row_Dummy.Product_Code)
                                If count2 <= 5 Then
                                    If dummy = "" Then
                                        dummy = row_Dummy.Product_Code
                                    Else
                                        dummy &= ", " & row_Dummy.Product_Code
                                    End If
                                    count2 += 1
                                End If
                            End If
                        End If
                    End If
                Next

                If dummy <> "" Then
                    MsgBox("Product code :" & dummy & " quantity < 0", MsgBoxStyle.Information, "Information")
                End If
                ObjCD.SetNumber(Me.dgvItem)
                Me.FormatDGV()
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Else
                MsgBox("Product code is not found with batch no : " & vBatch.ToString.Trim & ", please check stock and try again.", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GetQtyAvailable()
        Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
        Try
            If Me.TxtPrimaryCodeNo.Text <> -1 Then
                Dim Decimal_ As Decimal = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlObjLama = Me.objStockOutOutItem.GetJml(Me.TxtPrimaryStockIn.Text, TxtPrimaryCodeNo.Text, _
                txtSecondaryQty.Text, txtPrimaryLocation.Text, cboBatchNo.Text, txtPrimaryPDO.Text, DtpExpired.Value.Date)

                JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, _
                            Me.chkBatchNoControlled.Checked, Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
                JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV
                Me.txtQtyOnHand.Text = Me.ObjCD.DigitGrouping(JmlSetelahTambah, Decimal_)
                'If JmlGudang >= JmlObjLama Then
                '    '- CDbl(Me.TxtQtyperBatch.Text) - JmlObjLama
                'Else
                '    JmlSetelahTambah = JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatch.Text)
                'End If

                'JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGV

                'If JmlGudang >= JmlObjLama Then
                '    Me.txtQtyOnHand.Text = Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Decimal_)
                'Else
                '    Me.txtQtyOnHand.Text = Me.ObjCD.DigitGrouping(CStr(JmlObjLama - JmlDGV), Decimal_)
                'End If

            Else
                Me.txtQtyOnHand.Text = "0"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_OutAdjust_Create(ByVal vPrimaryTransaksi As Double)
        Try
            Me.baru = False
            Me.objStockOut.ResetObject(Me.dgvItem)
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust

            Me.objStockOut.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.TblItemCustom)
            Me.UserCode = Me.objStockOut.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_OutAdjust_Save() As Double
        Try
            Me.baru = True
            Me.objStockOut.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, -1, False, False, True)
            Return Me.objStockOut.Primary
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
            Obj.GetDataInput(TBLData, InputType, False, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Out_Adjust, Me.dgvItem, Me.TxtPrimaryStockIn.Text, Me.UserCode)
            'Obj.GetDataInput(TBLData, InputType, False, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_Out_Adjust)

            'ij.BackColorAll = Me.BackColorAll
            'ij.UserCode = Me.UserCode
            'ij.ParentForm1 = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Location_Transfer_Source
            'ij.ShowDialog()

            If InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add Then

            ElseIf InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Replace Then
                If TBLData.Rows.Count > 0 Then
                    Me.dgvItem.Rows.Clear()
                End If
            End If

            For Each rowData In TBLData
                With rowData
                    Index = Me.GetIndex2(.Primary_Product, .Batch, -1, -1, .Primary_Location, .Location_Controlled)
                    'cari data yang sama persis
                    If Index <> -1 Then
                        'tambahkan nilai qty
                        Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                        row.Cells(5).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(5).Value) + CDec(.Qty), ._Decimal)
                    Else
                        'tambahkan baris baru
                        Dim ExpiredDateStr As String
                        If .Expired_Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                        End If
                        Me.dgvItem.Rows.Add(Me.dgvItem.Rows.Count + 1, "", "", .Code, .Description, Me.ObjCD.DigitGrouping(.Qty, ._Decimal), _
                        .Unit, "-", .Unit, .Batch, _
                        ExpiredDateStr, .Location, .Warehouse, .Remarks, _
                        -1, .Primary_Product, .Primary_Location, .Primary_Unit, _
                        -1, -1, 0)
                    End If
                End With
            Next
            Me.ObjCD.SetNumber(Me.dgvItem)
            'format dgv
            Me.FormatDGV()

            'fokuskan ke baris terakhir
            If Me.dgvItem.Rows.Count > 0 Then
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                         Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsm_AddByBarcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsm_AddByBarcode.Click
        Dim Obj As New CustomSMESS.ClsCustomBarcode
        Dim TBLData As New Core_Module.DseBarcodeInput.DataTable1DataTable
        Dim InputType As CustomSMESS.ClsCustomBarcode.EnumInputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add
        Dim Index As Integer = -1
        Dim rowData As Core_Module.DseBarcodeInput.DataTable1Row
        Try
            Obj.GetDataInput(TBLData, InputType, False, CustomSMESS.ClsCustomBarcode.EnumTransaksiAcuan.OutAdjust)

            If InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add Then

            ElseIf InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Replace Then
                If TBLData.Rows.Count > 0 Then
                    Me.dgvItem.Rows.Clear()
                End If
            End If

            For Each rowData In TBLData
                With rowData
                    Index = Me.GetIndex2(.Primary_Product, .Batch, -1, -1, .Primary_Location, .Location_Controlled)
                    'cari data yang sama persis
                    If Index <> -1 Then
                        'tambahkan nilai qty
                        Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                        row.Cells(5).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(5).Value) + CDec(.Qty), ._Decimal)
                    Else
                        'tambahkan baris baru
                        Dim ExpiredDateStr As String
                        If .Expired_Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                        End If
                        Me.dgvItem.Rows.Add(Me.dgvItem.Rows.Count + 1, "", "", .Code, .Description, Me.ObjCD.DigitGrouping(.Qty, ._Decimal), _
                        .Unit, "-", .Unit, .Batch, _
                        ExpiredDateStr, .Location, .Warehouse, .Remarks, _
                        -1, .Primary_Product, .Primary_Location, .Primary_Unit, _
                        -1, -1, 0)
                    End If
                End With
            Next
            Me.ObjCD.SetNumber(Me.dgvItem)
            'format dgv
            Me.FormatDGV()

            'fokuskan ke baris terakhir
            If Me.dgvItem.Rows.Count > 0 Then
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                         Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsm_AddByBatchNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsm_AddByBatchNo.Click
        Try
            Me.AddByBatchNo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tsm_ByOnHand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsm_ByOnHand.Click
        Try
            Me.QuickClear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cms1_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cms1.Opening
        Try
            Me.tsm_ByOnHand.Enabled = Me.QuickAddByProdLevel
            Me.tsm_AddByRefDoc.Enabled = Me.QuickAddByRefDoc
            Me.tsm_AddByBarcode.Enabled = Me.QuickAddByBarcode
            Me.tsm_AddByBatchNo.Enabled = Me.QuickAddByBatchNo
            Me.tsm_ByOnHand.Visible = False
            Me.tsm_AddByBatchNo.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex2(ByVal vPrimaryProduct As Double, ByVal BatchNo As String, _
      ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal vPrimaryLocation As Double, ByVal vLocationControlled As Integer) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows

                If vLocationControlled = 1 Then
                    If IsNumeric(row.Cells(7).Value) = True Then
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo  Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(7).Value) = False Then
                        If row.Cells(15).Value = vPrimaryProduct And _
                        row.Cells(9).Value = BatchNo Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(15).Value = vPrimaryProduct And _
                        row.Cells(9).Value = BatchNo  Then
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
        Dim i As New CustomSMESS.FrmTransaksiStockOutAdj_AddByCustomProduct
        Try
            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TanggalTransaksi = Me.dtpDate.Value
            i.PrimaryStockIn = Me.TxtPrimaryStockIn.Text
            i.TblItemCustom = Me.TblItemCustom
            i.baru = Me.baru
            Me.FormatDGV()
            If Me.TxtRemark.Enabled = False Then
                i.AktifForm = False
            Else
                i.AktifForm = True
            End If

            If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblItemCustom = i.TblItemCustom
                Me.IsikanDgvFromTblDummy()
                Me.FormatDGV()
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturHakAksesFormCustom()
        Dim ObjUser As New Core_Module.ClsUser
        Try
            Me.llbAdvance.Parent = Me.gboItem
            Me.llbAdvance.BringToFront()
            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockOutAdj_AddByCustomProduct") = True Then
                'Me.cms1.Visible = False
                Me.dgvItem.ReadOnly = True
                Me.dgvItem.AllowUserToDeleteRows = False
                Me.TxtCodeNo.Enabled = False
                Me.BtnFindCodeNo.Enabled = False
                Me.txtDesc.Enabled = False
                Me.cboBatchNo.Enabled = False
                Me.DtpExpired.Enabled = False
                Me.cboLocation.Enabled = False
                Me.TxtWarehouse.Enabled = False
                Me.TxtQtyperBatch.Enabled = False
                Me.txtQtyOnHand.Enabled = False
                Me.TxtRemarkItem.Enabled = False
                Me.llbAdvance.Visible = True
            Else
                Me.llbAdvance.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.gboItem.Enabled = True
                'Me.cms1.Visible = True
                Me.dgvItem.AllowUserToDeleteRows = True

                Me.TxtCodeNo.Enabled = True
                Me.BtnFindCodeNo.Enabled = True
                Me.txtDesc.Enabled = True
                Me.cboBatchNo.Enabled = True
                Me.DtpExpired.Enabled = True
                Me.cboLocation.Enabled = True
                Me.TxtWarehouse.Enabled = True
                Me.TxtQtyperBatch.Enabled = True
                Me.txtQtyOnHand.Enabled = True
                Me.TxtRemarkItem.Enabled = True
            End If
            Me.llbAdvance.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsikanDgvFromTblDummy() As Boolean
        Dim Row As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemRow
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim NoPDO, Location, Warehouse As String
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ExpDate As String = ""
        Dim Idx As Integer = 0
        Try
            dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgvItem.Rows.Clear()
            'DefPrec = ObjCur.GetDecimal_DefaultCurrency

            'kosongkan dgv
            dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgvItem.Rows.Clear()
            For Each Row In TblItemCustom.Rows
                Idx = Me.GetIndex(Row.Primary_Product, Row.Primary_Storage_Area, Row.Batch_No, _
           -1, -1, Row.Expiry_Date)

                If Row.Primary_PdO = -1 Then
                    NoPDO = ""
                End If
                If Row.Primary_Storage_Area = -1 Then
                    Location = ""
                    Warehouse = ""
                Else
                    Location = Row.Storage_Area_Code
                    Warehouse = Row.Warehouse_Code
                End If
                If Row.Expiry_Date = #1/1/1950# Then
                    ExpDate = "-"
                Else
                    ExpDate = Format(Row.Expiry_Date, "dd MMM yyyy")
                End If
                Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(Row.Primary_Product)
                Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Row.Primary_Product, Row.Primary_Multi_Units)

                If Idx <> -1 Then
                    rowDGV = dgvItem.Rows(Idx)
                    If Row.Is_Kain = True Then
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(rowDGV.Cells(Me.ColQty.Index).Value) + Row.Qty, DefPrecDefault)
                    Else
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(rowDGV.Cells(Me.ColQty.Index).Value) + Row.Qty, DefPrecDefault)
                    End If
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColRemark.Index).Value) = False Then
                        Dim R As String = ""
                        If CStr(rowDGV.Cells(Me.ColRemark.Index).Value).EndsWith(",") = False Then
                            'tambahkan titik bila tidak diakhiri dengan titik
                            rowDGV.Cells(Me.ColRemark.Index).Value = CStr(rowDGV.Cells(Me.ColRemark.Index).Value) + ", "
                        End If
                        rowDGV.Cells(Me.ColRemark.Index).Value = CStr(rowDGV.Cells(Me.ColRemark.Index).Value) + Row.Remarks
                    Else
                        rowDGV.Cells(Me.ColRemark.Index).Value = Row.Remarks
                    End If
                Else
                    dgvItem.Rows.Add()
                    rowDGV = dgvItem.Rows(dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPONoDONo.Index).Value = Row.po_so_number
                        .Cells(Me.ColSuppDONo.Index).Value = Row.Supp__DO_No
                        .Cells(Me.ColCodeNo.Index).Value = Row.Product_Code
                        .Cells(Me.ColDescription.Index).Value = Row.Description

                        .Cells(Me.ColBatch.Index).Value = Row.Batch_No
                        .Cells(Me.ColExpiredDate.Index).Value = ExpDate
                        .Cells(Me.ColLocation.Index).Value = Location
                        .Cells(Me.ColWareHouse.Index).Value = Warehouse
                        .Cells(Me.ColRemark.Index).Value = Row.Remarks
                        .Cells(Me.ColPrimaryPOnoDoNo.Index).Value = Row.Primary_PO_Purchase
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Row.Primary_Product
                        .Cells(Me.ColPrimaryLocation.Index).Value = Row.Primary_Storage_Area
                        .Cells(Me.ColPrimaryUnits.Index).Value = Row.Primary_Units
                        .Cells(Me.ColPrimaryMultiUnit.Index).Value = Row.Primary_Multi_Units
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = Row.Primary_Secondary_Units
                        .Cells(Me.ColPrimaryPdO.Index).Value = Row.Primary_PdO

                        If Row.Is_Kain = True Then
                            .Cells(Me.ColIsKain.Index).Value = 1
                            .Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(Row.Qty, DefPrecDefault)
                            .Cells(Me.ColUnit.Index).Value = Row.unit
                            .Cells(Me.ColSecondaryQty.Index).Value = -1
                            .Cells(Me.ColSecondaryUnit.Index).Value = ""
                        Else
                            .Cells(Me.ColIsKain.Index).Value = 0
                            .Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(Row.Qty, DefPrecDefault)
                            .Cells(Me.ColUnit.Index).Value = Row.unit
                            .Cells(Me.ColSecondaryQty.Index).Value = -1
                            .Cells(Me.ColSecondaryUnit.Index).Value = ""
                        End If
                    End With
                End If
            Next
            ObjCD.SetNumber(dgvItem)
            dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class