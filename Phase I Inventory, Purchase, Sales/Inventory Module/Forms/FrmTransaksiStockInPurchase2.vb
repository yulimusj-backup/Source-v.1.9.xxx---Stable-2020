Option Explicit On
Option Strict Off
Imports Core_Module
Public Class FrmTransaksiStockInPurchase2

    Public UserCode As String = ""
    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
    Public ObjCD As New ClsCekData
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objStockIn As New clsTransaksiStockInOut
    Dim ObjProduct As New clsMasterProduct
    Dim ObjStockInOutItem As New clsTransaksiStockInOutItem
    Dim ObjSupplier As New clsMasterSupplier
    Dim ObjItemInOut As New clsTransaksiStockInOut
    Dim ObjItemBundling As New clsTransaksiStockBundlingItem
    Dim objPOItem As New ClsTransaksiPOItem
    Dim objPackingListItem As New ClsTransaksiPackingListItem
    Dim ObjPackingList As New ClsTransaksiPackingList
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjUser As New ClsUser
    Dim objPO As New ClsTransaksiPO
    Dim ObjLicense As New ClsLicense
    Dim ObjPurcReq As New ClsTransaksiPR
    Dim ObjNumb As New ClsMasterTransactionNumbering
    Dim TBL_PickUpOrderQty As New DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item2DataTable

    ''@rosadinnn
    ''@MUaT
    ''@update tableAdapter yg mengacu ke view, buat tableAdapter baru
    Dim TBL As New DseTransaksiPO.orderbalance_report_view0MultiDataTable
    Dim TBL2 As New dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiDataTable
    Dim TBL3 As New DseTransaksiPO.orderbalance_report_view0MultiDataTable
    Dim TBL4 As New dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiDataTable
    Dim TBL5 As New DseTransaksiPackingList.tbltrans_packinglist_view1MultiDataTable

    Dim buffer As String = ""
    Dim CekCbo As Boolean = False

    Public TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

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


            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, False, False)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, False, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, False, False)
            Me.ChkIgnoreTolerance.Enabled = True

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.SetInputCustom()
            If Me.TxtPrimaryPihakAcuan.Text <> "-1" And Me.TxtPrimaryPihakAcuan.Text <> "" Then
                Me.SetCBO()
            End If
            Me.cboCodeNo.Enabled = True
            Me.cboPurchaseOrder.Enabled = True
            Me.dgvPackingList.ReadOnly = True

            If Me.baru = False Then
                If Me.TxtPrimaryPackingList.Text <> "-1" Then
                    Me.TxtKodePihakAcuan.ReadOnly = True
                    Me.TxtKodePihakAcuan.BackColor = Me.TxtStockInNo.BackColor
                    Me.BtnCariPihakAcuan.Enabled = False
                    Me.TxtPackingListNo.ReadOnly = False
                    Me.TxtPackingListNo.BackColor = Me.TxtPrimaryPackingList.BackColor
                    Me.btnFindPackingList.Enabled = True
                Else
                    Me.TxtKodePihakAcuan.ReadOnly = False
                    Me.TxtKodePihakAcuan.BackColor = Me.TxtPrimaryPackingList.BackColor
                    Me.BtnCariPihakAcuan.Enabled = True
                    Me.TxtPackingListNo.ReadOnly = True
                    Me.TxtPackingListNo.BackColor = Me.TxtStockInNo.BackColor
                    Me.btnFindPackingList.Enabled = False
                End If
            Else
                Me.TxtKodePihakAcuan.ReadOnly = False
                Me.BtnCariPihakAcuan.Enabled = True
                Me.TxtKodePihakAcuan.BackColor = Me.TxtPrimaryStockIn.BackColor
                Me.TxtPackingListNo.ReadOnly = False
                Me.btnFindPackingList.Enabled = True
                Me.TxtPackingListNo.BackColor = Me.TxtPrimaryStockIn.BackColor
            End If

            Me.SetGboItem()
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetInputCustom()
        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockInPurchase_Custom") = False Then
                Me.SplitContainer2.Enabled = True
                'Me.ToolStripMenuItem1.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
                Me.LinkLabel1.Enabled = False
            Else
                Me.dgvItem.ReadOnly = True
                Me.SplitContainer2.Enabled = False
                'Me.ToolStripMenuItem1.Visible = True
                Me.dgvItem.AllowUserToDeleteRows = False
                Me.LinkLabel1.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)

            ''@rosadinnn
            ''@MUaT
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, False, False)

            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, False, False)

            ''@rosadinnn
            ''@MUAT
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, False, False)
            Me.ChkIgnoreTolerance.Enabled = False

            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.cboCodeNo.Enabled = False
            Me.cboPurchaseOrder.Enabled = False
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
            Me.TxtPrimaryPihakAcuan.Text = "-1"
            Me.TxtPrimaryPackingList.Text = "-1"
            Me.txtPrimaryPDO.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.dgvPackingList.Rows.Clear()
            Me.TblItemCustom.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)

            ''@rosadinnn
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.ChkIgnoreTolerance.Checked = False

            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.txtSecondaryQty.Visible = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnits.Left = Me.lblSecondaryUnits.Left
            'Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
            Me.lblSecQty.Visible = False
            Me.lblExpiryDate.Visible = False
            Me.DtpExpired.Value = Today
            Me.DtpExpired.Visible = True
            Me.lblLocation.Visible = False
            Me.lblWarehouse.Visible = False
            Me.TxtLocation.Visible = True
            Me.TxtWarehouse.Visible = True
            Me.btnFindLocation.Visible = True
            Me.lblBatchNo.Visible = False
            Me.txtBatchNo.Visible = True
            Me.txtBatchNo.TabStop = True
            Me.txtBatchNo.ReadOnly = False
            Me.cboCodeNo.DataSource = Nothing
            Me.cboPurchaseOrder.DataSource = Nothing
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnits.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimaryPDO.Text = "-1"
            Me.txtPrimaryPO.Text = "-1"
            Me.TxtQtyperBatch.Text = "0"
            Me.txtSecondaryQty.Text = "-1"
            Me.TxtOrderBalance.Text = "0"

            ''@rosadinnn
            ''@MUaT
            Me.TxtPrimaryUnitDefault.Text = "-1"
            Me.TxtQtyDefault.Text = "0"
            Me.LblUnitDefault.Text = "Units"
            Me.TxtRatioReference.Text = "1"
            Me.TxtRatioDefault.Text = "1"
            Me.LblQty.Text = "Qty"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.cboCodeNo.Focus()
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
                If Me.txtBatchNo.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtBatchNo.Focus()
                    Return False
                End If
            End If

            '4. nilai qtyDefault <= 0
            

            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyDefault.Text) = False Then
                MsgBox("Qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatch.Focus()
                Return False
            End If

            '4.a Untuk transaksi proses (in/back/out proses) no Pdo harus diisikan
            '4.b Untuk transaksi yang berkaitan dengan purchase (stock out purchase return dan stout in purchase)
            '    harus menyertakan no PO bila terhubungkan dengan software purchase
            '4.c Untuk transaksi yang berkaitan dengan sales (stock out sales dan stock in sales return) harrus
            '    menyertakan no DO bila terhubungkan dengan software sales
            '5. Bila location controlled, kode lokasi tidak diisikan 
            '5. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlled.Checked = True Then
                If Me.txtPrimaryLocation.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtLocation.Focus()
                    Return False
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
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.txtBatchNo.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColBatch.Index).Value = Me.txtBatchNo.Text And Me.txtBatchNo.Text <> "" And row.Cells(Me.ColPrimaryCodeNo.Index).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColExpiredDate.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.txtBatchNo.Text + " on " + _
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
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text, _
                    JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> Me.DtpExpired.Value Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNo.Text) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.DtpExpired.Focus()
                            Return False
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpired.Value Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNo.Text) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpired.Focus()
                        Return False
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")
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

            '9. Cek berkaitan dengan purchase order balance
            Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlDGVAll As Decimal
            If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)

                ''@rosadinnn
                ''@MUaT
                Dim ToleranceUp As Decimal = objPO.GetToleranceUp(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
                JmlPOB += ToleranceUp

            Else
                JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
            End If

            JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
            Me.txtPrimaryPO.Text)
            Dim SecQty As Decimal = 0
            If IsNumeric(Me.txtSecondaryQty.Text) = True Then
                SecQty = Me.txtSecondaryQty.Text
            End If
            Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
            Me.ObjProduct.GetProductProperties(Me.TxtPrimaryCodeNo.Text, LocationCont, BatchNoCont, _
            ExpiryDateCont, UseSecondary)
            JmlDGVAll = Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)

            ''@rosadinnn
            ''@MUaT
            Dim MaxCanInput As Decimal = 0
            If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGVAll - CDec(Me.TxtQtyperBatch.Text)
                MaxCanInput = JmlPOB + JmlObjLama - JmlDGVAll
            Else
                JmlPOBBaru = JmlPOB - JmlDGVAll - CDec(Me.TxtQtyperBatch.Text)
                MaxCanInput = JmlPOB - JmlDGVAll
            End If

            If JmlPOBBaru < 0 Then
                'MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlPOB + JmlObjLama - JmlDGVAll)) + _
                '" " + CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", _
                'MsgBoxStyle.Information, "Information")

                ''@rosadinnn
                ''@MUaT
                MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(MaxCanInput), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)) + _
                        " " + CStr(Me.lblUnits.Text) + ".", MsgBoxStyle.Information, "Information")

                Me.TxtQtyperBatch.Focus()
                Return False
            End If

            'cek bila purchase order belum diotorisasi maka tidak bisa masuk
            If Me.objPO.GetAuthorizationLevel(Me.txtPrimaryPO.Text) <> Me.objPO.GetAuthorizationLevelNeeded(Me.txtPrimaryPO.Text) Then
                MsgBox("This purchase order require more authorization, please check authorization needed.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'untuk po yang masuk, harus sama currency-nya
            If Me.dgvItem.RowCount > 0 Then
                Dim mTaxNew, mprimaryCurrency, mTaxOld, mCurrencyOld, vPrimaryPO_Old As Decimal
                vPrimaryPO_Old = Me.dgvItem.Rows(0).Cells(Me.ColPrimaryPOno.Index).Value

                Me.objPO.GetPrimaryCurrency_Tax(Me.txtPrimaryPO.Text, mTaxNew, mprimaryCurrency)
                Me.objPO.GetPrimaryCurrency_Tax(vPrimaryPO_Old, mTaxOld, mCurrencyOld)

                '   a. % pajak yang digunakan sama
                If mTaxNew <> mTaxOld Then
                    MsgBox("This purchase order item cannot be in one stock in purchase with the existing items. " + _
                    "For several purchase order item written in one stock in purchase their currency and tax must be identical", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '   b. primary currency sama
                If mprimaryCurrency <> mCurrencyOld Then
                    MsgBox("This purchase order item cannot be in one stock in purchase with the existing items. " + _
                    "For several purchase order item written in one stock in purchase their currency and tax must be identical", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Decimal, ByVal Secondary_Qty As _
    Decimal, ByVal vPrimaryLocation As Decimal, ByVal vBatchNo As String) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    OK = True
                    'cek secondary qty
                    'If OK = True Then
                    '    If row.Cells(Me.ColSecondaryQty.Index).Value <> Secondary_Qty Then
                    '        OK = False
                    '    End If
                    'End If
                    'cek location
                    If OK = True Then
                        If vPrimaryLocation <> row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If vBatchNo <> row.Cells(Me.ColBatch.Index).Value Then
                        OK = False
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

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Decimal, ByVal Secondary_Qty As _
   Decimal, ByVal vPrimaryLocation As Decimal, ByVal vBatchNo As String, ByVal vExcludeIndex As Integer) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Index <> vExcludeIndex Then
                    OK = True
                    'cek secondary qty
                    'If OK = True Then
                    '    If row.Cells(Me.ColSecondaryQty.Index).Value <> Secondary_Qty Then
                    '        OK = False
                    '    End If
                    'End If
                    'cek location
                    If OK = True Then
                        If vPrimaryLocation <> row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If vBatchNo <> row.Cells(Me.ColBatch.Index).Value Then
                        OK = False
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
    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Decimal, ByVal Secondary_Qty As _
    Decimal, ByVal vPrimaryLocation As Decimal, ByVal vBatchNo As String, _
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

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryPO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO Then
                    Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimaryPO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO Then
                    Jml = Jml + CDec(row.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim ObjPickUp As New ClsTransaksiPR_PickUpOrder
        Dim rowPickUp As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item2Row
        Dim Dummy As String = ""
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
            KoleksiNamaData.Add("stock in number")
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

            'cek supplier jangan sampai tidak diisikan
            If Me.TxtPrimaryPihakAcuan.Text = "-1" Then
                MsgBox("Please specify supplier's code", MsgBoxStyle.Information, "Information")
                Me.TxtKodePihakAcuan.Focus()
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
                If Me.objStockIn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objStockIn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtStockInNo.Text = Me.objStockIn.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objStockIn.TanggalTransaksi
                        Me.TxtStockInNo.Text = Me.objStockIn.NoTransaksi
                        Return False
                    End If
                End If
            End If

            'cek kecocokan dengan packing list
            If Me.TxtPrimaryPackingList.Text <> "-1" Then
                If Me.KonfirmasiJumlah() = False Then
                    Return False
                End If
            End If

            'kalo no do supplier kosong maka kosongkan data
            Me.TxtSuppDONumber.Text = Me.TxtSuppDONumber.Text.Trim
            If Me.TxtSuppDONumber.Text = "" Then
                MsgBox("Please specify supplier do number", MsgBoxStyle.Information, "Information")
                Me.TxtSuppDONumber.Focus()
                Return False
            End If

            'jika bukan baru, bandingkan minimum qty pick order dengan minimum qty pada dgv
            If Me.baru = False Then
                ObjPickUp.GetDetailPickUpQtyByPrimaryStockIn(Me.TxtPrimaryStockIn.Text, Me.TBL_PickUpOrderQty)
                For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                    Me.TBL_PickUpOrderQty.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
                    If Me.TBL_PickUpOrderQty.DefaultView.Count > 0 Then
                        rowPickUp = TBL_PickUpOrderQty.DefaultView(0).Row
                        rowPickUp.qty -= CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                        If rowPickUp.qty <= 0 Then
                            TBL_PickUpOrderQty.Rows.Remove(rowPickUp)
                        End If
                    End If
                Next
                If Me.TBL_PickUpOrderQty.Rows.Count > 0 Then
                    For Each rowPickUp In TBL_PickUpOrderQty.Rows
                        If Dummy = "" Then
                            Dummy = Me.ObjProduct.GetProductCode(rowPickUp.primary_code_no) + "(" + Me.ObjCD.DigitGrouping(rowPickUp.qty, _
                            Me.ObjProduct.GetAngkaDecimal(rowPickUp.primary_code_no)) + " " + Me.ObjProduct.GetUnitCode(rowPickUp.primary_code_no) + ")"
                        Else
                            Dummy = Dummy + ", " + Me.ObjProduct.GetProductCode(rowPickUp.primary_code_no) + "(" + Me.ObjCD.DigitGrouping(rowPickUp.qty, _
                            Me.ObjProduct.GetAngkaDecimal(rowPickUp.primary_code_no)) + " " + Me.ObjProduct.GetUnitCode(rowPickUp.primary_code_no) + ")"
                        End If
                    Next
                    MsgBox("The following item is insufficient for the pick up order generated by this transaction: " + Dummy + ". Please add more as specified.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Dim objLocationAccess As New FrmUserLocationPrivilegeSetting
            'Dim listOfPrimaryLocation As New List(Of Double)
            'For Each dataGridViewRow As DataGridViewRow In Me.dgvItem.Rows
            '    If dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value <> "-1" Then
            '        If listOfPrimaryLocation.Contains(dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value) = False Then listOfPrimaryLocation.Add(dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value)
            '    End If
            'Next
            'If objLocationAccess.HasAccessPrivileges(Me.UserCode, listOfPrimaryLocation, FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut.STOCK_IN_PURCAHSE) = False Then
            '    Return False
            'End If

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_IN_PURCAHSE, Me.TBLDetailLocation) = False Then
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
                    Dim PrimaryPOSO As Double = .Cells(Me.ColPrimaryPOno.Index).Value
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
            ObjPickUp = Nothing
            ObjAcc = Nothing
        End Try
    End Function

    Private Function KonfirmasiJumlah() As Boolean
        Dim Warning As String = ""
        Dim Unit As String = ""
        Dim Qty As Integer = 0
        Dim JmlIn As Decimal = 0

        Try
            'cek kembali jumlah, jika ada yang kurang maka konfirmasikan
            Dim rowPL As DataGridViewRow
            For Each rowPL In Me.dgvPackingList.Rows
                JmlIn = Me.GetJmlDGV_Multi(rowPL.Cells(Me.colPrimaryCodeNoDGVPAckingList.Index).Value, _
                rowPL.Cells(Me.ColPrimaryPODGVPackingLIst.Index).Value)

                ''@rosadinnn
                ''@MUaT
                If rowPL.Cells(Me.ColQtyMulti_PL.Index).Value <> JmlIn Then
                    Qty += 1
                    Unit = rowPL.Cells(Me.ColUnitMulti_PL.Index).Value

                    If Warning = "" Then
                        Warning = "There several items where packing list quantity don't match the stock in quantity. They are:" + vbCrLf
                    End If
                    Warning = CStr(Warning) + CStr(Qty) + ". Code No: " + rowPL.Cells(Me.colCodeNoDGVPackingList.Index).Value.ToString + _
                    " PO number: " + rowPL.Cells(Me.colPONoDGVPackingList.Index).Value.ToString + _
                    " --> Packing List Qty: " + rowPL.Cells(Me.ColQtyMulti_PL.Index).Value.ToString + " " + Unit + " " + _
                    "- Stock In Qty: " + JmlIn.ToString + " " + Unit + " " + vbCrLf
                End If
            Next

            If Warning <> "" Then
                Warning = Warning + vbCrLf + "Click ""Yes"" if you aware of this and the data is right."
                If MsgBox(Warning, MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.objStockIn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objStockIn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtStockInNo.Text = Me.objStockIn.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtStockInNo.Text = Me.objStockIn.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtStockInNo.Text = Me.objStockIn.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objStockIn
                .NoTransaksi = Me.TxtStockInNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Remarks = Me.TxtRemark.Text
                .PrimarySupplier = Me.TxtPrimaryPihakAcuan.Text
                .PrimaryGIT = Me.TxtPrimaryPackingList.Text
                .SupplierDoNo = Me.TxtSuppDONumber.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objStockIn
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.TxtPrimaryPackingList.Text = .PrimaryGIT
                Me.TxtSuppDONumber.Text = .SupplierDoNo
                Me.TxtPackingListNo.Text = Me.ObjPackingList.GetNoTransaksi(Me.TxtPrimaryPackingList.Text)
                If Me.TxtPrimaryPackingList.Text <> "-1" Then
                    Me.IsikanDGVPackingLIst(Me.dgvPackingList, Me.TxtPrimaryPackingList.Text)
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.TxtPrimaryPihakAcuan.Text = .PrimarySupplier
                Me.TxtKodePihakAcuan.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
            End With
            Me.FormatDGV()
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                'row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
                'If IsNumeric(row.Cells(7).Value) = True Then
                '    row.Cells(7).Value = Me.ObjCD.DigitGrouping(row.Cells(7).Value)
                'End If

                ''@rosadinnn
                ''@MUaT
                row.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyMulti.Index).Value)
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    row.Cells(Me.ColSecondaryQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecondaryQty.Index).Value)
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
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            Me.dgvPackingList.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.KosongkanField()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.ChkIgnoreTolerance.TabStop = False
            Me.CekCbo = True
            Me.ShowIcon = False
            Me.TxtSuppDONumber.MaxLength = 50
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockIn.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                            'konfirmasi proses selesai
                            Me.SalinKeField()
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            Me.KosongkanItem()

                            Me.BuatPickUpOrder()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockIn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            Me.KosongkanItem()

                            Me.BuatPickUpOrder()
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

    Private Sub BuatPickUpOrder()
        Dim AdaPurchReq As Boolean = False
        Dim row_ProductQty As DseTransaksiShippingForecast.tblmaster_productqtyRow
        Dim ObjPickUp As New ClsTransaksiPR_PickUpOrder

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                If ObjPurcReq.GetOrderBalance_Only(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) > 0 Then
                    AdaPurchReq = True
                    Exit For
                End If
            Next

            If AdaPurchReq = True Then
                Dim ijkl As New FrmTrans_MaterialPlanningDashboard
                ijkl.vPrimaryStockInPurchaseReference = Me.TxtPrimaryStockIn.Text
                ijkl.UserCode = Me.UserCode
                ijkl.vFromStockIn = True
                'isikan item stock in, yang di anggap stock in purchase
                ijkl.TBL_ProductQty1.Rows.Clear()
                For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                    ijkl.TBL_ProductQty1.DefaultView.RowFilter = "[Primary Product] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString
                    If ijkl.TBL_ProductQty1.DefaultView.Count = 0 Then
                        row_ProductQty = ijkl.TBL_ProductQty1.Newtblmaster_productqtyRow
                        row_ProductQty.Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                        row_ProductQty.Qty = rowDGV.Cells(Me.ColQty.Index).Value
                        row_ProductQty.Primary_Unit = rowDGV.Cells(Me.ColPrimaryUnits.Index).Value
                        ijkl.TBL_ProductQty1.Rows.Add(row_ProductQty)
                    Else
                        row_ProductQty = ijkl.TBL_ProductQty1.DefaultView(0).Row
                        row_ProductQty.Qty += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                    End If
                Next
                'kurangkan dengan item pick order yang sudah ada stock out
                ObjPickUp.GetDetailPickUpQtyByPrimaryStockIn(Me.TxtPrimaryStockIn.Text, Me.TBL_PickUpOrderQty)
                For Each rowPickUp As DseTransaksiPR_PickUpOrder.tbltrans_purchase_req_pickup_order_item2Row In Me.TBL_PickUpOrderQty.Rows
                    ijkl.TBL_ProductQty1.DefaultView.RowFilter = "[Primary Product] = " + rowPickUp.primary_code_no.ToString
                    If ijkl.TBL_ProductQty1.DefaultView.Count = 0 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        row_ProductQty = ijkl.TBL_ProductQty1.DefaultView(0).Row
                        row_ProductQty.Qty -= rowPickUp.qty
                        If row_ProductQty.Qty = 0 Then
                            ijkl.TBL_ProductQty1.Rows.Remove(row_ProductQty)
                        ElseIf row_ProductQty.Qty < 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                Next
                ijkl.ShowDialog()
                ijkl = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjPickUp = Nothing
        End Try
    End Sub

    Private Function GetJml_DGVItem3(ByVal vPrimaryProduct As Decimal, _
    ByVal vPrimaryPO As Decimal) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim Dummy As Decimal = 0
        Try
            For Each rowDGV In Me.dgvItem3.Rows
                If rowDGV.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO And _
                rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Dummy += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.dtpDate.Value.Date = Today And _
            Me.txtBatchNo.Text = "" And Me.TxtLocation.Text = "" And _
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
                        Me.objStockIn.CariDetail(Me.TxtPrimaryStockIn.Text, Me.dgvItem, Me.TblItemCustom)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanItem()
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
            Me.objStockIn.ResetObject(Me.dgvItem)
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            Me.TxtStockInNo.Text = Me.objStockIn.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiStockInOut_Purchase

        Try
            ij.JenisTransaksi = Me.TipeTransaksi
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()

            If Not ij.PrimaryTransaksiStockInOut = -1 Then
                If Me.objStockIn.CariDetail(ij.PrimaryTransaksiStockInOut, Me.dgvItem, Me.TblItemCustom) = True Then
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
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim ObjPickUp As New ClsTransaksiPR_PickUpOrder
        Dim RetVal As String = ""
        Dim RetVal1 As String = ""
        Dim Lanjut As Boolean = False

        Try
            Me.baru = True

            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                If Me.objStockIn.ApakahBolehEdit(Me.UserCode) Then
                    'jika ada pick up order yang merujuk pada transaksi ini, maka berikan peringatan bahwa pick up order ybs juga akan dihapuskan
                    RetVal = ObjPickUp.GetNoTransaksi_YangBelumStockOut_ByPrimaryStockIn(Me.TxtPrimaryStockIn.Text)
                    RetVal1 = ObjPickUp.GetNoTransaksi_YangSudahStockOut_ByPrimaryStockIn(Me.TxtPrimaryStockIn.Text)
                    If RetVal <> "" Or RetVal1 <> "" Then
                        If RetVal <> "" Then
                            If MsgBox("The following pick up order refer to this transaction, editing this transaction will automatically delete pick up order: " + RetVal + "." + _
                            vbCrLf + "Continue edit process anyway?", _
                            MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Lanjut = True
                            Else
                                Lanjut = False
                            End If
                        Else
                            If MsgBox("The following pick up order refer to this transaction and already stock out: " + RetVal1 + "." + _
                            vbCrLf + "Minimum qty is restricted by the pick up qty, continue edit process anyway?", _
                            MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Lanjut = True
                            Else
                                Lanjut = False
                            End If
                        End If
                    Else
                        Lanjut = True
                    End If
                    If Lanjut = True Then
                        Me.baru = False
                        Me.AktifkanDetail()
                        Me.dgvItem.Focus()
                        'nilai min juga harus diatas tanggal pengkuncian data terakhir
                        'msgbox("belum")
                        ObjPickUp.Hapus_ByPrimaryStockIn(Me.TxtPrimaryStockIn.Text)
                        Me.TxtKodePihakAcuan.ReadOnly = True
                        Me.TxtKodePihakAcuan.BackColor = Me.txtReadOnly.BackColor
                        Me.BtnCariPihakAcuan.Enabled = False
                        Me.TxtPackingListNo.ReadOnly = True
                        Me.TxtPackingListNo.BackColor = Me.txtReadOnly.BackColor
                        Me.btnFindPackingList.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjPickUp = Nothing
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim ObjPickUp As New ClsTransaksiPR_PickUpOrder
        Dim PickUpOrderNo As String = ""

        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                'tidak cek stock minus ketika cek apakah boleh hapus, karena akan di cek tepat sebelum di hapus,
                'pengecekan ada dalam fungsi hapus di kelas, sehingga memperhitungkan antrian
                If Me.objStockIn.ApakahBolehHapus(Me.UserCode, True, True, False, False) = False Then
                    Me.btnFind.Focus()
                Else
                    'cek apakah ada pick up order yang sudah stock out yang 
                    'jika ada, maka tidak boleh hapus stock in
                    PickUpOrderNo = ObjPickUp.GetNoTransaksi_YangSudahStockOut_ByPrimaryStockIn(Me.TxtPrimaryStockIn.Text)
                    If PickUpOrderNo <> "" Then
                        MsgBox("There is pick up order with stock out adjust refering to this transaction, you can't delete this data. The pick up order number(s) is : " _
                        + PickUpOrderNo + ".", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    PickUpOrderNo = ObjPickUp.GetNoTransaksi_YangBelumStockOut_ByPrimaryStockIn(Me.TxtPrimaryStockIn.Text)
                    If PickUpOrderNo = "" Then
                        hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    Else
                        hasil = MsgBox("There is pick up order refering to this transaction, deleting this transaction will also delete the respected pick up order." + _
                        " The pick up order number(s) is : " _
                        + PickUpOrderNo + ". Continue delete process?", MsgBoxStyle.YesNo, "Confirm")
                    End If

                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        'hapus semua pick order yang merujuk pada data ini
                        ObjPickUp.Hapus_ByPrimaryStockIn(Me.TxtPrimaryStockIn.Text)

                        If Me.objStockIn.Hapus(Me.TxtPrimaryStockIn.Text, Me.UserCode, Me.TblItemCustom) = True Then
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
            ObjPickUp = Nothing
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
                            If Me.objStockIn.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom) = True Then
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
                            If Me.objStockIn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) = True Then
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

    ''@rosadinnn
    ''@MUaT
    Private Sub SetGboItem()
        Try
            If Me.TxtPrimaryCodeNo.Text = "" Then
                Me.TxtPrimaryCodeNo.Text = "-1"
            End If
            If Me.txtPrimaryUnits.Text = "" Then
                Me.txtPrimaryUnits.Text = "-1"
            End If
            If Me.txtPrimarySecondaryUnits.Text = "" Then
                Me.txtPrimarySecondaryUnits.Text = "-1"
            End If

            ''@rosadinnn
            ''@MUaT
            Dim i As New TextBox
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
                    Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
                    Me.chkPdoControlled, i, Me.txtPrimarySecondaryUnits)

            Me.txtPrimaryPDO.Text = "-1"
            If Me.chkLocationControlled.Checked = False Then
                Me.txtPrimaryLocation.Text = "-1"
                Me.TxtLocation.Text = ""
                Me.TxtWarehouse.Clear()
                Me.TxtLocation.ReadOnly = True
                Me.lblLocation.Visible = True
                Me.lblWarehouse.Visible = True
                Me.TxtLocation.Visible = False
                Me.TxtWarehouse.Visible = False
                Me.btnFindLocation.Visible = False
            Else
                Me.TxtLocation.ReadOnly = False
                Me.lblLocation.Visible = False
                Me.lblWarehouse.Visible = False
                Me.TxtLocation.Visible = True
                Me.TxtWarehouse.Visible = True
                Me.btnFindLocation.Visible = True
            End If
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.lblExpiryDate.Visible = False
                Me.DtpExpired.Visible = True
                Me.DtpExpired.Value = Today
            Else
                Me.lblExpiryDate.Visible = True
                Me.DtpExpired.Visible = False
                Me.DtpExpired.Value = #1/1/1950#
            End If
            If Me.chkBatchNoControlled.Checked = True Then
                Me.txtBatchNo.ReadOnly = False
                Me.txtBatchNo.TabStop = True
                Me.lblBatchNo.Visible = False
                Me.txtBatchNo.Visible = True
            Else
                Me.txtBatchNo.ReadOnly = True
                Me.txtBatchNo.TabStop = False
                Me.lblBatchNo.Visible = True
                Me.txtBatchNo.Visible = False
                Me.txtBatchNo.Text = ""
            End If
            If Me.chkUseSecondaryUnit.Checked = False Then
                Me.txtSecondaryQty.Visible = False
                Me.lblSecondaryUnits.Visible = False
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
            Else
                Me.txtSecondaryQty.Visible = True
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                Me.txtSecondaryQty.Text = "0"
            End If
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then

                ''@rosadinnn
                ''@MUaT
                'Me.lblUnits.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)

                Me.LblUnits2.Text = Me.lblUnits.Text
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            Else
                Me.lblUnits.Text = "Units"
                Me.LblUnits2.Text = Me.lblUnits.Text
                Me.lblSecondaryUnits.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCBO(Optional ByVal vIndexPO As Double = -1, Optional ByVal vIndexProduct As Double = -1)
        Try
            If Me.SplitContainer2.Enabled = True Then
                Me.cboPurchaseOrder.DataSource = Nothing
                Me.cboCodeNo.DataSource = Nothing
                Me.IsiCboPurchaseOrder(Me.TxtPrimaryPihakAcuan.Text)

                If vIndexPO <> -1 And Me.TBL.DefaultView.Count > vIndexPO Then
                    'geser po ke yg lama
                    Me.cboPurchaseOrder.SelectedIndex = vIndexPO
                    'geser product
                    If vIndexProduct <> -1 And Me.TBL3.DefaultView.Count > vIndexProduct Then
                        Me.cboCodeNo.SelectedIndex = vIndexProduct
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.cboCodeNo.SelectedValue)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.cboCodeNo.SelectedValue)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        Me.TxtQtyperBatch.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.cboCodeNo.SelectedValue <> "" Then
                        i.txtCari.Text = Me.cboCodeNo.SelectedValue
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.cboCodeNo.SelectedValue = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            Me.TxtQtyperBatch.Focus()
                        End If
                    Else
                        '  Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.cboCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As New frmCariMasterProduct

        Try
            If Me.cboCodeNo.SelectedValue <> "" Then
                If Me.ObjProduct.GetJml(Me.cboCodeNo.SelectedValue) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.cboCodeNo.SelectedValue)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        Me.TxtQtyperBatch.Focus()
                    End If
                Else
                    If Me.cboCodeNo.SelectedValue <> "" Then
                        i.txtCari.Text = Me.cboCodeNo.SelectedValue
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.cboCodeNo.SelectedValue = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            Me.TxtQtyperBatch.Focus()
                        End If
                    Else
                        'Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.cboCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Function GetIndex(ByVal vPrimaryProduct As Decimal, ByVal vPrimaryLocation As Decimal, ByVal BatchNo As String, _
    ByVal SecondaryQty As Decimal, ByVal PrimaryPdO As Decimal, ByVal ExpiryDate As Date, ByVal vPrimaryPO As Decimal) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And SecondaryQty = -1 _
                        And row.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# And row.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO Then
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
        Dim rowDgv As DataGridViewRow
        Dim DummyPO As String = ""
        Dim PrimaryLocationLama As Double = -1

        Try
            If Me.ApakahDataItemBenar = True Then
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value, Me.txtPrimaryPO.Text)
                PrimaryLocationLama = Me.txtPrimaryLocation.Text
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)

                    row.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatch.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))
                    Me.SetQtyDefaultOnDgv(row.Index)
                    'row.Cells(5).Value = CDec(row.Cells(5).Value) + CDec(Me.TxtQtyperBatch.Text)

                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Or Me.DtpExpired.Visible = False Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If
                    Dim ProductCode As String
                    Dim DV As Data.DataRowView
                    DV = Me.cboCodeNo.SelectedValue
                    ProductCode = DV.Item(3)

                    ''@rosadinnn
                    ''@MUaT
                    Me.dgvItem.Rows.Add()
                    rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDgv
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPONo.Index).Value = Me.cboPurchaseOrder.Text
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatch.Text
                        .Cells(Me.ColUnitCodeMulti.Index).Value = Me.lblUnits.Text
                        If Me.txtPrimaryUnits.Text <> Me.TxtPrimaryUnitDefault.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                            ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)) _
                            & " " & Me.lblUnits.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                            Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) & " " & Me.LblUnitDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyDefault.Text
                        .Cells(Me.ColUnit.Index).Value = Me.LblUnitDefault.Text
                        .Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                        .Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                        .Cells(Me.ColBatch.Index).Value = Me.txtBatchNo.Text
                        .Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColLocation.Index).Value = Me.TxtLocation.Text
                        .Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryPOno.Index).Value = Me.txtPrimaryPO.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.TxtPrimaryUnitDefault.Text
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                        .Cells(Me.ColPrimaryPdO.Index).Value = Me.txtPrimaryPDO.Text
                        .Cells(Me.ColPrimaryUnitMulti.Index).Value = Me.txtPrimaryUnits.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.TxtRatioDefault.Text
                        .Cells(Me.ColRatioReference.Index).Value = Me.TxtRatioReference.Text
                    End With

                    'Me.dgvItem.Rows.Add("", Me.cboPurchaseOrder.SelectedValue, "", ProductCode, Me.txtDesc.Text, Me.TxtQtyperBatch.Text, _
                    'Me.lblUnits.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.txtBatchNo.Text, _
                    'ExpiredDateStr, Me.TxtLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnits.Text, _
                    'Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)

                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If

                'perhitungan menggunakan nilai jumlah dgv yang baru
                Dim JmlObjLama, JmlPOBBaru, JmlDGV, JmlPOB As Decimal
                JmlDGV = Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                    JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
                Else
                    JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
                End If
                JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)
                JmlPOBBaru = JmlPOB + JmlObjLama - JmlDGV

                'If JmlPOBBaru <= 0 Then
                '    'remove row, pada code no ybs
                '    Dim rowView As Data.DataRowView = Me.cboCodeNo.SelectedItem
                '    Dim row As DseTransaksiPO.orderbalance_report_view0MultiRow = rowView.Row
                '    Me.TBL3.Rows.Remove(row)
                '    Me.cboCodeNo_SelectedValueChanged(sender, e)
                '    Me.cboCodeNo.Focus()
                '    If Me.cboCodeNo.Items.Count = 0 Then
                '        Me.SetGboItem()
                '        Me.cboPurchaseOrder.Focus()
                '        Me.SetCBO()
                '    End If
                'Else
                '    If Me.cboCodeNo.SelectedIndex <> Me.cboCodeNo.Items.Count - 1 Then
                '        Me.cboCodeNo.SelectedIndex += 1
                '        Me.cboCodeNo.Focus()
                '    Else
                '        'focus pada cbo code no
                '        Me.cboCodeNo.Focus()
                '        'hitung kembali order balance dan sisa free qty
                '        Me.SetOrderBalance()
                '    End If
                'End If
                DummyPO = Me.cboPurchaseOrder.Text
                Me.SetCBO(Me.cboPurchaseOrder.SelectedIndex, Me.cboCodeNo.SelectedIndex)
               
                'kosongkan gbo item
                Me.txtBatchNo.Text = ""
                Me.TxtRemarkItem.Text = ""
                Me.TxtQtyperBatch.Text = "0"
                Me.DtpExpired.Value = Today
                If Me.txtPrimarySecondaryUnits.Text = "-1" Then
                    Me.txtSecondaryQty.Text = "-1"
                Else
                    Me.txtSecondaryQty.Text = "0"
                End If
                If Me.TxtPrimaryCodeNo.Text = "-1" Then
                    Me.TxtOrderBalance.Text = "0"
                End If
                Me.txtPrimaryLocation.Text = "-1"
                Me.TxtLocation.Clear()
                Me.FormattxtSupplierNPacking(False)
                'format dgv
                Me.FormatDGV()

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                If Me.cboPurchaseOrder.Text = DummyPO Then
                    Me.cboCodeNo.Focus()
                Else
                    Me.cboPurchaseOrder.Focus()
                End If

                If Me.TxtLocation.Visible = True Then
                    Me.txtPrimaryLocation.Text = PrimaryLocationLama
                    Me.TxtLocation.Text = Me.ObjLocation.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                    Me.TxtWarehouse.Text = Me.ObjLocation.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormattxtSupplierNPacking(ByVal Aktifkan As Boolean)
        Try
            If Aktifkan = False Then
                If Me.TxtPackingListNo.Text <> "" Then
                    Me.TxtPackingListNo.Enabled = False
                    Me.TxtPackingListNo.ReadOnly = True
                    Me.TxtPackingListNo.BackColor = Me.txtReadOnly.BackColor
                    Me.btnFindPackingList.Enabled = False
                End If
                If Me.TxtKodePihakAcuan.Text <> "" Then
                    Me.TxtPihakAcuan.Enabled = False
                    Me.TxtKodePihakAcuan.Enabled = False
                    Me.TxtKodePihakAcuan.ReadOnly = True
                    Me.TxtKodePihakAcuan.BackColor = Me.txtReadOnly.BackColor
                End If
            Else
                If Me.TxtPackingListNo.Text <> "" Then
                    Me.TxtPackingListNo.Enabled = True
                    Me.btnFindPackingList.Enabled = True
                    Me.TxtPackingListNo.ReadOnly = False
                    Me.TxtPackingListNo.BackColor = Me.TxtPrimaryStockIn.BackColor
                Else
                    Me.TxtPackingListNo.Enabled = True
                    Me.btnFindPackingList.Enabled = True
                    Me.TxtPackingListNo.ReadOnly = False
                    Me.TxtPackingListNo.BackColor = Me.TxtPrimaryStockIn.BackColor

                    Me.TxtPihakAcuan.Enabled = True
                    Me.TxtKodePihakAcuan.Enabled = True
                    Me.BtnCariPihakAcuan.Enabled = True
                    Me.TxtKodePihakAcuan.ReadOnly = False
                    Me.TxtKodePihakAcuan.BackColor = Me.TxtPrimaryStockIn.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub FormatDGV()
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

    ''@rosadinnn
    ''@MUaT
    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Expiry Date
            '5. Location Code
            '6. Remarks
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColSecondaryQty.Index Or e.ColumnIndex = Me.ColBatch.Index Or e.ColumnIndex = Me.ColLocation.Index _
                Or e.ColumnIndex = Me.ColRemark.Index Or e.ColumnIndex = Me.ColExpiredDate.Index Then
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

    ''@rosadinnn
    ''@MUaT
    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim frmRatio As New FrmTransaksiStockInPurchase2_SetQtyRatio
        Dim row As DataGridViewRow

        Dim JmlDGV As Decimal = 0
        Dim ExpiryDate As Date = #1/1/1950#
        Dim SecondaryQty As Decimal = -1
        Dim lanjut As Boolean = True
        Dim Use_POB_Keseluruhan As Boolean = False
        Dim PackingList_Ada_StockIn As Boolean = False

        Dim BufferQtyDefault As String = ""
        Dim BufferRatioReference As String = ""
        Dim BufferRatioDefault As String = ""
        Dim BufferRatioDesc As String = ""
        Dim BufferLblQty As String = ""

        Dim DefPrec As Integer = 0

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If

            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                If Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value Or _
                    Me.TxtRemark.Enabled = False Then
                    Exit Sub
                End If
                If Me.dgvItem.ReadOnly = True Then
                    Exit Sub
                End If

                BufferQtyDefault = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                BufferRatioReference = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                BufferRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                BufferRatioDesc = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                BufferLblQty = Me.LblQty.Text

                frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)
                frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                frmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value
                frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value

                If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                    Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value & _
                                                                            " " & Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value & " = " & _
                                                                            Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value & " " & _
                                                                            Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    If Me.txtPrimaryPO.Text = Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value.ToString And Me.TxtPrimaryCodeNo.Text = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value.ToString Then
                        Me.TxtRatioReference.Text = frmRatio.TxtRatioMulti.Text
                        Me.TxtRatioDefault.Text = frmRatio.TxtRatioDefault.Text
                        Me.LblQty.Text = "Qty (" & Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value & ")"
                    End If
                    Me.SetQtyDefaultOnDgv(e.RowIndex)

                    If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) = 0 Then
                        MsgBox("Ratio value restored to previous value, because ratio you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.dgvItem.Item(Me.ColRatioDesc.Index, e.RowIndex).Value = BufferRatioDesc
                        Me.LblQty.Text = BufferLblQty
                        Me.SetQtyDefaultOnDgv(e.RowIndex)
                    End If

                    'cek apakah jumlah digudang akan menjadi minus setelah perubahan data, hanya untuk goods.
                    If Me.ObjProduct.ApakahGoods(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) = True Then
                        JmlDGV = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                        If IsDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value) = True Then
                            ExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(Me.dgvItem.Item(Me.ColSecondaryQty.Index, e.RowIndex).Value) = True Then
                            SecondaryQty = Me.dgvItem.Item(Me.ColSecondaryQty.Index, e.RowIndex).Value
                        Else
                            SecondaryQty = -1
                        End If
                        If Me.ApakahEditSah(row, Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryPdO.Index, e.RowIndex).Value, _
                                    Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value, ExpiryDate, _
                                    SecondaryQty, JmlDGV, True) = False Then
                            lanjut = False
                        End If
                    End If

                    'cek apakah jumlah purchase order balance menjadi minus setelah perubahan data
                    Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri, JmlOB_Purchase As Decimal
                    JmlDGV = Me.GetJmlDGV_Multi(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value)
                    If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                        JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                        JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value)
                        JmlSendiri = Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value
                        'perhitungan menggunakan nilai jumlah dgv yang baru
                        JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama
                    Else
                        If Me.baru = True Then
                            JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                                     Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                            JmlPOBBaru = JmlPOB - JmlDGV
                        Else
                            If Me.ObjPackingList.GetNoStockInPurchase(Me.TxtPrimaryPackingList.Text) <> "" Then
                                'jika bukan baru maka selain order balance packing list maka juga harus memperhitungkan
                                'order balance keseluruhan
                                'hitung order balance keseluruhan
                                JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value, _
                                        Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                                JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                                        Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value)
                                JmlSendiri = Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value
                                'perhitungan menggunakan nilai jumlah dgv yang baru
                                JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama
                                JmlOB_Purchase = JmlPOBBaru

                                'hitung order balance packing list
                                JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                                        Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                                JmlPOBBaru = JmlPOB - JmlDGV

                                If JmlPOBBaru > JmlOB_Purchase Then
                                    JmlPOBBaru = JmlOB_Purchase
                                    Use_POB_Keseluruhan = True
                                Else
                                    Use_POB_Keseluruhan = False
                                End If
                                PackingList_Ada_StockIn = True
                            Else
                                JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                                        Me.dgvItem.Item(Me.ColPrimaryPOno.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                                JmlPOBBaru = JmlPOB - JmlDGV
                                PackingList_Ada_StockIn = False
                            End If
                        End If
                    End If

                    If JmlPOBBaru < 0 Then

                        DefPrec = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)

                        If Me.TxtPackingListNo.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                            MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB - (JmlDGV - JmlSendiri) + JmlObjLama), DefPrec) + _
                            " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                            MsgBoxStyle.Information, "Information")
                            lanjut = False
                        Else
                            If Me.baru = True Then
                                MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                                " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                MsgBoxStyle.Information, "Information")
                                lanjut = False
                            Else
                                If PackingList_Ada_StockIn = True Then
                                    If Use_POB_Keseluruhan = True Then
                                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlSendiri + JmlPOBBaru), DefPrec) + _
                                        " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                        MsgBoxStyle.Information, "Information")
                                        lanjut = False
                                    Else
                                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                                        " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                        MsgBoxStyle.Information, "Information")
                                        lanjut = False
                                    End If
                                Else
                                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                                    " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                    MsgBoxStyle.Information, "Information")
                                    lanjut = False
                                End If
                            End If
                        End If
                    End If

                    If lanjut = False Then
                        Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value = BufferQtyDefault
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioDesc
                        Me.LblQty.Text = BufferLblQty
                    End If

                    Me.SetCBO()
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
                            ExpiryDate = CDate(.Cells(Me.ColExpiredDate.Index).Value)
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
                        ElseIf e.ColumnIndex = Me.ColSecondaryQty.Index Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColBatch.Index Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColLocation.Index Then
                            Me.EndEditLocation(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColExpiredDate.Index Then
                            Me.EndEditExpiryDate(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColQty.Index Then
                            Me.EndEditQtyDefault(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
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

    Private Sub EndEditExpiryDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan format yang dituliskan adalah format tanggal
                If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'format ulang tanggal sehingga formatnya konsisten
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), _
                "dd MMM yyyy")
                'cek apakah back date terhadap tanggal transaksi, konfirmasikan data ini
                If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'pastikan tanggal expiry date tidak konflik dengan data yang sudah ada
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'bandingkan dengan yang sudah ada didalam dgv
                Dim row1 As DataGridViewRow
                For Each row1 In Me.dgvItem.Rows
                    If row1.Cells(Me.ColBatch.Index).Value = row.Cells(Me.ColBatch.Index).Value And CStr(row.Cells(Me.ColBatch.Index).Value) <> "" _
                    And row1.Index <> row.Index And row1.Cells(Me.ColPrimaryCodeNo.Index).Value = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                        If CDate(row1.Cells(Me.ColExpiredDate.Index).Value).Date <> CDate(row.Cells(Me.ColExpiredDate.Index).Value).Date Then
                            MsgBox("You already entered expiry date for batch no: " + CStr(row1.Cells(Me.ColBatch.Index).Value) + " on " + _
                            Format(CDate(row1.Cells(Me.ColExpiredDate.Index).Value), "dd MMM yyyy"), MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                'bandingkan dengan yang sudah ada didalam transaksi lain. Detail lihat point no 8.c pada 
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
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value, JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> CDate(Me.dgvItem(Me.ColExpiredDate.Index, e.RowIndex).Value) Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItem(Me.ColExpiredDate.Index, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")

                'expiry date, tidak perlu cek kemungkinan merger, karena tidak mungkin
                'hanya expiry date saja yang berbeda sedangkan batch no-nya sama
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatch.Index).Value, CDate(Me.buffer), SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim Use_POB_Keseluruhan As Boolean = False
        Dim PackingList_Ada_StockIn As Boolean = False

        Dim DefPrec As Integer = 0

        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data, hanya untuk goods.
                If Me.ObjProduct.ApakahGoods(.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then

                    Me.SetQtyDefaultOnDgv(e.RowIndex)

                    If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) = 0 Then
                        MsgBox("Qty restored to previous value, because qty you entered make qty default=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyDefaultOnDgv(e.RowIndex)
                        Exit Sub
                    End If

                    JmlDGV = .Cells(Me.ColQty.Index).Value

                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, _
                                    .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatch.Index).Value, ExpiryDate, _
                                    SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyDefaultOnDgv(e.RowIndex)
                        Exit Sub
                    End If

                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value))

                'cek apakah jumlah purchase order balance menjadi minus setelah perubahan data
                Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri, JmlOB_Purchase As Decimal

                JmlDGV = Me.GetJmlDGV_Multi(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPOno.Index).Value)
                If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                    JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPOno.Index).Value, _
                            .Cells(Me.ColPrimaryCodeNo.Index).Value)

                    ''@rosadinnn
                    ''@MUaT
                    Dim ToleranceUp As Decimal = objPO.GetToleranceUp(.Cells(ColPrimaryPOno.Index).Value, .Cells(ColPrimaryCodeNo.Index).Value)
                    JmlPOB += ToleranceUp

                    JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPOno.Index).Value)
                    JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value
                    'perhitungan menggunakan nilai jumlah dgv yang baru
                    JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama
                Else
                    If Me.baru = True Then
                        JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                         .Cells(Me.ColPrimaryPOno.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                        JmlPOBBaru = JmlPOB - JmlDGV
                    Else
                        If Me.ObjPackingList.GetNoStockInPurchase(Me.TxtPrimaryPackingList.Text) <> "" Then
                            'jika bukan baru maka selain order balance packing list maka juga harus memperhitungkan
                            'order balance keseluruhan
                            'hitung order balance keseluruhan
                            JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPOno.Index).Value, _
                                    .Cells(Me.ColPrimaryCodeNo.Index).Value)
                            JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                    .Cells(Me.ColPrimaryPOno.Index).Value)
                            JmlSendiri = .Cells(Me.ColQty.Index).Value
                            'perhitungan menggunakan nilai jumlah dgv yang baru
                            JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama
                            JmlOB_Purchase = JmlPOBBaru

                            'hitung order balance packing list
                            JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                            .Cells(Me.ColPrimaryPOno.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                            JmlPOBBaru = JmlPOB - JmlDGV

                            ''@rosadinnn
                            ''@MUaT
                            ''@hilangakn pengalihan ke OB Purchase
                            'If JmlPOBBaru > JmlOB_Purchase Then
                            '    JmlPOBBaru = JmlOB_Purchase
                            '    Use_POB_Keseluruhan = True
                            'Else
                            '    Use_POB_Keseluruhan = False
                            'End If
                            'PackingList_Ada_StockIn = True

                        Else
                            JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                             .Cells(Me.ColPrimaryPOno.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                            JmlPOBBaru = JmlPOB - JmlDGV
                            PackingList_Ada_StockIn = False
                        End If
                    End If
                End If

                

                If JmlPOBBaru < 0 Then

                    DefPrec = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)

                    If Me.TxtPackingListNo.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB - (JmlDGV - JmlSendiri) + JmlObjLama), DefPrec) + _
                        " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Else
                        If Me.baru = True Then
                            MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                            " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            If PackingList_Ada_StockIn = True Then
                                If Use_POB_Keseluruhan = True Then
                                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlSendiri + JmlPOBBaru), DefPrec) + _
                                    " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                    MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Else
                                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                                    " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                    MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            Else
                                MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                                " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        End If
                    End If
                End If

                Me.SetCBO()

                Me.SetQtyDefaultOnDgv(e.RowIndex)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub EndEditQtyDefault(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
            ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
            ByVal row As DataGridViewRow)

        Dim Use_POB_Keseluruhan As Boolean = False
        Dim PackingList_Ada_StockIn As Boolean = False

        Dim DefPrec As Integer = 0

        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data, hanya untuk goods.
                If Me.ObjProduct.ApakahGoods(.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then

                    Me.SetQtyMultiOnDgv(e.RowIndex)
                    JmlDGV = .Cells(Me.ColQty.Index).Value

                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, _
                                .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatch.Index).Value, ExpiryDate, _
                                SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyMultiOnDgv(e.RowIndex)

                        Exit Sub
                    End If

                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, _
                ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                'cek apakah jumlah purchase order balance menjadi minus setelah perubahan data
                Dim JmlObjLama, JmlPOBBaru, JmlPOB, JmlSendiri, JmlOB_Purchase As Decimal
                JmlDGV = Me.GetJmlDGV_Multi(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColPrimaryPOno.Index).Value)
                If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                    JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPOno.Index).Value, _
                            .Cells(Me.ColPrimaryCodeNo.Index).Value)
                    JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            .Cells(Me.ColPrimaryPOno.Index).Value)
                    JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value
                    'perhitungan menggunakan nilai jumlah dgv yang baru
                    JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama
                Else
                    If Me.baru = True Then
                        JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                         .Cells(Me.ColPrimaryPOno.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                        JmlPOBBaru = JmlPOB - JmlDGV
                    Else
                        If Me.ObjPackingList.GetNoStockInPurchase(Me.TxtPrimaryPackingList.Text) <> "" Then
                            'jika bukan baru maka selain order balance packing list maka juga harus memperhitungkan
                            'order balance keseluruhan
                            'hitung order balance keseluruhan
                            JmlPOB = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(.Cells(Me.ColPrimaryPOno.Index).Value, _
                                    .Cells(Me.ColPrimaryCodeNo.Index).Value)
                            JmlObjLama = Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                    .Cells(Me.ColPrimaryPOno.Index).Value)
                            JmlSendiri = .Cells(Me.ColQtyMulti.Index).Value
                            'perhitungan menggunakan nilai jumlah dgv yang baru
                            JmlPOBBaru = JmlPOB - JmlDGV + JmlObjLama
                            JmlOB_Purchase = JmlPOBBaru

                            'hitung order balance packing list
                            JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                            .Cells(Me.ColPrimaryPOno.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                            JmlPOBBaru = JmlPOB - JmlDGV

                            If JmlPOBBaru > JmlOB_Purchase Then
                                JmlPOBBaru = JmlOB_Purchase
                                Use_POB_Keseluruhan = True
                            Else
                                Use_POB_Keseluruhan = False
                            End If
                            PackingList_Ada_StockIn = True
                        Else
                            JmlPOB = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, _
                             .Cells(Me.ColPrimaryPOno.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                            JmlPOBBaru = JmlPOB - JmlDGV
                            PackingList_Ada_StockIn = False
                        End If
                    End If
                End If

                If JmlPOBBaru < 0 Then

                    DefPrec = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, e.RowIndex).Value)

                    If Me.TxtPackingListNo.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                        MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB - (JmlDGV - JmlSendiri) + JmlObjLama), DefPrec) + _
                        " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Else
                        If Me.baru = True Then
                            MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                            " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                            MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            If PackingList_Ada_StockIn = True Then
                                If Use_POB_Keseluruhan = True Then
                                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlSendiri + JmlPOBBaru), DefPrec) + _
                                    " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                    MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Else
                                    MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                                    " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                    MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            Else
                                MsgBox("Maximum value you can enter is " + ObjCD.DigitGrouping(CStr(JmlPOB), DefPrec) + _
                                " " + CStr(Me.dgvItem.Item(Me.ColUnitCodeMulti.Index, e.RowIndex).Value) + ".", _
                                MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        End If
                    End If
                End If

                Me.SetCBO()

                Me.SetQtyMultiOnDgv(e.RowIndex)
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
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
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
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                    End If
                End If
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Trim = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'batch no
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                Me.buffer, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value) = True Then
                        '0. basis data
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
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
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
                            ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                            Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value, JenisTransaksi)
                            If ExpiryDateLama <> #1/1/1950# Then
                                If ExpiryDateLama.Date <> CDate(Me.dgvItem(Me.ColExpiredDate.Index, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) + " is " + _
                                    Format(ExpiryDateLama, "dd MMM yyyy") + _
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
                        Next
                        '2. transaksi stock bundling untuk item hasil bundling
                        ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
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
                                    BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value)
                                    Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        '3. proses stock taking
                        'msgbox("belum")

                        '4. data grid view
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
                            Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = Format(CDate(BufferExpiryDate), "dd MMM yyyy")
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
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                Dim Obj As New clsMasterStorageArea
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
                Obj = Nothing
                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'update data primary location
                    Dim Obj1 As New clsMasterStorageArea
                    .Cells(Me.ColPrimaryLocation.Index).Value = Obj1.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                    'update data gudang
                    .Cells(Me.ColWareHouse.Index).Value = Obj1.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)

                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'kembalikan data gudang
                            .Cells(Me.ColPrimaryLocation.Index).Value = Obj1.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                            'update data gudang
                            .Cells(Me.ColWareHouse.Index).Value = Obj1.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                            Obj1 = Nothing
                            Exit Sub
                        End If
                        Me.GabungkanBaris()
                    End If

                    Obj1 = Nothing
                End If
            End With
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

                    If row.Cells(Me.ColPONo.Index).Value = row1.Cells(Me.ColPONo.Index).Value And row.Cells(Me.ColSuppDONo.Index).Value = row1.Cells(Me.ColSuppDONo.Index).Value And _
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
                    If CStr(row.Cells(Me.ColPONo.Index).Value) = CStr(row1.Cells(Me.ColPONo.Index).Value) And CStr(row.Cells(Me.ColSuppDONo.Index).Value) = CStr(row1.Cells(Me.ColSuppDONo.Index).Value) And _
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
            row.Cells(Me.ColQty.Index).Value = CDec(row.Cells(Me.ColQty.Index).Value) + CDec(row1.Cells(Me.ColQty.Index).Value)
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

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Decimal, _
    ByVal PrimaryPdO As Decimal, ByVal PrimaryProduct As Decimal, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal jmlDGV As Decimal, Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal, JmlDGVLain As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.ObjStockInOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV_Default(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGVTotal
            JmlDGVLain = JmlDGVTotal - jmlDGV

            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make your stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Minimum value you must can enter is " + Me.ObjCD.DigitGrouping(jmlDGV - JmlSetelahHapus, Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + _
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

    Private Sub dgvItem_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEnter

    End Sub

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

    Private Sub TxtLocation_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocation.KeyDown
        Dim i As New FrmCariMasterStorageArea
        Dim Obj As New clsMasterStorageArea

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    Me.txtPrimaryLocation.Text = Obj.GetPrimary(Me.TxtLocation.Text)
                    Me.TxtWarehouse.Text = Obj.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                    Me.TxtQtyperBatch.Focus()
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtWarehouse.Text = Obj.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                        Me.TxtLocation.Text = Obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        Me.TxtQtyperBatch.Focus()
                    Else
                        Me.TxtLocation.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        Me.TxtWarehouse.Clear()
                        Me.TxtLocation.Focus()
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

    Private Sub TxtLocation_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.Leave
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            If Me.TxtLocation.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtLocation.Text) = True Then
                    Me.TxtQtyperBatch.Focus()
                    Me.txtPrimaryLocation.Text = obj.GetPrimary(Me.TxtLocation.Text)
                    Me.TxtWarehouse.Text = obj.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        Me.TxtWarehouse.Text = obj.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                        Me.TxtQtyperBatch.Focus()
                    Else
                        Me.TxtLocation.Clear()
                        Me.TxtWarehouse.Clear()
                        Me.txtPrimaryLocation.Text = "-1"
                        Me.TxtLocation.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryLocation.Text = "-1"
                Me.TxtWarehouse.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindLocation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocation.Click
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocation.Text <> "" Then
                i.txtCari.Text = Me.TxtLocation.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                Me.TxtLocation.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                Me.TxtWarehouse.Text = obj.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                Me.TxtQtyperBatch.Focus()
            Else
                Me.TxtWarehouse.Clear()
                Me.TxtLocation.Clear()
                Me.TxtLocation.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
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
            Me.SetCBO()
            Me.ObjCD.SetNumber(Me.dgvItem)
            If Me.dgvItem.Rows.Count = 0 Then
                Me.FormattxtSupplierNPacking(True)
            End If
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
        Dim PrimaryLocation, PrimaryPdO, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
        JmlDGV, JmlSetelahHapus As Decimal
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim row As DataGridViewRow = e.Row

        Try
            If Me.ObjProduct.ApakahGoods(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                With row
                    PrimaryLocation = .Cells(Me.ColPrimaryLocation.Index).Value
                    PrimaryPdO = .Cells(Me.ColPrimaryPdO.Index).Value
                    PrimaryProduct = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    BatchNo = .Cells(Me.ColBatch.Index).Value
                    If IsDate(.Cells(Me.ColExpiredDate.Index).Value) = True Then
                        ExpiryDate = CDate(.Cells(Me.ColExpiredDate.Index).Value)
                    Else
                        ExpiryDate = #1/1/1950#
                    End If
                    If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                        SecondaryQty = .Cells(Me.ColSecondaryQty.Index).Value
                    Else
                        SecondaryQty = -1
                    End If
                    JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                    'karena akan dihapus, perhitungan menggunakan nilai jumlah dgv =0
                    JmlDGV = Me.GetJmlDGV_Default(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, e.Row.Index)

                    JmlObjLama = Me.ObjStockInOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
                    SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                    JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
                    If JmlSetelahHapus < 0 Then
                        MsgBox("You cannot delete this data, because this will make your stock quantity minus.", _
                        MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                    End If
                End With
            End If
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

    Private Sub txtBatchNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.GotFocus
        Try
            Me.txtBatchNo.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.LostFocus
        Try
            Me.txtBatchNo.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.GotFocus
        Try
            Me.TxtLocation.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocation_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocation.LostFocus
        Try
            Me.TxtLocation.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then

                    ''@rosadinnn
                    ''@MUaT
                    Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))
                    Me.SetQtyDefault()
                    'Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.GotFocus
        Try
            Me.TxtQtyperBatch.BackColor = Core_Module.GetSMESSBackColorAll
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

    Private Sub TxtKodePihakAcuan_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtKodePihakAcuan.KeyDown
        If Me.TxtKodePihakAcuan.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjSupplier.GetJml(Me.TxtKodePihakAcuan.Text, Core_Module.EnumSearchCriteria.ByCode)
                If Jml = 1 Then
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtSuppDONumber.Focus()
                    Me.SetCBO()
                    Me.TxtPackingListNo.ReadOnly = True
                    Me.TxtPackingListNo.BackColor = Me.TxtStockInNo.BackColor
                    Me.btnFindPackingList.Enabled = False
                    Me.TxtPackingListNo.TabStop = False
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
                        Me.TxtSuppDONumber.Focus()
                        Me.SetCBO()
                        Me.TxtPackingListNo.ReadOnly = True
                        Me.TxtPackingListNo.BackColor = Me.TxtStockInNo.BackColor
                        Me.btnFindPackingList.Enabled = False
                        Me.TxtPackingListNo.TabStop = False
                    Else
                        Me.TxtKodePihakAcuan.Clear()
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtKodePihakAcuan.Focus()
                        Me.TxtPackingListNo.ReadOnly = False
                        Me.TxtPackingListNo.BackColor = Me.TxtPrimaryStockIn.BackColor
                        Me.btnFindPackingList.Enabled = True
                        Me.TxtPackingListNo.TabStop = True
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
        If Me.TxtKodePihakAcuan.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier

        Try
            If Me.TxtKodePihakAcuan.Text <> "" Then
                If Me.ObjSupplier.GetJml(Me.TxtKodePihakAcuan.Text) = 1 Then
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtSuppDONumber.Focus()
                    Me.SetGboItem()
                    Me.SetCBO()
                    Me.TxtPackingListNo.ReadOnly = True
                    Me.TxtPackingListNo.BackColor = Me.TxtStockInNo.BackColor
                    Me.btnFindPackingList.Enabled = False
                    Me.TxtPackingListNo.TabStop = False
                Else
                    If Me.TxtKodePihakAcuan.Text <> "" Then
                        i.txtCari.Text = Me.TxtKodePihakAcuan.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.TxtPrimaryPihakAcuan.Text = i.PrimarySupplier
                        Me.TxtKodePihakAcuan.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtPihakAcuan.Text = Me.ObjSupplier.GetSupplierName(Me.TxtPrimaryPihakAcuan.Text)
                        Me.TxtSuppDONumber.Focus()
                        Me.SetCBO()
                        Me.TxtPackingListNo.ReadOnly = True
                        Me.TxtPackingListNo.BackColor = Me.TxtStockInNo.BackColor
                        Me.btnFindPackingList.Enabled = False
                        Me.TxtPackingListNo.TabStop = False
                    Else
                        Me.TxtKodePihakAcuan.Clear()
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtKodePihakAcuan.Focus()
                        Me.TxtPackingListNo.ReadOnly = False
                        Me.TxtPackingListNo.BackColor = Me.TxtPrimaryStockIn.BackColor
                        Me.btnFindPackingList.Enabled = True
                        Me.TxtPackingListNo.TabStop = True
                    End If
                End If
            Else
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtPihakAcuan.Clear()
                Me.KosongkanItem()
                Me.TxtPackingListNo.ReadOnly = False
                Me.TxtPackingListNo.BackColor = Me.TxtPrimaryStockIn.BackColor
                Me.btnFindPackingList.Enabled = True
                Me.TxtPackingListNo.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnCariPihakAcuan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCariPihakAcuan.Click
        If Me.TxtKodePihakAcuan.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

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
                Me.TxtSuppDONumber.Focus()
                Me.TxtPackingListNo.ReadOnly = True
                Me.TxtPackingListNo.BackColor = Me.TxtStockInNo.BackColor
                Me.btnFindPackingList.Enabled = False
                Me.TxtPackingListNo.TabStop = False
                Me.SetCBO()
            Else
                Me.TxtKodePihakAcuan.Clear()
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.KosongkanItem()
                Me.TxtKodePihakAcuan.Focus()
                Me.TxtPackingListNo.ReadOnly = False
                Me.TxtPackingListNo.BackColor = Me.TxtPrimaryStockIn.BackColor
                Me.btnFindPackingList.Enabled = True
                Me.TxtPackingListNo.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtKodePihakAcuan_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodePihakAcuan.GotFocus
        Try
            If Me.TxtKodePihakAcuan.Enabled = True And Me.TxtKodePihakAcuan.ReadOnly = False And Me.TxtRemark.Enabled = True Then
                Me.TxtKodePihakAcuan.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodePihakAcuan_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodePihakAcuan.LostFocus
        Try
            If Me.TxtKodePihakAcuan.Enabled = True And Me.TxtKodePihakAcuan.ReadOnly = False Then
                Me.TxtKodePihakAcuan.BackColor = Me.txtPrimaryCategory.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodePihakAcuan_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtKodePihakAcuan.TextChanged
        Try
            If Me.TxtKodePihakAcuan.Text = "" Then
                Me.TxtPihakAcuan.Text = ""
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.cboPurchaseOrder.DataSource = Nothing
                Me.cboCodeNo.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindPackingList_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnFindPackingList.Click
        If Me.TxtPackingListNo.ReadOnly = True Then
            Exit Sub
        End If
        Dim StockInNo As String = ""
        Dim i As New FrmCariTransaksiPackingList
        Dim obj As New ClsTransaksiPackingList
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtPackingListNo.Text <> "" Then
                i.txtCari.Text = Me.TxtPackingListNo.Text
            End If
            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                StockInNo = Me.ObjPackingList.GetNoStockInPurchase(i.PrimaryTransaksi)
                If StockInNo <> "" Then
                    MsgBox("The packing list list is alredy refered by stock in transaction: " + _
                    StockInNo, MsgBoxStyle.Information, "Information")
                    Me.TxtPrimaryPackingList.Text = "-1"
                    Me.TxtPackingListNo.Text = ""
                    Me.SetCBO()
                    Me.TxtPackingListNo.Focus()
                    Me.TxtOrderBalance.Text = "0"
                    Me.dgvPackingList.Rows.Clear()
                    Me.KosongkanItem()
                    Exit Sub
                End If

                If Me.TxtPihakAcuan.Text <> "" Then
                    Dim SupName, SupCode, NoTrans As String
                    SupName = ""
                    SupCode = ""
                    NoTrans = ""
                    obj.GetInfoSupplierNameNCodeNTransNo(i.PrimaryTransaksi, SupName, SupCode, NoTrans)

                    If SupCode <> Me.TxtKodePihakAcuan.Text Then
                        MsgBox("Supplier for this Packing List is " + SupName + ".", MsgBoxStyle.Information, "Information")
                        If IsNumeric(Me.TxtPrimaryPackingList.Text) = False Then
                            Me.TxtPrimaryPackingList.Text = "-1"
                        End If
                        Me.TxtPackingListNo.Text = obj.GetNoTransaksi(Me.TxtPrimaryPackingList.Text)
                        Me.TxtPackingListNo.Focus()
                    Else
                        Me.TxtPrimaryPackingList.Text = i.PrimaryTransaksi
                        Me.TxtPackingListNo.Text = NoTrans
                        Me.TxtKodePihakAcuan.Text = SupCode
                        Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                        Me.TxtPihakAcuan.Text = SupName
                        Me.TxtKodePihakAcuan.Enabled = False
                        Me.BtnCariPihakAcuan.Enabled = False
                        Me.objPackingListItem.TransferPackingList_Multi(Me, Me.TxtPrimaryPackingList.Text)
                        Me.SetCBO()
                        Me.IsikanDGVPackingLIst(Me.dgvPackingList, Me.TxtPrimaryPackingList.Text)
                    End If
                Else
                    Me.TxtPrimaryPackingList.Text = i.PrimaryTransaksi
                    obj.GetInfoSupplierNameNCodeNTransNo(Me.TxtPrimaryPackingList.Text, Me.TxtPihakAcuan.Text, Me.TxtKodePihakAcuan.Text, Me.TxtPackingListNo.Text)
                    Me.IsikanDGVPackingLIst(Me.dgvPackingList, Me.TxtPrimaryPackingList.Text)
                    Me.TxtKodePihakAcuan.Enabled = False
                    Me.BtnCariPihakAcuan.Enabled = False
                    Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                    Me.objPackingListItem.TransferPackingList_Multi(Me, Me.TxtPrimaryPackingList.Text)
                    Me.SetCBO()
                    Me.TxtRemark.Focus()
                End If
            Else
                Me.TxtPrimaryPackingList.Text = "-1"
                Me.TxtPackingListNo.Text = ""
                Me.SetCBO()
                Me.TxtPackingListNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtPackingListNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPackingListNo.GotFocus
        Try
            If Me.TxtPackingListNo.Enabled = True Then
                If Me.TxtPackingListNo.ReadOnly = False Then
                    Me.TxtPackingListNo.BackColor = Me.BackColorAll
                Else
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPackingListNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPackingListNo.KeyDown
        If Me.TxtPackingListNo.ReadOnly = True Then
            Exit Sub
        End If

        Dim StockInNo As String = ""
        Dim obj As New ClsTransaksiPackingList

        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Decimal = obj.GetPrimary(Me.TxtPackingListNo.Text)
                If primary > -1 Then
                    StockInNo = Me.ObjPackingList.GetNoStockInPurchase(primary)
                    If StockInNo <> "" Then
                        MsgBox("The packing list list is alredy refered by stock in transaction: " + _
                        StockInNo, MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryPackingList.Text = "-1"
                        Me.TxtPackingListNo.Text = ""
                        Me.SetCBO()
                        Me.TxtPackingListNo.Focus()
                        Me.TxtOrderBalance.Text = "0"
                        Me.dgvPackingList.Rows.Clear()
                        Exit Sub
                    End If

                    If Me.TxtPihakAcuan.Text <> "" Then
                        Dim SupName, SupCode, NoTrans As String
                        SupName = ""
                        SupCode = ""
                        NoTrans = ""

                        obj.GetInfoSupplierNameNCodeNTransNo(primary, SupName, SupCode, NoTrans)
                        If SupCode <> Me.TxtKodePihakAcuan.Text Then
                            MsgBox("Supplier for this Packing List is " + SupName + ".", MsgBoxStyle.Information, "Information")
                            If IsNumeric(Me.TxtPrimaryPackingList.Text) = False Then
                                Me.TxtPrimaryPackingList.Text = "-1"
                            End If
                            Me.TxtPackingListNo.Text = obj.GetNoTransaksi(Me.TxtPrimaryPackingList.Text)
                            Me.TxtPackingListNo.Focus()
                        Else
                            Me.TxtPrimaryPackingList.Text = primary
                            Me.TxtPackingListNo.Text = NoTrans
                            Me.TxtKodePihakAcuan.Text = SupCode
                            Me.TxtPihakAcuan.Text = SupName
                            Me.TxtKodePihakAcuan.Enabled = False
                            Me.BtnCariPihakAcuan.Enabled = False
                            Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                            Me.objPackingListItem.TransferPackingList_Multi(Me, Me.TxtPrimaryPackingList.Text)
                            Me.SetCBO()
                            Me.IsikanDGVPackingLIst(Me.dgvPackingList, Me.TxtPrimaryPackingList.Text)
                        End If
                    Else
                        Me.TxtPrimaryPackingList.Text = primary
                        obj.GetInfoSupplierNameNCodeNTransNo(Me.TxtPrimaryPackingList.Text, Me.TxtPihakAcuan.Text, Me.TxtKodePihakAcuan.Text, Me.TxtPackingListNo.Text)
                        Me.IsikanDGVPackingLIst(Me.dgvPackingList, Me.TxtPrimaryPackingList.Text)
                        Me.TxtKodePihakAcuan.Enabled = False
                        Me.BtnCariPihakAcuan.Enabled = False
                        Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                        Me.SetCBO()
                        Me.TxtRemark.Focus()
                    End If
                Else
                    Me.btnFindPackingList.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtPackingListNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPackingListNo.Leave
        If Me.TxtPackingListNo.ReadOnly = True Then
            Exit Sub
        End If

        Dim obj As New ClsTransaksiPackingList
        Dim StockInNo As String = ""

        Try
            If Me.TxtPackingListNo.Text <> "" Then
                Dim primary As Decimal = obj.GetPrimary(Me.TxtPackingListNo.Text)
                If primary > -1 Then
                    StockInNo = Me.ObjPackingList.GetNoStockInPurchase(primary)
                    If StockInNo <> "" Then
                        MsgBox("The packing list list is alredy refered by stock in transaction: " + _
                        StockInNo, MsgBoxStyle.Information, "Information")
                        Me.TxtPrimaryPackingList.Text = "-1"
                        Me.TxtPackingListNo.Text = ""
                        Me.SetCBO()
                        Me.TxtPackingListNo.Focus()
                        Me.TxtOrderBalance.Text = "0"
                        Me.dgvPackingList.Rows.Clear()
                        Exit Sub
                    End If

                    If Me.TxtPihakAcuan.Text <> "" Then
                        Dim SupName, SupCode, NoTrans As String
                        SupName = ""
                        SupCode = ""
                        NoTrans = ""

                        obj.GetInfoSupplierNameNCodeNTransNo(primary, SupName, SupCode, NoTrans)
                        If SupCode <> Me.TxtKodePihakAcuan.Text Then
                            MsgBox("Supplier for this Packing List is " + SupName + ".", MsgBoxStyle.Information, "Information")
                            If IsNumeric(Me.TxtPrimaryPackingList.Text) = False Then
                                Me.TxtPrimaryPackingList.Text = "-1"
                            End If
                            Me.TxtPackingListNo.Text = obj.GetNoTransaksi(Me.TxtPrimaryPackingList.Text)
                            Me.TxtPackingListNo.Focus()
                        Else
                            Me.TxtPrimaryPackingList.Text = primary
                            Me.TxtPackingListNo.Text = NoTrans
                            Me.TxtKodePihakAcuan.Text = SupCode
                            Me.TxtPihakAcuan.Text = SupName
                            Me.TxtKodePihakAcuan.Enabled = False
                            Me.BtnCariPihakAcuan.Enabled = False
                            Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                            Me.objPackingListItem.TransferPackingList_Multi(Me, Me.TxtPrimaryPackingList.Text)
                            Me.SetCBO()
                            Me.IsikanDGVPackingLIst(Me.dgvPackingList, Me.TxtPrimaryPackingList.Text)
                        End If
                    Else
                        Me.TxtPrimaryPackingList.Text = primary
                        obj.GetInfoSupplierNameNCodeNTransNo(Me.TxtPrimaryPackingList.Text, Me.TxtPihakAcuan.Text, Me.TxtKodePihakAcuan.Text, Me.TxtPackingListNo.Text)
                        Me.IsikanDGVPackingLIst(Me.dgvPackingList, Me.TxtPrimaryPackingList.Text)
                        Me.TxtKodePihakAcuan.Enabled = False
                        Me.BtnCariPihakAcuan.Enabled = False
                        Me.TxtPrimaryPihakAcuan.Text = ObjSupplier.GetPrimary(Me.TxtKodePihakAcuan.Text)
                        Me.SetCBO()
                        Me.TxtRemark.Focus()
                    End If
                Else
                    Me.btnFindPackingList.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub TxtPackingListNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPackingListNo.LostFocus
        Try
            If Me.TxtPackingListNo.Enabled = True Then
                If Me.TxtPackingListNo.ReadOnly = False Then
                    Me.TxtPackingListNo.BackColor = Me.TxtPrimaryPackingList.BackColor
                Else
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPackingListNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPackingListNo.TextChanged
        Try
            If Me.TxtPackingListNo.Text = "" Then
                Me.TxtPrimaryPackingList.Text = "-1"
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtKodePihakAcuan.Text = ""
                Me.TxtPihakAcuan.Text = ""
                Me.TxtKodePihakAcuan.Enabled = True
                Me.TxtKodePihakAcuan.ReadOnly = False
                Me.TxtKodePihakAcuan.BackColor = Me.TxtPrimaryStockIn.BackColor
                Me.BtnCariPihakAcuan.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDGVPackingLIst(ByVal vDGV As DataGridView, ByVal vPrimaryGIT As Decimal)
        Dim objItem As New ClsTransaksiPackingListItem
        Try
            objItem.IsiDGV2_Multi(vDGV, vPrimaryGIT)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objItem = Nothing
        End Try
    End Sub

    Private Function ApakahDGVItemNGITCocok() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowItem, rowGIT, rowItem2, rowGIT2 As DataGridViewRow
        Dim Cocok, Ketemu As Boolean
        Try
            Me.dgvPackingList2.Rows.Clear()
            Me.DGVItem2.Rows.Clear()
            Me.DGVItemLebihBesarDariGIT.Rows.Clear()
            Me.DGVItemLebihKecilDariGIT.Rows.Clear()
            For Each rowItem In Me.dgvItem.Rows
                Cocok = False
                For Each rowGIT In Me.dgvPackingList.Rows
                    If rowItem.Cells(Me.ColPONo.Index).Value = rowGIT.Cells(Me.colPONoDGVPackingList.Index).Value And _
                    rowItem.Cells(Me.ColCodeNo.Index).Value = rowGIT.Cells(Me.colCodeNoDGVPackingList.Index).Value Then
                        Cocok = True
                        Exit For
                    End If
                Next
                If Cocok = False Then
                    If Me.TxtPackingListNo.Text <> "" Then
                        MsgBox("Can not find Product: " + rowItem.Cells(Me.ColCodeNo.Index).Value + " in Packing List Item.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next
            For Each rowItem In Me.dgvItem.Rows
                If Me.DGVItem2.Rows.Count = 0 Then
                    Me.DGVItem2.Rows.Add("", rowItem.Cells(Me.ColPONo.Index).Value, "", rowItem.Cells(Me.ColCodeNo.Index).Value, _
                    rowItem.Cells(Me.ColDescription.Index).Value, rowItem.Cells(Me.ColQty.Index).Value, _
                    rowItem.Cells(Me.ColUnit.Index).Value, rowItem.Cells(Me.ColSecondaryQty.Index).Value, _
                    rowItem.Cells(Me.ColSecondaryUnit.Index).Value, rowItem.Cells(Me.ColBatch.Index).Value, _
                    rowItem.Cells(Me.ColExpiredDate.Index).Value, rowItem.Cells(Me.ColLocation.Index).Value, _
                    rowItem.Cells(Me.ColWareHouse.Index).Value, rowItem.Cells(Me.ColRemark.Index).Value, _
                    rowItem.Cells(Me.ColPrimaryPOno.Index).Value, rowItem.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    rowItem.Cells(Me.ColPrimaryLocation.Index).Value, rowItem.Cells(Me.ColPrimaryUnits.Index).Value, _
                    rowItem.Cells(Me.ColPrimarySecondaryUnits.Index).Value, rowItem.Cells(Me.ColPrimaryPdO.Index).Value)
                Else
                    Ketemu = False
                    For Each rowItem2 In Me.DGVItem2.Rows
                        If rowItem2.Cells(Me.ColPONo.Index).Value = rowItem.Cells(Me.ColPONo.Index).Value And _
                           rowItem2.Cells(Me.ColCodeNo.Index).Value = rowItem.Cells(Me.ColCodeNo.Index).Value Then
                            rowItem2.Cells(Me.ColQty.Index).Value = CDec(rowItem2.Cells(Me.ColQty.Index).Value) + CDec(rowItem.Cells(Me.ColQty.Index).Value)
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        Me.DGVItem2.Rows.Add("", rowItem.Cells(Me.ColPONo.Index).Value, "", rowItem.Cells(Me.ColCodeNo.Index).Value, _
                        rowItem.Cells(Me.ColDescription.Index).Value, rowItem.Cells(Me.ColQty.Index).Value, _
                        rowItem.Cells(Me.ColUnit.Index).Value, rowItem.Cells(Me.ColSecondaryQty.Index).Value, _
                        rowItem.Cells(Me.ColSecondaryUnit.Index).Value, rowItem.Cells(Me.ColBatch.Index).Value, _
                        rowItem.Cells(Me.ColExpiredDate.Index).Value, rowItem.Cells(Me.ColLocation.Index).Value, _
                        rowItem.Cells(Me.ColWareHouse.Index).Value, rowItem.Cells(Me.ColRemark.Index).Value, _
                        rowItem.Cells(Me.ColPrimaryPOno.Index).Value, rowItem.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        rowItem.Cells(Me.ColPrimaryLocation.Index).Value, rowItem.Cells(Me.ColPrimaryUnits.Index).Value, _
                        rowItem.Cells(Me.ColPrimarySecondaryUnits.Index).Value, rowItem.Cells(Me.ColPrimaryPdO.Index).Value)
                    End If
                End If
            Next
            For Each rowGIT In Me.dgvPackingList.Rows
                If Me.dgvPackingList2.Rows.Count = 0 Then
                    Me.dgvPackingList2.Rows.Add("0", rowGIT.Cells(Me.colPONoDGVPackingList.Index).Value, rowGIT.Cells(Me.colCodeNoDGVPackingList.Index).Value, _
                    rowGIT.Cells(Me.colDescDGVPackingLIst.Index).Value, rowGIT.Cells(Me.ColSuppCodeNoDGVPackingLIst.Index).Value, _
                    rowGIT.Cells(Me.ColSuppDescDGVPackingList.Index).Value, _
                    rowGIT.Cells(Me.ColQtyMasterDGVPackingList.Index).Value, rowGIT.Cells(Me.ColQtyPerMasterDGVPackingList.Index).Value, _
                    rowGIT.Cells(Me.ColMasterUnitDGVPackingList.Index).Value, rowGIT.Cells(Me.colQtyDGVPackingList.Index).Value, _
                    rowGIT.Cells(Me.ColTotalQtyDGVPackingList.Index).Value, rowGIT.Cells(Me.colUnitsDGVPackingList.Index).Value, _
                    rowGIT.Cells(Me.colRemarksDGVPackingList.Index).Value, _
                    rowGIT.Cells(Me.colPrimaryCodeNoDGVPAckingList.Index).Value, rowGIT.Cells(Me.colPrimaryUnitsDGVPackingLIst.Index).Value, _
                    rowGIT.Cells(Me.ColPrimaryMasterUnitsDGVPAckingLIst.Index).Value, rowGIT.Cells(Me.ColPrimaryPODGVPackingLIst.Index).Value)
                Else
                    Ketemu = False
                    For Each rowGIT2 In Me.dgvPackingList2.Rows
                        If rowGIT2.Cells(Me.colPONoDGVPackingList.Index).Value = rowGIT.Cells(Me.colPONoDGVPackingList.Index).Value And _
                           rowGIT2.Cells(Me.colCodeNoDGVPackingList.Index).Value = rowGIT.Cells(Me.colCodeNoDGVPackingList.Index).Value Then
                            rowGIT2.Cells(Me.ColTotalQtyDGVPackingList.Index).Value = CDec(rowGIT2.Cells(Me.ColTotalQtyDGVPackingList.Index).Value) + CDec(rowGIT.Cells(Me.ColTotalQtyDGVPackingList.Index).Value)
                            Ketemu = True
                            Exit For
                        End If
                    Next
                    If Ketemu = False Then
                        Me.dgvPackingList2.Rows.Add("0", rowGIT.Cells(Me.colPONoDGVPackingList.Index).Value, rowGIT.Cells(Me.colCodeNoDGVPackingList.Index).Value, _
                        rowGIT.Cells(Me.colDescDGVPackingLIst.Index).Value, rowGIT.Cells(Me.ColSuppCodeNoDGVPackingLIst.Index).Value, _
                        rowGIT.Cells(Me.ColSuppDescDGVPackingList.Index).Value, _
                        rowGIT.Cells(Me.ColQtyMasterDGVPackingList.Index).Value, rowGIT.Cells(Me.ColQtyPerMasterDGVPackingList.Index).Value, _
                        rowGIT.Cells(Me.ColMasterUnitDGVPackingList.Index).Value, rowGIT.Cells(Me.colQtyDGVPackingList.Index).Value, _
                        rowGIT.Cells(Me.ColTotalQtyDGVPackingList.Index).Value, rowGIT.Cells(Me.colUnitsDGVPackingList.Index).Value, _
                        rowGIT.Cells(Me.colRemarksDGVPackingList.Index).Value, _
                        rowGIT.Cells(Me.colPrimaryCodeNoDGVPAckingList.Index).Value, rowGIT.Cells(Me.colPrimaryUnitsDGVPackingLIst.Index).Value, _
                        rowGIT.Cells(Me.ColPrimaryMasterUnitsDGVPAckingLIst.Index).Value, rowGIT.Cells(Me.ColPrimaryPODGVPackingLIst.Index).Value)
                    End If
                End If
            Next
            For Each rowItem2 In Me.DGVItem2.Rows
                For Each rowGIT2 In Me.dgvPackingList2.Rows
                    If rowItem2.Cells(Me.ColPONo.Index).Value = rowGIT2.Cells(Me.colPONoDGVPackingList.Index).Value And _
                       rowItem2.Cells(Me.ColQty.Index).Value < rowGIT2.Cells(Me.ColTotalQtyDGVPackingList.Index).Value And _
                       rowItem2.Cells(Me.ColCodeNo.Index).Value = rowGIT2.Cells(Me.colCodeNoDGVPackingList.Index).Value Then
                        Me.DGVItemLebihKecilDariGIT.Rows.Add(rowItem2.Cells(Me.ColPONo.Index).Value, rowItem2.Cells(Me.ColCodeNo.Index).Value, _
                        rowItem2.Cells(Me.ColBatch.Index).Value, rowItem2.Cells(Me.ColQty.Index).Value)
                    ElseIf rowItem2.Cells(Me.ColPONo.Index).Value = rowGIT2.Cells(Me.colPONoDGVPackingList.Index).Value And _
                       rowItem2.Cells(Me.ColQty.Index).Value > rowGIT2.Cells(Me.ColTotalQtyDGVPackingList.Index).Value And _
                       rowItem2.Cells(Me.ColCodeNo.Index).Value = rowGIT2.Cells(Me.colCodeNoDGVPackingList.Index).Value Then
                        Me.DGVItemLebihBesarDariGIT.Rows.Add(rowItem2.Cells(Me.ColPONo.Index).Value, rowItem2.Cells(Me.ColCodeNo.Index).Value, _
                        rowItem2.Cells(Me.ColBatch.Index).Value, rowItem2.Cells(Me.ColQty.Index).Value)
                    End If
                Next
            Next
            If Me.DGVItemLebihBesarDariGIT.Rows.Count > 0 Then
                Dim i As String = ""
                For Each rowItem In Me.DGVItemLebihBesarDariGIT.Rows
                    If i = "" Then
                        i = "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is bigger than GIT qty"
                    Else
                        i = i + ", " + "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is bigger than GIT qty"
                    End If
                Next
                i = i + "."
                MsgBox("Can not save because " + i, MsgBoxStyle.OkOnly, "Confirm")
                Return False
            End If
            If Me.DGVItemLebihKecilDariGIT.Rows.Count > 0 Then
                Dim i As String = ""
                For Each rowItem In Me.DGVItemLebihKecilDariGIT.Rows
                    If i = "" Then
                        i = "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is more less than GIT qty"
                    Else
                        i = i + ", " + "PO Number: " + rowItem.Cells(Me.colPONoDGVPerbandingan.Index).Value + "(Product: " + rowItem.Cells(Me.colCodeNoDGVPerbandingan.Index).Value + ") the qty is more less than GIT qty"
                    End If
                Next
                i = i + "."
                If MsgBox(i + " If you want to save click 'Yes' else you want to change or edit the qty click 'No'.", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub cboPurchaseOrder_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPurchaseOrder.SelectedValueChanged
    '    Try
    '        If Me.cboPurchaseOrder.SelectedIndex <> -1 Then
    '            Dim RowView As Data.DataRowView = Me.cboPurchaseOrder.Items(Me.cboPurchaseOrder.SelectedIndex)
    '            Dim row2 As DseTransaksiPO.orderbalance_report_view0Row = RowView.Row
    '            Dim PrimaryPOBaru As Decimal = row2.Primary_PO
    '            If Me.txtPrimaryPO.Text <> CStr(PrimaryPOBaru) Then
    '                Me.txtPrimaryPO.Text = PrimaryPOBaru
    '                Me.IsiCboProduct(Me.txtPrimaryPO.Text)
    '            End If
    '        Else
    '            Me.txtPrimaryPO.Text = "-1"
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub cboPurchaseOrder_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPurchaseOrder.SelectedValueChanged
        Try
            If Me.cboPurchaseOrder.SelectedIndex <> -1 Then
                Dim RowView As Data.DataRowView = Me.cboPurchaseOrder.Items(Me.cboPurchaseOrder.SelectedIndex)
                Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow = RowView.Row

                If Me.CekCbo = True Then
                    Dim PrimaryPOBaru As Decimal = row2.Primary_PO
                    If Me.txtPrimaryPO.Text <> CStr(PrimaryPOBaru) Then
                        Me.txtPrimaryPO.Text = PrimaryPOBaru
                        Me.IsiCboProduct(Me.txtPrimaryPO.Text)
                    End If
                End If
            Else
                Me.txtPrimaryPO.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTBLProduct(ByVal vPrimaryPO As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiRow
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Primary_Supplier = Me.TxtPrimaryPihakAcuan.Text And _
                        row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        row1.Primary_PO_Purchase = row.Cells(Me.ColPrimaryPOno.Index).Value Then
                            'kalo ketemu ubah data order balance
                            row1.TotalQty = row1.TotalQty + CDec(row.Cells(Me.ColQty.Index).Value)

                            ''@rosadinnn
                            ''@MUaT
                            row1.TotalQty_Multi += CDec(row.Cells(Me.ColQtyMulti.Index).Value)

                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then

                        ''@rosadinnn
                        ''@MUaT
                        'Me.TBL4.Rows.Add(Me.TxtPrimaryStockIn.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        'row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColQty.Index).Value, row.Cells(Me.ColBatch.Index).Value, _
                        'row.Cells(Me.ColPrimaryPOno.Index).Value, row.Cells(Me.ColPONo.Index).Value, Me.TxtPrimaryPihakAcuan.Text, _
                        'Me.TxtKodePihakAcuan.Text, CDec(row.Cells(Me.ColQtyMulti.Index).Value))
                        row1 = Me.TBL4.Newtbltransaksistockinpurchase_viewMultiRow
                        With row1
                            .Primary_Transaksi = Me.TxtPrimaryStockIn.Text
                            .Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                            .Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                            .TotalQty = row.Cells(Me.ColQty.Index).Value
                            .Batch_No = row.Cells(Me.ColBatch.Index).Value
                            .Primary_PO_Purchase = row.Cells(Me.ColPrimaryPOno.Index).Value
                            .No_PO = row.Cells(Me.ColPONo.Index).Value
                            .Primary_Supplier = Me.TxtPrimaryPihakAcuan.Text
                            .Supplier_Code = Me.TxtKodePihakAcuan.Text
                            .TotalQty_Multi = row.Cells(Me.ColQtyMulti.Index).Value
                            .Primary_Unit_Multi = row.Cells(Me.ColPrimaryUnitMulti.Index).Value
                            .Unit_Code_Multi = row.Cells(Me.ColUnitCodeMulti.Index).Value
                            .Ratio_Reference = row.Cells(Me.ColRatioReference.Index).Value
                            .Ration_Default = row.Cells(Me.ColRatioDefault.Index).Value
                            .Primary_Unit_Default = row.Cells(Me.ColPrimaryUnits.Index).Value
                            .Unit_Code_Default = row.Cells(Me.ColUnit.Index).Value
                            .Tanggal_Transaksi = Me.dtpDate.Value
                        End With
                        Me.TBL4.Rows.Add(row1)

                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub IsiCboPurchaseOrder(ByVal vPrimarySupplier As Decimal)
        Dim SortingPO As Boolean = False
        Dim TypeReset As Integer

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data awal
            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()

            'isikan dari basis data-persedian terakhir
            If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then

                ''@rosadinnn
                ''@MUaT
                If Me.ChkIgnoreTolerance.Checked = False Then
                    Me.objPOItem.IsikanTBLPO_Multi_Tolerance(vPrimarySupplier, Me.TBL, Me.cboPurchaseOrder)
                Else
                    Me.objPOItem.IsikanTBLPO_Multi(vPrimarySupplier, Me.TBL, Me.cboPurchaseOrder)
                End If

            Else
                Me.IsikanTBLPODariPackingList_Multi(vPrimarySupplier, Me.TxtPrimaryPackingList.Text)
            End If

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiRow
            Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow
            Dim ketemu As Boolean

            If Me.baru = False Then

                ''@rosadinnn
                ''@MUaT
                ''@untuk pob dari packing list jangan ditambahkan dengan jmlobjlama
                If Me.TxtPrimaryPackingList.Text = "-1" Or Me.TxtPrimaryPackingList.Text = "" Then
                    Me.ObjStockInOutItem.IsiTBLPO_Multi(Me.TxtPrimaryStockIn.Text, Me.TBL2)
                    'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                    While Me.TBL2.Rows.Count > 0
                        row1 = Me.TBL2.Rows(0)
                        ketemu = False
                        For Each row2 In Me.TBL.Rows
                            If row1.Primary_Supplier = row2.Primary_Supplier _
                            And row1.Primary_Product = row2.Primary_Product And _
                            row2.Primary_PO = row1.Primary_PO_Purchase Then
                                row2.OrderBalance = row2.OrderBalance + row1.TotalQty

                                ''@rosadinnn
                                ''@MUaT
                                row2.OrderBalance_Multi = row2.OrderBalance_Multi + row1.TotalQty_Multi

                                ketemu = True
                                Exit For
                            End If
                        Next
                        If ketemu = False Then
                            'tambahkan baris baru

                            ''@rosadinnn
                            ''@MUaT
                            row2 = TBL.Neworderbalance_report_view0MultiRow
                            row2.Primary_PO = row1.Primary_PO_Purchase
                            row2.Primary_Supplier = row1.Primary_Supplier
                            row2.Primary_Product = row1.Primary_Product
                            row2.Product_Code = row1.Product_Code
                            row2.PO_Number = row1.No_PO
                            row2.OrderBalance = row1.TotalQty
                            row2.Primary_Quick_Acuan = -1
                            row2.OrderBalance_Multi = row1.TotalQty_Multi
                            row2.Primary_Unit_Multi = row1.Primary_Unit_Multi
                            row2.Unit_Code_Multi = row1.Unit_Code_Multi
                            row2.Ratio_Default = row1.Ration_Default
                            row2.Ratio_Reference = row1.Ratio_Reference
                            row2.Primary_Unit_Default = row1.Primary_Unit_Default
                            row2.Unit_Code_Default = row1.Unit_Code_Default
                            row2.Tolerance_Down = 0
                            row2.Tolerance_Up = 0
                            row2.Tanggal_Transaksi = row1.Tanggal_Transaksi
                            Me.TBL.Rows.Add(row2)

                        End If
                        Me.TBL2.Rows.Remove(row1)
                    End While
                End If
            End If

            'kurangkan item yang sudah dimasukkan ke dgv semua

            Me.IsikanTBLSO()
            While Me.TBL2.Rows.Count > 0
                row1 = Me.TBL2.Rows(0)
                ketemu = False
                For Each row2 In Me.TBL.Rows
                    If row1.Primary_Supplier = row2.Primary_Supplier _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_PO = row1.Primary_PO_Purchase Then
                        row2.OrderBalance = row2.OrderBalance - row1.TotalQty

                        ''@rosadinnn
                        ''@MUaT
                        row2.OrderBalance_Multi -= row1.TotalQty_Multi

                        'If row2.OrderBalance = 0 Then
                        '    Me.TBL.Rows.Remove(row2)
                        'End If

                        ''@rosadinnn
                        ''@MUaT
                        ''ubah if diatas (asumsi ob sekarang adalah ob multi)
                        'If row2.OrderBalance_Multi = 0 Then
                        '    Me.TBL.Rows.Remove(row2)
                        'End If

                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL2.Rows.Remove(row1)
            End While
            For i As Integer = TBL.Rows.Count - 1 To 0 Step -1
                row2 = TBL.Rows(i)
                If Me.ChkIgnoreTolerance.Checked = True Then
                    If row2.OrderBalance_Multi + row2.Tolerance_Up <= 0 Then
                        Me.CekCbo = False
                        Me.TBL.Rows.Remove(row2)
                        Me.CekCbo = True
                    End If
                Else
                    If row2.OrderBalance_Multi - row2.Tolerance_Down <= 0 Then
                        Me.CekCbo = False
                        Me.TBL.Rows.Remove(row2)
                        Me.CekCbo = True
                    End If
                End If
            Next


            'pastikan hanya 1 so disebutkan hanya 1x
            Dim idx As Integer = 0
            Dim idx2 As Integer = 0
            Dim dummyRow As DseTransaksiPO.orderbalance_report_view0MultiRow
            While idx < Me.TBL.Rows.Count
                row2 = Me.TBL(idx)
                idx2 = idx + 1
                While idx2 <= Me.TBL.Rows.Count - 1
                    dummyRow = Me.TBL.Rows(idx2)
                    If dummyRow.PO_Number = row2.PO_Number Then
                        Me.CekCbo = False
                        Me.TBL.Rows.Remove(dummyRow)
                        Me.CekCbo = True
                    Else
                        idx2 += 1
                    End If
                End While

                'pindah ke baris selanjutnya
                idx += 1
            End While

            'berkaitan pengurutan po bila no po autonumber (SMESS)
            If ObjSupplier.ApakahLokal(Me.TxtPrimaryPihakAcuan.Text) = True Then
                If ObjNumb.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                    If ObjNumb.TypeResetAutonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = ClsMasterTransactionNumbering.EnumReset.Monthly Then
                        TypeReset = ClsMasterTransactionNumbering.EnumReset.Monthly
                    Else
                        TypeReset = ClsMasterTransactionNumbering.EnumReset.Yearly
                    End If
                    SortingPO = True
                End If
            Else
                If ObjNumb.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) = True Then
                    If ObjNumb.TypeResetAutonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) = ClsMasterTransactionNumbering.EnumReset.Monthly Then
                        TypeReset = ClsMasterTransactionNumbering.EnumReset.Monthly
                    Else
                        TypeReset = ClsMasterTransactionNumbering.EnumReset.Yearly
                    End If
                    SortingPO = True
                End If
            End If

            If SortingPO = True Then
                'jika penomoran po : autonumber maka sorting berdasarkan type reset number (monthly/yearly)
                Dim row As DseTransaksiPO.orderbalance_report_view0MultiRow

                For Each row In Me.TBL
                    If TypeReset = ClsMasterTransactionNumbering.EnumReset.Monthly Then
                        row.PO_Date_Sort = Format(row.Tanggal_Transaksi, "MM/yyyy")
                    ElseIf TypeReset = ClsMasterTransactionNumbering.EnumReset.Yearly Then
                        row.PO_Date_Sort = Format(row.Tanggal_Transaksi, "yyyy/MM")
                    End If
                Next

                Me.TBL.DefaultView.Sort = "[PO Date Sort] ASC, [PO Number] ASC"
            End If

            Me.cboPurchaseOrder.DataSource = Me.TBL.DefaultView
            Me.cboPurchaseOrder.ValueMember = "PO Number"
            Me.cboPurchaseOrder.DisplayMember = "PO Number"
            Me.cboPurchaseOrder.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    ''@rosadinnn
    ''MUaT
    Private Sub IsikanTBLPODariPackingList_Multi(ByVal vPrimarySupplier As Decimal, ByVal vPrimaryPackingList As Decimal)
        Dim rowPL As DseTransaksiPackingList.tbltrans_packinglist_view1MultiRow
        Dim rowPOB As DseTransaksiPO.orderbalance_report_view0MultiRow
        Dim TBL5_Dum As New DseTransaksiPackingList.tbltrans_packinglist_view1MultiDataTable

        Try
            Me.ObjPackingList.IsiTBLPackingList(vPrimarySupplier, _
            vPrimaryPackingList, Me.TBL5)
            TBL5_Dum = Me.TBL5.Copy
            For Each rowPL In TBL5_Dum.Rows
                rowPOB = Me.TBL.Neworderbalance_report_view0MultiRow
                With rowPL
                    rowPOB.OrderBalance = .Order_Balance
                    rowPOB.PO_Number = .PO_Number
                    rowPOB.Primary_PO = .Primary_PO
                    rowPOB.Primary_Product = .Primary_Product
                    rowPOB.Primary_Supplier = .Primary_Supplier
                    rowPOB.Product_Code = .Product_Code

                    ''@rosadinnn
                    ''@MUaT
                    rowPOB.Primary_Quick_Acuan = -1
                    rowPOB.OrderBalance_Multi = .Order_Balance_Multi
                    rowPOB.Primary_Unit_Multi = .Primary_Unit_Multi
                    rowPOB.Unit_Code_Multi = .Unit_Code_Multi
                    rowPOB.Ratio_Reference = .Ratio_Reference
                    rowPOB.Ratio_Default = .Ratio_Default
                    rowPOB.Primary_Unit_Default = .Primary_Unit_Default
                    rowPOB.Unit_Code_Default = .Unit_Code_Default
                    rowPOB.Tolerance_Down = 0
                    rowPOB.Tolerance_Up = 0
                    'rowPOB.PO_Number = ""
                    rowPOB.Tanggal_Transaksi = "01/01/1950"

                End With
                Me.TBL.Rows.Add(rowPOB)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub IsikanTBLSO()
    '    Try
    '        Dim row As DataGridViewRow
    '        Dim row1 As dseTransaksiStockInOut.tbltransaksistockinpurchase_viewRow
    '        For Each row In Me.dgvItem.Rows
    '            Dim ketemu As Boolean = False
    '            For Each row1 In Me.TBL2.Rows
    '                If row1.Primary_Supplier = Me.TxtPrimaryPihakAcuan.Text And _
    '                row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
    '                row1.Primary_PO_Purchase = row.Cells(Me.ColPrimaryPOno.Index).Value Then
    '                    'kalo ketemu ubah data order balance
    '                    row1.TotalQty = row1.TotalQty + CDec(row.Cells(Me.ColQty.Index).Value)
    '                    ketemu = True
    '                    Exit For
    '                End If
    '            Next
    '            'kalo ga ketemu tambahkan baris baru
    '            If ketemu = False Then
    '                Me.TBL2.Rows.Add(Me.TxtPrimaryStockIn.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
    '               row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColQty.Index).Value, row.Cells(Me.ColBatch.Index).Value, _
    '                row.Cells(Me.ColPrimaryPOno.Index).Value, row.Cells(Me.ColPONo.Index).Value, Me.TxtPrimaryPihakAcuan.Text, Me.TxtKodePihakAcuan.Text)
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub IsikanTBLSO()
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiRow
            For Each row In Me.dgvItem.Rows
                Dim ketemu As Boolean = False
                For Each row1 In Me.TBL2.Rows
                    If row1.Primary_Supplier = Me.TxtPrimaryPihakAcuan.Text And _
                    row1.Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                    row1.Primary_PO_Purchase = row.Cells(Me.ColPrimaryPOno.Index).Value Then
                        'kalo ketemu ubah data order balance
                        row1.TotalQty = row1.TotalQty + CDec(row.Cells(Me.ColQty.Index).Value)

                        ''@rosadinnn
                        ''@MUaT
                        row1.TotalQty_Multi += CDec(row.Cells(Me.ColQtyMulti.Index).Value)

                        ketemu = True
                        Exit For
                    End If
                Next
                'kalo ga ketemu tambahkan baris baru
                If ketemu = False Then

                    ''@rosadinnn
                    ''@MUAT
                    ' Me.TBL2.Rows.Add(Me.TxtPrimaryStockIn.Text, row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'row.Cells(Me.ColCodeNo.Index).Value, row.Cells(Me.ColQty.Index).Value, row.Cells(Me.ColBatch.Index).Value, _
                    ' row.Cells(Me.ColPrimaryPOno.Index).Value, row.Cells(Me.ColPONo.Index).Value, Me.TxtPrimaryPihakAcuan.Text, Me.TxtKodePihakAcuan.Text, CDec(row.Cells(Me.ColQtyMulti.Index).Value))
                    row1 = Me.TBL2.Newtbltransaksistockinpurchase_viewMultiRow
                    With row1
                        .Primary_Transaksi = Me.TxtPrimaryStockIn.Text
                        .Primary_Product = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                        .Product_Code = row.Cells(Me.ColCodeNo.Index).Value
                        .TotalQty = row.Cells(Me.ColQty.Index).Value
                        .Batch_No = row.Cells(Me.ColBatch.Index).Value
                        .Primary_PO_Purchase = row.Cells(Me.ColPrimaryPOno.Index).Value
                        .No_PO = row.Cells(Me.ColPONo.Index).Value
                        .Primary_Supplier = Me.TxtPrimaryPihakAcuan.Text
                        .Supplier_Code = Me.TxtKodePihakAcuan.Text
                        .TotalQty_Multi = row.Cells(Me.ColQtyMulti.Index).Value
                        .Primary_Unit_Multi = row.Cells(Me.ColPrimaryUnitMulti.Index).Value
                        .Unit_Code_Multi = row.Cells(Me.ColUnitCodeMulti.Index).Value
                        .Ratio_Reference = row.Cells(Me.ColRatioReference.Index).Value
                        .Ration_Default = row.Cells(Me.ColRatioDefault.Index).Value
                        .Primary_Unit_Default = row.Cells(Me.ColPrimaryUnits.Index).Value
                        .Unit_Code_Default = row.Cells(Me.ColUnit.Index).Value
                        .Tanggal_Transaksi = Me.dtpDate.Value
                    End With
                    Me.TBL2.Rows.Add(row1)

                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub IsiCboProduct(ByVal vPrimaryPO As Decimal)
        Try
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'isikan dari basis data-persedian terakhir
            Me.cboCodeNo.DataSource = Nothing
            If Me.TxtPrimaryPackingList.Text = "" Or Me.TxtPrimaryPackingList.Text = "-1" Then
                Me.objPOItem.IsikanTBLProduct(Me.txtPrimaryPO.Text, Me.TBL3, Me.cboCodeNo, Me.ChkIgnoreTolerance.Checked)
            Else
                Me.IsikanTBLProductDariPackingList(Me.TxtPrimaryPackingList.Text, vPrimaryPO)
            End If

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiRow
            Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjStockInOutItem.IsiTBLProductByPO(Me.TxtPrimaryStockIn.Text, _
                Me.TBL4, Me.txtPrimaryPO.Text)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    ketemu = False
                    For Each row2 In Me.TBL3.Rows
                        If row1.Primary_Supplier = row2.Primary_Supplier _
                        And row1.Primary_Product = row2.Primary_Product And _
                        row2.Primary_PO = row1.Primary_PO_Purchase Then
                            row2.OrderBalance = row2.OrderBalance + row1.TotalQty

                            ''@rosadinnn
                            ''@MUaT
                            row2.OrderBalance_Multi += row1.TotalQty_Multi

                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru

                        ''@rosadinnn
                        ''@MUaT
                        'Me.TBL3.Rows.Add(row1.Primary_PO_Purchase, row1.Primary_Supplier, row1.Primary_Product, _
                        'row1.Product_Code, row1.No_PO, row1.TotalQty, CDec(row1.TotalQty_Multi))
                        row2 = Me.TBL3.Neworderbalance_report_view0MultiRow
                        With row2
                            .Primary_PO = row1.Primary_PO_Purchase
                            .Primary_Supplier = row1.Primary_Supplier
                            .Primary_Product = row1.Primary_Product
                            .Product_Code = row1.Product_Code
                            .PO_Number = row1.No_PO
                            .OrderBalance = row1.TotalQty
                            .Primary_Quick_Acuan = -1
                            .OrderBalance_Multi = row1.TotalQty_Multi
                            .Primary_Unit_Multi = row1.Primary_Unit_Multi
                            .Unit_Code_Multi = row1.Unit_Code_Multi
                            .Ratio_Default = row1.Ration_Default
                            .Ratio_Reference = row1.Ratio_Reference
                            .Primary_Unit_Default = row1.Primary_Unit_Default
                            .Unit_Code_Default = row1.Unit_Code_Default
                            row2.Tolerance_Down = 0
                            row2.Tolerance_Up = 0
                            .Tanggal_Transaksi = row1.Tanggal_Transaksi
                        End With
                        Me.TBL3.Rows.Add(row2)

                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If

            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsiTBLProduct(Me.txtPrimaryPO.Text)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                ketemu = False
                For Each row2 In Me.TBL3.Rows
                    If row1.Primary_Supplier = row2.Primary_Supplier _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_PO = row1.Primary_PO_Purchase Then
                        row2.OrderBalance = row2.OrderBalance - row1.TotalQty

                        ''@rosadinnn
                        ''@MUaT 
                        ''pindahkan penghapusan dibawan ke pengulangan yg memperhitungkan toleransi
                        row2.OrderBalance_Multi -= row1.TotalQty_Multi
                        '' -> : mengganti if dibawah (asumsi ob sekarang adalah ob multi)
                        'If row2.OrderBalance_Multi = 0 Then
                        '    Me.TBL3.Rows.Remove(row2)
                        'End If

                        'If row2.OrderBalance = 0 Then
                        '    Me.TBL3.Rows.Remove(row2)
                        'End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL3.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4.Rows.Remove(row1)
            End While

            ''@rosadinnn
            ''@MUaT
            For i As Integer = TBL3.Rows.Count - 1 To 0 Step -1
                row2 = TBL3.Rows(i)
                If Me.ChkIgnoreTolerance.Checked = True Then
                    If row2.OrderBalance_Multi + row2.Tolerance_Up <= 0 Then
                        Me.CekCbo = False
                        Me.TBL3.Rows.Remove(row2)
                        Me.CekCbo = True
                    End If
                Else
                    If row2.OrderBalance_Multi - row2.Tolerance_Down <= 0 Then
                        Me.CekCbo = False
                        Me.TBL3.Rows.Remove(row2)
                        Me.CekCbo = True
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

    'Private Sub IsikanTBLProductDariPackingList(ByVal vPrimaryPackingList As Decimal, _
    'ByVal vPrimaryPO As Decimal)
    '    Dim rowPL As DseTransaksiPackingList.tbltrans_packinglist_view1Row
    '    Dim rowPOB As DseTransaksiPO.orderbalance_report_view0Row

    '    Try
    '        'isikan order balance dari packing list
    '        Me.ObjPackingList.IsiTBLPackingList_PO(vPrimaryPackingList, Me.TBL5, vPrimaryPO)
    '        'masukkan ke tabel 1
    '        For Each rowPL In Me.TBL5.Rows
    '            rowPOB = Me.TBL3.Neworderbalance_report_view0Row
    '            With rowPL
    '                rowPOB.OrderBalance = .Order_Balance
    '                rowPOB.PO_Number = .PO_Number
    '                rowPOB.Primary_PO = .Primary_PO
    '                rowPOB.Primary_Product = .Primary_Product
    '                rowPOB.Primary_Supplier = .Primary_Supplier
    '                rowPOB.Product_Code = .Product_Code
    '            End With
    '            Me.TBL3.Rows.Add(rowPOB)
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub IsikanTBLProductDariPackingList(ByVal vPrimaryPackingList As Decimal, _
    ByVal vPrimaryPO As Decimal)
        Dim rowPL As DseTransaksiPackingList.tbltrans_packinglist_view1MultiRow
        Dim rowPOB As DseTransaksiPO.orderbalance_report_view0MultiRow

        Try
            'isikan order balance dari packing list
            Me.ObjPackingList.IsiTBLPackingList_PO(vPrimaryPackingList, Me.TBL5, vPrimaryPO)
            'masukkan ke tabel 1
            For Each rowPL In Me.TBL5.Rows
                rowPOB = Me.TBL3.Neworderbalance_report_view0MultiRow
                With rowPL
                    rowPOB.OrderBalance = .Order_Balance
                    rowPOB.PO_Number = .PO_Number
                    rowPOB.Primary_PO = .Primary_PO
                    rowPOB.Primary_Product = .Primary_Product
                    rowPOB.Primary_Supplier = .Primary_Supplier
                    rowPOB.Product_Code = .Product_Code

                    ''@rosadinnn
                    ''@MUaT
                    rowPOB.Primary_Quick_Acuan = -1
                    rowPOB.OrderBalance_Multi = .Order_Balance_Multi
                    rowPOB.Primary_Unit_Multi = .Primary_Unit_Multi
                    rowPOB.Unit_Code_Multi = .Unit_Code_Multi
                    rowPOB.Ratio_Reference = .Ratio_Reference
                    rowPOB.Ratio_Default = .Ratio_Default
                    rowPOB.Primary_Unit_Default = .Primary_Unit_Default
                    rowPOB.Unit_Code_Default = .Unit_Code_Default
                    rowPOB.Tolerance_Down = 0
                    rowPOB.Tolerance_Up = 0
                    rowPOB.PO_Number = ""
                    rowPOB.Tanggal_Transaksi = "01/01/1950"

                End With
                Me.TBL3.Rows.Add(rowPOB)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Try
            If Me.cboCodeNo.SelectedIndex <> -1 Then
                Dim RowView As Data.DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
                Dim row2 As DseTransaksiPO.orderbalance_report_view0MultiRow = RowView.Row
                If Me.CekCbo = True Then
                    Dim PrimaryCodeNoBaru As Double = row2.Primary_Product
                    If Me.TxtPrimaryCodeNo.Text <> CStr(PrimaryCodeNoBaru) Then

                        Me.TxtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)

                        Me.txtPrimaryUnits.Text = row2.Primary_Unit_Multi
                        Me.lblUnits.Text = row2.Unit_Code_Multi
                        Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping("0", ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))

                        Me.TxtPrimaryUnitDefault.Text = row2.Primary_Unit_Default
                        Me.LblUnitDefault.Text = row2.Unit_Code_Default
                        Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping("0", Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))

                        Me.TxtRatioReference.Text = Me.ObjCD.DigitGrouping(row2.Ratio_Reference, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))
                        Me.TxtRatioDefault.Text = Me.ObjCD.DigitGrouping(row2.Ratio_Default, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))

                        Dim idx As Integer = -1
                        idx = Me.GetIndexRatio(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text)
                        If Me.txtPrimaryUnits.Text <> Me.TxtPrimaryUnitDefault.Text Then
                            If idx <> -1 Then
                                Me.LblQty.Text = "Qty (" & Me.dgvItem.Item(Me.ColRatioDesc.Index, idx).Value & ")"
                                Me.TxtRatioReference.Text = Me.dgvItem.Item(Me.ColRatioReference.Index, idx).Value
                                Me.TxtRatioDefault.Text = Me.dgvItem.Item(Me.ColRatioDefault.Index, idx).Value
                            Else
                                Me.LblQty.Text = "Qty (" & Me.ObjCD.DigitGrouping(row2.Ratio_Reference, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)) & _
                                " " & row2.Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(row2.Ratio_Default, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) & _
                                " " & row2.Unit_Code_Default & ")"
                            End If
                            Me.TxtQtyDefault.Visible = True
                            Me.LblUnitDefault.Visible = True
                        Else
                            Me.LblQty.Text = "Qty"
                            Me.TxtQtyDefault.Visible = False
                            Me.LblUnitDefault.Visible = False
                        End If

                        Me.LblUnits2.Text = Me.lblUnits.Text

                        Me.SetOrderBalance()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = ""

                Me.txtPrimaryUnits.Text = "-1"
                Me.lblUnits.Text = "Unit"
                Me.TxtQtyperBatch.Text = "0"

                Me.TxtPrimaryUnitDefault.Text = "-1"
                Me.LblUnitDefault.Text = "Unit"
                Me.TxtQtyDefault.Text = "0"

                Me.TxtRatioReference.Text = "0"
                Me.TxtRatioDefault.Text = "0"

                Me.LblQty.Text = "Qty"
                Me.TxtOrderBalance.Text = "0"
            End If

            Me.SetGboItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppDONumber.GotFocus
        If Me.TxtSuppDONumber.Enabled = True And Me.TxtRemark.Enabled = True Then
            If Me.TxtSuppDONumber.ReadOnly = False Then
                Me.TxtSuppDONumber.BackColor = Me.BackColorAll
            End If
        End If
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuppDONumber.LostFocus
        If Me.TxtSuppDONumber.Enabled = True Then
            If Me.TxtSuppDONumber.ReadOnly = False Then
                Me.TxtSuppDONumber.BackColor = Me.TxtPrimaryPackingList.BackColor
            End If
        End If
    End Sub

    Private Sub SetOrderBalance()
        Try
            If Me.TxtPrimaryPackingList.Text <> "-1" Or Me.TxtPrimaryPackingList.Text = "" Then

                ''@rosadinnn
                ''@MUaT
                ''@belum -> dikerjakan sekalian ngerjain packing list

                'berdasarkan packing list
                'pastikan nilai qty item packing list yang belum masuk stock in tidak lebih besar dari batas yang
                'mengakibatkan nilai order balance
                Dim OB_PackingList As Decimal = Me.objPackingListItem.GetJml_Multi(Me.TxtPrimaryPackingList.Text, Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) _
                        - Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) '- CDec(Me.TxtQtyperBatch.Text)

                Dim OB_Purchase As Decimal = Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) _
                                + Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, _
                                Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) - Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text)

                If Me.baru = False And Me.ObjPackingList.GetNoStockInPurchase(Me.TxtPrimaryPackingList.Text) <> "" Then
                    'jika order balance secara keseluruhan lebih kecil dari order balance packing list
                    'maka cek, jangan sampe kelebihan masukkin angka order balance yang
                    'dilihat oleh user

                    ''@rosadinnn
                    ''@MUaT
                    ''@hilangakan pengalihan ke OB PO
                    'If OB_Purchase < OB_PackingList Then
                    '    'jika menggunakan order balance packing list akan mengakibatkan order balance minus
                    '    'gunakan order balance keseluruhan

                    '    ''@rosadinnn
                    '    ''@MUaT
                    '    ''@update DefPrec
                    '    Me.TxtOrderBalance.Text = Me.ObjCD.DigitGrouping(OB_Purchase, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))

                    'Else
                    '    'gunakan order balance packing list

                    '    ''@rosadinnn
                    '    ''@MUaT
                    '    ''@update DefPrec
                    Me.TxtOrderBalance.Text = Me.ObjCD.DigitGrouping(OB_PackingList, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))

                    'End If
                Else
                    'jika order balance keseluruhan lebih besar atau sama dengan order balance packing
                    'list maka gunakan order balance packing list

                    ''@rosadinnn
                    ''@MUaT
                    ''@update DefPrec
                    Me.TxtOrderBalance.Text = Me.ObjCD.DigitGrouping(OB_PackingList, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))

                End If
            Else
                'berdasarkan po number
                If IsNumeric(Me.TxtPrimaryCodeNo.Text) = True And IsNumeric(Me.txtPrimaryPO.Text) = True Then
                    '   Me.TxtOrderBalance.Text = Me.ObjCD.DigitGrouping(Me.objPOItem.GetJmlPurchaseOrderBalance(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) _
                    '+ Me.ObjStockInOutItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) - _
                    'Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))

                    ''@rosadinnn
                    ''@MUaT
                    Me.TxtOrderBalance.Text = Me.ObjCD.DigitGrouping(Me.objPOItem.GetJmlPurchaseOrderBalance_Multi(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) _
                                        + Me.ObjStockInOutItem.GetJml_Multi(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text) - _
                                        Me.GetJmlDGV_Multi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryPO.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")

            ''@rosadinnn
            ''@MUaT
        Finally
            Me.SetFieldToleranceUp()

        End Try
    End Sub

    Private Sub txtBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.Leave
        Try
            Me.txtBatchNo.Text = Me.txtBatchNo.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub LblQty_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblQty.Click
        Dim frmRatio As New FrmTransaksiStockInPurchase2_SetQtyRatio
        Try
            If Me.TxtRemark.Enabled = False Or Me.LblQty.Text = "Qty" Then
                Exit Sub
            End If
            If Me.GetIndexRatio(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text) <> -1 Then
                MsgBox("There is already ratio defined below. You cannot set different ratio for exist product.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)
            frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
            frmRatio.vRatioMulti = Me.TxtRatioReference.Text
            frmRatio.vUnitMulti = Me.lblUnits.Text
            frmRatio.vRatioDefault = Me.TxtRatioDefault.Text
            frmRatio.vUnitDefault = Me.LblUnitDefault.Text
            If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TxtRatioReference.Text = frmRatio.TxtRatioMulti.Text
                Me.TxtRatioDefault.Text = frmRatio.TxtRatioDefault.Text
                Me.LblQty.Text = "Qty (" & frmRatio.TxtRatioMulti.Text & " " & frmRatio.LlbUnitMulti.Text & " = " & frmRatio.TxtRatioDefault.Text & " " & frmRatio.LlbUnitDefault.Text & ")"
                SetQtyDefault()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Try
            Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatch.Text) * CDec(Me.TxtRatioDefault.Text) / CDec(Me.TxtRatioReference.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyMulti()
        Try
            Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyDefault.Text) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub TxtQtyDefault_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.GotFocus
        Try
            Me.TxtQtyDefault.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyDefault_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyDefault.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyDefault, sender, e)
    End Sub

    Private Sub TxtQtyDefault_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyDefault.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyDefault, sender, e)
    End Sub

    Private Sub TxtQtyDefault_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyDefault_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.Leave
        Try
            Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(TxtQtyDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
            SetQtyMulti()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexRatio(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Integer
        Try
            For Each rowDgv As DataGridViewRow In Me.dgvItem.Rows
                With rowDgv
                    If .Cells(Me.ColPrimaryPOno.Index).Value = vPrimaryPO And .Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
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
                                                            CDec(Me.dgvItem.Item(Me.ColRatioReference.Index, vRowIdx).Value / _
                                                            CDec(Me.dgvItem.Item(Me.ColRatioDefault.Index, vRowIdx).Value)), _
                                                            ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, vRowIdx).Value, _
                                                            Me.dgvItem.Item(Me.ColPrimaryUnitMulti.Index, vRowIdx).Value))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvPackingList_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvPackingList.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")

        End Try
    End Sub

    Private Sub TxtQtyDefault_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.LostFocus
        Try
            Me.TxtQtyDefault.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetFieldToleranceUp()
        Dim DefPrec As Integer
        Try
            DefPrec = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnits.Text)
            If Me.TxtPrimaryPackingList.Text = "-1" Or Me.TxtPrimaryPackingList.Text = "" Then
                Me.ChkIgnoreTolerance.Enabled = True
                Me.TxtToleranceUp.Text = ObjCD.DigitGrouping(Me.objPO.GetToleranceUp(Me.txtPrimaryPO.Text, Me.TxtPrimaryCodeNo.Text), DefPrec)
            Else
                Me.ChkIgnoreTolerance.Enabled = False
                Me.ChkIgnoreTolerance.Checked = True
                Me.TxtToleranceUp.Text = ObjCD.DigitGrouping(0, DefPrec)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkIgnoreTolerance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkIgnoreTolerance.CheckedChanged
        Try
            If Me.TxtPrimaryPihakAcuan.Text <> "" And Me.TxtPrimaryPihakAcuan.Text <> "-1" And Me.TxtRemark.Enabled = True Then
                Me.SetCBO()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtOrderBalance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtOrderBalance.TextChanged
        Try
            If Me.TxtOrderBalance.Text = "0" Then
                Me.TxtToleranceUp.Text = "0"
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
                i.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInPurchase
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
                If Me.objStockIn.CariDetail(vPrimary, Me.dgvItem, Me.TblItemCustom) _
                 = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub EditTanggal(ByVal vTglTujuan As Date)
        'Dim hasil As MsgBoxResult
        Try
            If Me.TxtStockInNo.Text <> "" Then
                'If Me.objStockIn.ApakahBolehEdit(Me.UserCode) = True Then
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
                    '    If Me.objStockIn.SimpanBaru(Me.dgvItem, Me.UserCode) Then
                    '        MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    '        Me.SalinKeField()
                    '        Me.NonAktifkanDetail()
                    '        Me.btnFind.Focus()
                    '    End If
                    'End If
                Else
                    'konfirmasi disimpan atau tidak
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.objStockIn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                        'MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                        Me.SalinKeField()
                        Me.NonAktifkanDetail()
                        Me.btnEdit.Focus()
                    Else
                        MsgBox("Logic error.")
                    End If
                End If
                'End If
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_StockInPurch_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal vTblPrimaryPO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryPL As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core)
        Dim TblTempPO, TblTempPL As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.objStockIn.ResetObject(Me.dgvItem)
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            Me.objStockIn.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.TblItemCustom)

            'Update Primary Packing List & Primary PO
            Me.objStockIn.PrimaryGIT = ObjCutOff_Core.GetPrimaryUpdate(Me.objStockIn.PrimaryGIT, TblTempPL)
            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryPOno.Index).Value) = False And rowD.Cells(Me.ColPrimaryPOno.Index).Value <> "-1" Then
                    TblTempPO.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryPOno.Index).Value)
                    If TblTempPO.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOno.Index).Value), vTblPrimaryPO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOno.Index).Value), PrimaryUpdate, TblTempPO)
                        rowD.Cells(Me.ColPrimaryPOno.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryPOno.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryPOno.Index).Value), TblTempPO)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_StockInPurch_Save() As Double
        Try
            Me.baru = True
            Me.objStockIn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, -1, False, False, True)
            Return Me.objStockIn.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub LblQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblQty.TextChanged
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(Me.LinkLabel1, 0, 10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiStockInPurchase2_AddByCustomProduct

        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryPihakAcuan.Text = -1 And Me.TxtPrimaryPihakAcuan.Text <> "" Then
                    MsgBox("To start entering item data, you must specify supplier data first.", MsgBoxStyle.Information, "Information")
                    Me.TxtKodePihakAcuan.Focus()
                    Exit Sub
                End If
            End If

            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TblItemCustom = Me.TblItemCustom.Copy
            i.PrimaryTransaksi = Me.TxtPrimaryStockIn.Text
            i.TglTransaksi = Me.dtpDate.Value
            i.PrimarySupplier = Me.TxtPrimaryPihakAcuan.Text
            i.SupplierCode = Me.TxtPihakAcuan.Text
            i.Baru = Me.baru
            'i.PrimaryPackingList = Me.txtPrimaryDO.Text
            If Me.TxtRemark.Enabled = False Then
                i.CanEditForm = False
            End If

            If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblItemCustom.Clear()
                Me.TblItemCustom = i.TblItemCustom
                Me.IsikanDgvFromTblDummy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini dibuat untuk admiral
    Private Function IsikanDgvFromTblDummy() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemRow
        Dim ObjP As New Core_Module.clsMasterProduct
        Dim rd As DataGridViewRow
        Dim DummyPO As String = ""
        Dim PrimaryLocationLama As Double = -1
        Dim IdxRows As Integer = -1

        Try
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each row In TblItemCustom.Rows
                With row
                    IdxRows = Me.GetIndex(.Primary_Product, .Primary_Storage_Area, .Batch_No, _
                    -1, .Primary_PdO, .Expiry_Date, .Primary_PO_Purchase)
                    PrimaryLocationLama = Me.txtPrimaryLocation.Text
                    'cari data yang sama persis
                    If IdxRows <> -1 Then
                        'tambahkan nilai qty
                        rd = Me.dgvItem.Rows(IdxRows)
                        rd.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQtyMulti.Index).Value) + .Qty_Multi_Unit, .dec_multi)
                        If .Is_Kain = True Then
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQty.Index).Value) + (.Qty_Multi_Unit * .Qty_At), .dec)
                        Else
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQty.Index).Value) + .Qty, .dec)
                        End If
                        If .Is_Kain = False Then
                            rd.Cells(Me.ColRatioDefault.Index).Value = .Ratio_Default
                            rd.Cells(Me.ColRatioReference.Index).Value = .Rasio
                            rd.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Rasio, .dec_multi) & " " & .unit_multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, .dec) & " " & .unit
                        Else
                            rd.Cells(Me.ColRatioDefault.Index).Value = rd.Cells(Me.ColQty.Index).Value
                            rd.Cells(Me.ColRatioReference.Index).Value = rd.Cells(Me.ColQtyMulti.Index).Value
                            rd.Cells(Me.ColRatioDesc.Index).Value = rd.Cells(Me.ColQtyMulti.Index).Value & " " & .unit_multi & " = " & rd.Cells(Me.ColQty.Index).Value & " " & .unit
                        End If

                        If String.IsNullOrEmpty(rd.Cells(Me.ColRemark.Index).Value) = False Then
                            rd.Cells(Me.ColRemark.Index).Value = rd.Cells(Me.ColRemark.Index).Value & ", " & .Remarks
                        Else
                            rd.Cells(Me.ColRemark.Index).Value = .Remarks
                        End If
                    Else
                        'tambahkan baris baru
                        Dim ExpiredDateStr As String
                        If .Expiry_Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(.Expiry_Date, "dd MMM yyyy")
                        End If

                        rd = Me.dgvItem.Rows(Me.dgvItem.Rows.Add())
                        rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColPONo.Index).Value = .po_so_number
                        rd.Cells(Me.ColSuppDONo.Index).Value = .Supp__DO_No
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Multi_Unit, .dec_multi)
                        If .Is_Kain = True Then
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Multi_Unit * .Qty_At, .dec)
                        Else
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, .dec)
                        End If
                        rd.Cells(Me.ColUnitCodeMulti.Index).Value = .unit_multi
                        rd.Cells(Me.ColUnit.Index).Value = .unit
                        If .Primary_Units <> .Primary_Multi_Units Then
                            If .Is_Kain = False Then
                                rd.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Rasio, .dec_multi) & " " & .unit_multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, .dec) & " " & .unit
                            Else
                                rd.Cells(Me.ColRatioDesc.Index).Value = rd.Cells(Me.ColQtyMulti.Index).Value & " " & .unit_multi & " = " & rd.Cells(Me.ColQty.Index).Value & " " & .unit
                            End If
                        Else
                            rd.Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        rd.Cells(Me.ColSecondaryQty.Index).Value = -1
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = .unit
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        rd.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        rd.Cells(Me.ColLocation.Index).Value = .Storage_Area_Code
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse_Code
                        rd.Cells(Me.ColRemark.Index).Value = .Remarks
                        rd.Cells(Me.ColPrimaryPOno.Index).Value = .Primary_PO_Purchase
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                        rd.Cells(Me.ColPrimaryPdO.Index).Value = .Primary_PdO
                        rd.Cells(Me.ColPrimaryUnitMulti.Index).Value = .Primary_Multi_Units
                        rd.Cells(Me.ColRatioDefault.Index).Value = .Ratio_Default
                        rd.Cells(Me.ColRatioReference.Index).Value = .Rasio
                    End If
                End With
            Next

            If Me.dgvItem.RowCount > 0 Then
                Me.FormattxtSupplierNPacking(False)
            Else
                Me.FormattxtSupplierNPacking(True)
            End If

            'format dgv
            Me.FormatDGV()

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class