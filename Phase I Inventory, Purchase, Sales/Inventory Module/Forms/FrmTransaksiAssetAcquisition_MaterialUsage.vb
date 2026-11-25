Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiAssetAcquisition_MaterialUsage
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
    Public InOutAdjust As clsTransaksiStockInOut.EnumInOutAdjust = clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition

    Dim baru As Boolean
    Dim objStockOut As New clsTransaksiStockInOut
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim objStockOutOutItem As New clsTransaksiStockInOutItem
    Dim ObjAsset As New Core_Module.ClsMasterAsset

    Dim TBL As New Core_Module.dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New Core_Module.dseMasterProduct.stock_location_viewDataTable
    Dim TBL3 As New Core_Module.dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item1DataTable
    Dim TBL5 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item2DataTable
    Dim TBL6 As New dseTransaksiStockInOut.tbltransaksi_stockinout_item3DataTable

    Dim buffer As String = ""
    Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
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
            Me.txtSecondaryQty.Text = "0"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.lblSecondaryUnits.Visible = False
            Me.lblUnits.Left = Me.lblSecondaryUnits.Left
            Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
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
            If Me.objStockOut.GetProductGood(Me.TxtPrimaryCodeNo.Text) <> clsMasterProduct.EnumTipeProduct.Goods Then
                MsgBox("Product code isn't good type, you can't input this product.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
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
            '10. Bila transaksi berkaitan dengan sales dan terinstall software sales maka ada batasan 
            '    berkaitan dengan jumlah item DO
            '11. Bila transaksi berkaitan dengan purchase dan terinstall software purchase maka ada batasan
            '    berkaitan dengan jumlah item DO
            'kalo sampe sini return true
            '12.hanya boleh input type good, asset

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal BatchNoControlled As Boolean, ByVal LocationControlled As Boolean, ByVal UseSecondaryUnit As Boolean) As Double
        Try
            Dim row As DataGridViewRow
            Dim Jml As Double = 0
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
                        Jml = Jml + CDbl(row.Cells(5).Value)
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
            'cek data yang harus diisi
            KoleksiNamaData.Add("stock out number")
            KoleksiTextBox.Add(Me.TxtStockInNo)
            If Me.InOutAdjust <> clsTransaksiStockInOut.EnumInOutAdjust.NoTipe Then
                KoleksiNamaData.Add("Asset Code")
                KoleksiTextBox.Add(Me.TxtCodeAsset)
            End If
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

            If Me.InOutAdjust = clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                If Me.ObjAsset.GetStatus(Me.TxtPrimaryAsset.Text) <> ClsMasterAsset.EnumStatusAsset.On_Progress Then
                    MsgBox("Asset status isn't on progress, you can't save this data." _
                    , MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.InOutAdjust = clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                If Me.ObjAsset.GetStatus(Me.TxtPrimaryAsset.Text) <> ClsMasterAsset.EnumStatusAsset.Active Then
                    MsgBox("Asset status isn't active, you can't save this data." _
                    , MsgBoxStyle.Information, "Information")
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
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objStockOut
                .NoTransaksi = Me.TxtStockInNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Remarks = Me.TxtRemark.Text
                .PrimaryAsset = Me.TxtPrimaryAsset.Text
                .InOutAdjust = Me.InOutAdjust
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
                Me.TxtPrimaryAsset.Text = .PrimaryAsset
                Me.TxtCodeAsset.Text = Me.ObjAsset.GetCode(.PrimaryAsset)
                Me.TxtAssetDesc.Text = Me.ObjAsset.GetDesc(.PrimaryAsset)
                Me.InOutAdjust = .InOutAdjust
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
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
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
                        If Me.objStockOut.SimpanBaru_Material_Usage(Me.dgvItem, Me.UserCode) Then
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
                        If Me.objStockOut.SimpanPerubahan_MaterialUsage(Me.dgvItem, Me.UserCode) Then
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
                        Me.objStockOut.CariDetail_MaterialUsage(Me.TxtPrimaryStockIn.Text, Me.dgvItem)
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
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.TxtCodeAsset.Focus()
            Me.objStockOut.ResetObject(Me.dgvItem)
            Me.objStockOut.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
            Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value, Me.InOutAdjust)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTransaksiAssetAcquisition_MaterialUsage_cari
        Try
            hj.JenisTransaksi = Me.TipeTransaksi
            hj.InOutAdjust = Me.InOutAdjust
            hj.BackColorAll = Me.BackColorAll
            If Me.InOutAdjust = clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                hj.Text = "Find Acquistion Material Usage"
            ElseIf Me.InOutAdjust = clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                hj.Text = "Find Revaluation Material Usage"
            End If

            hj.ShowDialog()
            If Not hj.PrimaryTransaksiStockInOut = -1 Then
                If Me.objStockOut.CariDetail_MaterialUsage(hj.PrimaryTransaksiStockInOut, Me.dgvItem) = True Then
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
                If Me.objStockOut.ApakahBolehEdit(Me.UserCode, True, True) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
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
                If Me.objStockOut.ApakahBolehHapus(Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockOut.Hapus(Me.TxtPrimaryStockIn.Text, Me.UserCode, TblItemCustom) = True Then
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
                            If Me.objStockOut.SimpanBaru_Material_Usage(Me.dgvItem, Me.UserCode) = True Then
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
                            If Me.objStockOut.SimpanPerubahan_MaterialUsage(Me.dgvItem, Me.UserCode) = True Then
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
                Me.lblUnits.Left = Me.lblSecondaryUnits.Left
                Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
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
        Dim i As New Core_Module.frmCariMasterProduct
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
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If

            If Me.objStockOut.GetProductGood(Me.TxtPrimaryCodeNo.Text) <> clsMasterProduct.EnumTipeProduct.Goods Then
                MsgBox("Product code isn't good type, you can't input this product.", MsgBoxStyle.Information, _
                "Information")
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

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New Core_Module.frmCariMasterProduct

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

                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If

                If Me.objStockOut.GetProductGood(Me.TxtPrimaryCodeNo.Text) <> clsMasterProduct.EnumTipeProduct.Goods Then
                    MsgBox("Product code isn't good type, you can't input this product.", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtCodeNo.Clear()
                    Me.TxtPrimaryCodeNo.Text = "-1"
                    Me.txtPrimaryCategory.Text = "-1"
                    Me.KosongkanItem()
                    Me.TxtCodeNo.Focus()
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

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New Core_Module.frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
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
                    Me.TxtQtyperBatch.Focus()
                End If
                If Me.objStockOut.GetProductGood(Me.TxtPrimaryCodeNo.Text) <> clsMasterProduct.EnumTipeProduct.Goods Then
                    MsgBox("Product code isn't good type, you can't input this product.", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtCodeNo.Clear()
                    Me.TxtPrimaryCodeNo.Text = "-1"
                    Me.txtPrimaryCategory.Text = "-1"
                    Me.KosongkanItem()
                    Me.TxtCodeNo.Focus()
                End If
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
        Dim RowDGV As DataGridViewRow

        Try
            If Me.ApakahDataItemBenar = True Then
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQty.Index).Value = CDbl(row.Cells(Me.ColQty.Index).Value) + CDbl(Me.TxtQtyperBatch.Text)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItem.Rows.Add()
                    RowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                    RowDGV.Cells(Me.ColNo.Index).Value = ""
                    RowDGV.Cells(Me.ColPONoDONo.Index).Value = ""
                    RowDGV.Cells(Me.ColSuppDONo.Index).Value = ""
                    RowDGV.Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                    RowDGV.Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                    RowDGV.Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatch.Text
                    RowDGV.Cells(Me.ColUnit.Index).Value = Me.lblUnits.Text
                    RowDGV.Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                    RowDGV.Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                    RowDGV.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text
                    RowDGV.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                    RowDGV.Cells(Me.ColLocation.Index).Value = Me.cboLocation.Text
                    RowDGV.Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                    RowDGV.Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                    RowDGV.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = -1
                    RowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                    RowDGV.Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                    RowDGV.Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnits.Text
                    RowDGV.Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                    RowDGV.Cells(Me.ColPrimaryPdO.Index).Value = Me.txtPrimaryPDO.Text

                    'Me.dgvItem.Rows.Add("", "", "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatch.Text, _
                    'Me.lblUnits.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.cboBatchNo.Text, _
                    'ExpiredDateStr, Me.cboLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    '-1, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnits.Text, _
                    'Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                End If
                'kosongkan gbo item
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
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

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
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
            '''2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '''   tidak perlu ditampilkan
            ''If TidakAdaSecondary = True Then
            ''    Me.dgvItem.Columns(7).Visible = False
            ''    Me.dgvItem.Columns(8).Visible = False
            ''Else
            ''    Me.dgvItem.Columns(7).Visible = True
            ''    Me.dgvItem.Columns(8).Visible = True
            ''End If
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
                        ElseIf e.ColumnIndex = 7 Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 9 Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 11 Then
                            Me.EndEditLocation(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
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
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                JmlDGV = .Cells(5).Value
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
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
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Double, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard
        Dim ObjAsset As New Core_Module.ClsMasterAsset


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
            'cek status asset
            If ObjAsset.GetStatus(Me.TxtPrimaryAsset.Text) <> ClsMasterAsset.EnumStatusAsset.On_Progress Then
                MsgBox("Asset status isn't on progress, you can't edit thie data.", _
                MsgBoxStyle.Information, "Information")
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

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.objStockOut.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objStockOut.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtStockInNo.Text = Me.objStockOut.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value, Me.InOutAdjust)
                    End If
                Else
                    Me.TxtStockInNo.Text = Me.objStockOut.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value, Me.InOutAdjust)
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

    Private Sub BtnFindAsset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindAsset.Click
        Dim frm As New Core_Module.FrmMasterAssetCari
        Try
            frm.txtCari.Text = Me.TxtCodeAsset.Text
            frm.BackColorAll = Me.BackColorAll
            frm.ShowDialog()
            If frm.PrimaryAsset <> -1 Then
                Me.TxtPrimaryAsset.Text = frm.PrimaryAsset
                Me.TxtCodeAsset.Text = Me.ObjAsset.GetCode(frm.PrimaryAsset)
                Me.TxtAssetDesc.Text = Me.ObjAsset.GetDesc(frm.PrimaryAsset)
                Me.TxtRemark.Focus()
            Else
                Me.TxtPrimaryAsset.Text = -1
                Me.TxtCodeAsset.Text = ""
                Me.TxtAssetDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeAsset_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAsset.GotFocus
        Me.TxtCodeAsset.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCodeAsset_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeAsset.KeyDown
        If e.KeyValue = 13 Then
            If Me.ObjAsset.ApakahSudahAda(Me.TxtCodeAsset.Text) = True Then
                Me.TxtPrimaryAsset.Text = Me.ObjAsset.GetPrimary(Me.TxtCodeAsset.Text)
                Me.TxtAssetDesc.Text = Me.ObjAsset.GetDesc(Me.TxtPrimaryAsset.Text)
                Me.TxtRemark.Focus()
            Else
                Me.BtnFindAsset.PerformClick()
            End If
        End If
    End Sub

    Private Sub TxtCodeAsset_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAsset.Leave
        Try
            If Me.TxtCodeAsset.Text <> "" Then
                If Me.ObjAsset.ApakahSudahAda(Me.TxtCodeAsset.Text) = True Then
                    Me.TxtPrimaryAsset.Text = Me.ObjAsset.GetPrimary(Me.TxtCodeAsset.Text)
                    Me.TxtAssetDesc.Text = Me.ObjAsset.GetDesc(Me.TxtPrimaryAsset.Text)
                    'If Me.ObjAsset.GetStatus(Me.TxtPrimaryAsset.Text) <> Me.ObjAsset.EnumStatusAsset.On_Progress Then
                    '    MsgBox("Asset status isn't on progress, you can't edit thie data.", _
                    '    MsgBoxStyle.Information, "Information")
                    '    Me.TxtPrimaryAsset.Text = -1
                    '    Me.TxtCodeAsset.Text = ""
                    '    Me.TxtAssetDesc.Text = ""
                    'End If
                Else
                    Me.BtnFindAsset.PerformClick()
                End If
            Else
                Me.TxtPrimaryAsset.Text = -1
                Me.TxtAssetDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeAsset_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeAsset.LostFocus
        Me.TxtCodeAsset.BackColor = Me.TxtPrimaryCodeNo.BackColor
    End Sub
End Class