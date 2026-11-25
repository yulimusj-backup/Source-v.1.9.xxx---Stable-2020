Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiStockInAdj
    Public UserCode As String = ""
    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    Public BeginingBalance As Boolean = False
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim objStockIn As New clsTransaksiStockInOut
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjStockInOutItem As New clsTransaksiStockInOutItem
    Dim ObjItemInOut As New clsTransaksiStockInOut
    Dim ObjItemBundling As New clsTransaksiStockBundlingItem
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjAcc As New Core_Module.ClsTranskasiAccClosing
    Dim objItem As New ClsTransaksiMaterialPlanningGroupItem
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjUser As New ClsUser

    Dim ObjStockBundling As New ClsTransaksiStockBundling
    Dim DTAItemAft As New dseTransaksiStockBundlingTableAdapters.transaksi_stockbundling_itemafter_viewTableAdapter
    Dim TBLItemAft As New dseTransaksiStockBundling.transaksi_stockbundling_itemafter_viewDataTable
    Dim RowItemAft As dseTransaksiStockBundling.transaksi_stockbundling_itemafter_viewRow

    Dim buffer As String = ""
    Dim AutoNumber As Boolean = True
    Dim QuickAddByRefDoc As Boolean = False
    Dim QuickAddByBarcode As Boolean = False
    Dim QuickAddByProdLevel As Boolean = False

    Public TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False

            Me.SetGboItem()
            Me.FormatDGV()
            Me.lblQuickInput.Enabled = True
            Me.llbSetAmount.Enabled = True
            Me.lblQuickInput2.Enabled = True
            Me.llbQuickInputBundling.Enabled = True
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockInAdj_QuickInputProductLevel") = True Then
                Me.QuickAddByProdLevel = True
            Else
                Me.QuickAddByProdLevel = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockInAdj_QuickInputDocRef") = True Then
                Me.QuickAddByRefDoc = True
            Else
                Me.QuickAddByRefDoc = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockInAdj_QuickInputBarcode") = True Then
                Me.QuickAddByBarcode = True
            Else
                Me.QuickAddByBarcode = False
            End If
            If Me.QuickAddByBarcode = True Or Me.QuickAddByProdLevel = True Or Me.QuickAddByRefDoc = True Then
                Me.llbAdvance.Enabled = True
            Else
                Me.llbAdvance.Enabled = False
            End If
            Me.AturHakAksesFormCustom()

            Me.objStockIn.IsiCboSettingActive(Me.CboAdjType, True)
            If Me.TxtPrimarySetting.Text <> "" Or Me.TxtPrimarySetting.Text <> "-1" Then
                Me.CboAdjType.ValueMember = Me.objStockIn.GetCodeSettingByPrimary(Me.TxtPrimarySetting.Text)
                If Me.objStockIn.ApakahAdjustmentTypeAktif(Me.TxtPrimarySetting.Text) = True Then
                    Me.CboAdjType.SelectedItem = Me.CboAdjType.ValueMember
                    Me.TxtAdjustmentTypeDesc.Text = Me.objStockIn.GetDescriptionSettingAdjustment(Me.TxtPrimarySetting.Text)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
            Me.lblQuickInput.Enabled = False
            Me.llbSetAmount.Enabled = False
            Me.lblQuickInput2.Enabled = False
            Me.llbQuickInputBundling.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.TxtPrimarySetting.Text = "-1"
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
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.txtSecondaryQty.Visible = False
            Me.lblSecondaryUnits.Visible = False
            Me.lblUnits.Left = Me.lblSecondaryUnits.Left
            Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
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
                If Me.txtBatchNo.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtBatchNo.Focus()
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
                    If row.Cells(9).Value = Me.txtBatchNo.Text And Me.txtBatchNo.Text <> "" And row.Cells(15).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(10).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.txtBatchNo.Text + " on " + _
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
        Dim MaxDate As Date = Today
        Dim TglConvSMESS As Date = #1/1/1950#

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ApakahTransferOtomatis(Me.dtpDate.Value) = False And Me.BeginingBalance = False Then
                'tidak boleh bikin stock in out, pada tanggal sebelum convertion date -> bukan begining balance
                MsgBox("Please specify date after convertion date", MsgBoxStyle.Information, "Information")
                Return False
            ElseIf Me.BeginingBalance = True Then
                'tidak boleh bikin stock in out, pada tanggal setelah convertion date -> begining balance
                TglConvSMESS = Me.ObjAcc.GetSMESSConvertionDate
                If Me.dtpDate.Value.Date >= TglConvSMESS.Date Then
                    MsgBox("Minimum transaction date is on SMESS convertion date: " + Format(TglConvSMESS, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                    Me.dtpDate.Focus()
                    Return False
                End If
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

            'kalo begining balance, maka nilai amountnya harus positive
            If Me.BeginingBalance = True Then
                For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                    If rowDGV.Cells(Me.ColAmount.Index).Value <= 0 Then
                        MsgBox("Please specify positive amount for: " + rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString, _
                        MsgBoxStyle.Information, "Information")

                        Me.dgvItem.ClearSelection()
                        rowDGV.Selected = True
                        Return False
                    End If
                Next

                'tanggalnya harus sebelum conversion date
                MaxDate = Me.objStockIn.GetAccountingConvertionDate.Date
                If MaxDate < Me.dtpDate.Value.Date Then
                    MsgBox("Maximum allowable date is before: " + Format(MaxDate - TimeSpan.FromDays(1), "dd MMM yyyy"), _
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
            'If objLocationAccess.HasAccessPrivileges(Me.UserCode, listOfPrimaryLocation, FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut.STOCK_IN_ADJUST) = False Then
            '    Return False
            'End If

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_IN_ADJUST, Me.TBLDetailLocation) = False Then
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
            With Me.objStockIn
                .NoTransaksi = Me.TxtStockInNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Remarks = Me.TxtRemark.Text
                .PrimaryAdjType = Me.TxtPrimarySetting.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim AdjCode As String = ""

        Try
            With Me.objStockIn
                Me.objStockIn.IsiCboSettingActive(Me.CboAdjType, False) 'update ahdar
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.TxtPrimarySetting.Text = .PrimaryAdjType
                AdjCode = Me.objStockIn.GetCodeSettingByPrimary(.PrimaryAdjType)
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

    Private Sub FrmTransaksiStockInAdj_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try 'update ahdar

            If Me.TxtStockInNo.Text <> "" Then
                If Me.TxtRemark.Enabled = False Then
                    Me.objStockIn.IsiCboSettingActive(Me.CboAdjType, False)
                    If Me.TxtPrimarySetting.Text <> "" Or Me.TxtPrimarySetting.Text <> "-1" Then
                        Me.CboAdjType.ValueMember = Me.objStockIn.GetCodeSettingByPrimary(Me.TxtPrimarySetting.Text)
                        Me.CboAdjType.SelectedItem = Me.CboAdjType.ValueMember
                        Me.TxtAdjustmentTypeDesc.Text = Me.objStockIn.GetDescriptionSettingAdjustment(Me.TxtPrimarySetting.Text)
                    End If

                Else
                    Me.objStockIn.IsiCboSettingActive(Me.CboAdjType, True)
                    If Me.TxtPrimarySetting.Text <> "" Or Me.TxtPrimarySetting.Text <> "-1" Then
                        Me.CboAdjType.ValueMember = Me.objStockIn.GetCodeSettingByPrimary(Me.TxtPrimarySetting.Text)
                        If Me.objStockIn.ApakahAdjustmentTypeAktif(Me.TxtPrimarySetting.Text) = True Then
                            Me.CboAdjType.SelectedItem = Me.CboAdjType.ValueMember
                            Me.TxtAdjustmentTypeDesc.Text = Me.objStockIn.GetDescriptionSettingAdjustment(Me.TxtPrimarySetting.Text)
                        End If
                    End If
                End If
            End If

            If Me.CboAdjType.SelectedIndex = -1 Or Me.CboAdjType.SelectedIndex = Nothing Then
                Me.TxtAdjustmentTypeDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ShowIcon = False
            Me.ObjStockBundling.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Purchase1
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
            If Me.AutoNumber = True Then
                Me.TxtStockInNo.ReadOnly = True
            Else
                Me.TxtStockInNo.ReadOnly = False
            End If

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.KosongkanField()
            Me.objStockIn.IsiCboSetting(Me.CboAdjType)

            If Me.UserCode = "USER" Then
                Me.lblQuickInput.Visible = True
                'If Me.BeginingBalance = True Then
                '    Me.llbSetAmount.Visible = True
                'Else
                '    Me.llbSetAmount.Visible = False
                'End If
            Else
                Me.lblQuickInput.Visible = False
                Me.llbSetAmount.Visible = False
            End If
            Me.LlbUpdDecProd.Visible = False
            Me.lblQuickInput2.Visible = False

            Me.DTAItemAft.Connection.ConnectionString = Core_Module.GetConstring
            Me.llbQuickInputBundling.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            Me.CboAdjType.Anchor = Me.TxtStockInNo.Anchor
            Me.TxtAdjustmentTypeDesc.Anchor = Me.TxtRemark.Anchor
            Me.TxtAdjustmentTypeDesc.Left += 50
            Me.TxtAdjustmentTypeDesc.Width -= 50
            Me.CboAdjType.Width += 50


            Me.AturHakAksesFormCustom()
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
                        If Me.objStockIn.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
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
                        If Me.objStockIn.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
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
            Me.TxtCodeNo.Text = "" And Me.txtBatchNo.Text = "" And Me.TxtLocation.Text = "" And _
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
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
            If Me.AutoNumber = True Then
                Me.TxtStockInNo.Text = Me.objStockIn.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
            End If
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        If Me.BeginingBalance = True Then
            Me.FindBeginingBalance()
            Exit Sub
        End If

        Dim hj As New FrmCariTransaksiStockInOut
        Try
            hj.JenisTransaksi = Me.TipeTransaksi
            hj.ShowDialog()
            If Not hj.PrimaryTransaksiStockInOut = -1 Then
                If Me.objStockIn.CariDetail(hj.PrimaryTransaksiStockInOut, Me.dgvItem, Me.TblItemCustom) = True Then
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

    Private Sub FindBeginingBalance()
        Dim hj As New FrmCariTransaksiStockInOut2

        Try
            hj.JenisTransaksi = Me.TipeTransaksi
            hj.ShowDialog()
            If Not hj.PrimaryTransaksiStockInOut = -1 Then
                If Me.objStockIn.CariDetail(hj.PrimaryTransaksiStockInOut, Me.dgvItem, Me.TblItemCustom) = True Then
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
                'update baru by: 2024-2-29 'update ahdar
                If Me.objStockIn.ApakahAdjustmentTypeAktif(Me.TxtPrimarySetting.Text) = False Then
                    MsgBox("Please activate the adjustment type first so you can edit this transaction.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.objStockIn.ApakahBolehEdit(Me.UserCode) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
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
                'update baru by: 2024-2-29
                If Me.objStockIn.ApakahAdjustmentTypeAktif(Me.TxtPrimarySetting.Text) = False Then 'update ahdar
                    MsgBox("Please activate the adjustment type first so you can delete this transaction.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'tidak cek stock minus ketika cek apakah boleh hapus, karena akan di cek tepat sebelum di hapus,
                'pengecekan ada dalam fungsi hapus di kelas, sehingga memperhitungkan antrian
                If Me.objStockIn.ApakahBolehHapus(Me.UserCode, True, True, False, False) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
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

    Private Sub SetGboItem()
        Try
            If IsNumeric(Me.TxtPrimaryCodeNo.Text) = False Then
                Me.TxtPrimaryCodeNo.Text = "-1"
            End If
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
            Me.chkPdoControlled, Me.txtPrimaryUnits, Me.txtPrimarySecondaryUnits)

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
            End If

            If Me.chkUseSecondaryUnit.Checked = False Then
                Me.txtSecondaryQty.Visible = False
                Me.lblSecondaryUnits.Visible = False
                Me.lblUnits.Left = Me.lblSecondaryUnits.Left
                Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
            Else
                Me.txtSecondaryQty.Visible = True
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                Me.lblUnits.Left = Me.lblSecQty.Left - 37
                Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
                Me.txtSecondaryQty.Text = "0"
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
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        Me.TxtQtyperBatch.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
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
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
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
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        Me.TxtQtyperBatch.Focus()
                    End If
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
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
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
                If Me.txtBatchNo.Visible = True Then
                    Me.txtBatchNo.Focus()
                ElseIf Me.TxtLocation.Visible = True Then
                    Me.TxtLocation.Focus()
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
        Dim PrimaryLocationLama As Double = -1

        Try
            If Me.ApakahDataItemBenar = True Then
                PrimaryLocationLama = Me.txtPrimaryLocation.Text

                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(5).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(5).Value) + CDbl(Me.TxtQtyperBatch.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItem.Rows.Add(Me.dgvItem.Rows.Count + 1, "", "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatch.Text, _
                    Me.lblUnits.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.txtBatchNo.Text, _
                    ExpiredDateStr, Me.TxtLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    -1, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnits.Text, _
                    Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text, 0, 0)
                    'set ulang nomor
                    'Me.ObjCD.SetNumber(Me.dgvItem)
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

    Public Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
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
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
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
            '4. Expiry Date
            '5. Location Code
            '6. Remarks
            '7. Amount
            '8. Unit Cost
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 Or e.ColumnIndex = 11 _
                Or e.ColumnIndex = 13 Or e.ColumnIndex = 10 Or e.ColumnIndex = Me.ColAmount.Index Or _
                e.ColumnIndex = Me.ColUnitCost.Index Then
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
                    ElseIf e.ColumnIndex = Me.ColAmount.Index Or _
                    e.ColumnIndex = Me.ColQty.Index Or _
                    e.ColumnIndex = Me.ColUnitCost.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
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
            Dim DefPrec1 As Integer = 0

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> 10 Then
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
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
                        ElseIf e.ColumnIndex = 10 Then
                            Me.EndEditExpiryDate(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColAmount.Index Then
                            'pastikan positive 
                            If CDec(Me.buffer) = CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            ElseIf Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please specify positive number for amount value.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If

                            'digit grouping
                            DefPrec1 = Me.ObjCur.GetDecimal_DefaultCurrency
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                            DefPrec1)
                            If CDec(Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value) <> 0 Then
                                Me.dgvItem.Item(Me.ColUnitCost.Index, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value) / _
                                CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value), DefPrec1)
                            Else
                                Me.dgvItem.Item(Me.ColUnitCost.Index, e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(0, DefPrec1)
                            End If
                        ElseIf e.ColumnIndex = Me.ColUnitCost.Index Then
                            'pastikan positive 
                            If CDec(Me.buffer) = CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            ElseIf Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please specify positive number for unit cost value.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If

                            'digit grouping
                            DefPrec1 = Me.ObjCur.GetDecimal_DefaultCurrency
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                            DefPrec1)
                            Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) * _
                            CDec(Me.dgvItem.Item(Me.ColUnitCost.Index, e.RowIndex).Value), DefPrec1)
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

    Private Sub EndEditExpiryDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item(15, e.RowIndex).Value, _
                Me.dgvItem.Item(9, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(9, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'bandingkan dengan yang sudah ada didalam dgv
                Dim row1 As DataGridViewRow
                For Each row1 In Me.dgvItem.Rows
                    If row1.Cells(9).Value = row.Cells(9).Value And CStr(row.Cells(9).Value) <> "" _
                    And row1.Index <> row.Index And row1.Cells(15).Value = row.Cells(15).Value Then
                        If CDate(row1.Cells(10).Value).Date <> CDate(row.Cells(10).Value).Date Then
                            MsgBox("You already entered expiry date for batch no: " + CStr(row1.Cells(9).Value) + " on " + _
                            Format(CDate(row1.Cells(10).Value), "dd MMM yyyy"), MsgBoxStyle.Information, _
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
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItem.Item(15, e.RowIndex).Value, _
                    Me.dgvItem.Item(9, e.RowIndex).Value, JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> CDate(Me.dgvItem(10, e.RowIndex).Value) Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(9, e.RowIndex).Value) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItem.Item(15, e.RowIndex).Value, _
                Me.dgvItem.Item(9, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItem(10, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(9, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")

                'expiry date, tidak perlu cek kemungkinan merger, karena tidak mungkin
                'hanya expiry date saja yang berbeda sedangkan batch no-nya sama
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                .Cells(9).Value, CDate(Me.buffer), SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim DefPrecCur As Integer = 0

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
                If Me.ObjProduct.ApakahGoods(.Cells(15).Value) = True Then
                    JmlDGV = .Cells(5).Value
                    If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                    .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    DefPrecCur = Me.ObjCur.GetDecimal_DefaultCurrency
                    Me.dgvItem.Item(Me.ColAmount.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) * _
                    CDec(Me.dgvItem.Item(Me.ColUnitCost.Index, e.RowIndex).Value), DefPrecCur)
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
    ByVal SecondaryQty As Double, ByVal JmlDGV As Double, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                    End If
                End If
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper.Trim
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                Me.buffer, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItem.Item(10, e.RowIndex).Value) = True Then
                        '0. basis data
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
                            ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItem.Item(15, e.RowIndex).Value, _
                            Me.dgvItem.Item(9, e.RowIndex).Value, JenisTransaksi)
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
                        Next
                        '2. transaksi stock bundling untuk item hasil bundling
                        ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItem.Item(15, e.RowIndex).Value, _
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
                        '3. proses stock taking
                        'msgbox("belum")

                        '4. data grid view
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
                            Me.dgvItem.Item(10, e.RowIndex).Value = Format(CDate(BufferExpiryDate), "dd MMM yyyy")
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
                If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'update data primary location
                    Dim Obj1 As New clsMasterStorageArea
                    .Cells(16).Value = Obj1.GetPrimary(.Cells(11).Value)
                    'update data gudang
                    .Cells(12).Value = Obj1.GetWarehouseCode(.Cells(16).Value)

                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'update primary location
                            .Cells(16).Value = Obj1.GetPrimary(.Cells(11).Value)
                            'update data gudang
                            .Cells(12).Value = Obj1.GetWarehouseCode(.Cells(16).Value)
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
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Double, Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal As Decimal '2026-03-23: ganti ke decimal. kalo double ngaco
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim decimalPoint As Integer = 0

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.ObjStockInOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties1(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary, decimalPoint)
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Minimum value you must can enter is " + ObjCD.DigitGrouping(JmlObjLama - JmlGudang, decimalPoint) + _
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

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.objStockIn.CariDetail(vPrimary, Me.dgvItem, Me.TblItemCustom) = True Then
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
        'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
        Dim PrimaryLocation, PrimaryPdO, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
        JmlDGV, JmlSetelahHapus As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim row As DataGridViewRow = e.Row
        Try
            'hanya cek batasa ini untuk goods.
            If Me.ObjProduct.ApakahGoods(row.Cells(15).Value) = True Then
                With row
                    PrimaryLocation = .Cells(16).Value
                    PrimaryPdO = .Cells(19).Value
                    PrimaryProduct = .Cells(15).Value
                    BatchNo = .Cells(9).Value
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
                    JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                    'karena akan dihapus, perhitungan menggunakan nilai jumlah dgv =0
                    JmlDGV = 0
                    JmlObjLama = Me.ObjStockInOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
                    SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                    JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
                    If JmlSetelahHapus < 0 Then
                        MsgBox("You cannot delete this data, because this will make you stock quantity minus.", _
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

    Private Sub txtBatchNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.GotFocus
        Try
            Me.txtBatchNo.BackColor = Me.BackColorAll
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
            Me.TxtLocation.BackColor = Me.BackColorAll
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

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.objStockIn.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objStockIn.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtStockInNo.Text = Me.objStockIn.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        If Me.AutoNumber = True Then
                            Me.TxtStockInNo.Text = Me.objStockIn.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
                        End If
                    End If
                Else
                    If Me.AutoNumber = True Then
                        Me.TxtStockInNo.Text = Me.objStockIn.GetNoBaru(Me.TipeTransaksi, Me.dtpDate.Value)
                    End If
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

    Private Sub txtBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.Leave
        Try
            Me.txtBatchNo.Text = Me.txtBatchNo.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInAdj_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            If Me.BeginingBalance = True Then
                Me.ColAmount.Visible = True
                Me.Text = "Begining Balance"
                Me.LblNo.Text = "Document No"
                Me.ColUnitCost.Visible = True
            Else
                Me.ColAmount.Visible = False
                Me.ColUnitCost.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtRemark.TextChanged

    End Sub

    Private Sub CboAdjType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboAdjType.SelectedValueChanged
        Try
            Me.TxtPrimarySetting.Text = Me.objStockIn.GetPrimarySettingByCode(Me.CboAdjType.Text)
            Me.TxtAdjustmentTypeDesc.Text = Me.objStockIn.GetDescriptionSettingAdjustment(Me.TxtPrimarySetting.Text)
            If Me.CboAdjType.SelectedIndex = -1 Then
                Me.TxtAdjustmentTypeDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahSudahAda(ByVal vPrimaryProduct As Double) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblQuickInput.LinkClicked
        Try
            Me.QuickInput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub QuickInput()
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tblmaster_product1TableAdapter
        Dim TBL As New dseTransaksiStockInOut.tblmaster_product1DataTable
        Dim row As dseTransaksiStockInOut.tblmaster_product1Row
        Dim Tipe As clsMasterProduct.EnumTipeProduct = clsMasterProduct.EnumTipeProduct.Goods
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim BatchRef As String = "AWAL"
        Dim PrimaryLocationRef As Double = -1
        Dim ExpiryRef As Date = #1/1/2500#
        Dim QtyRef As Decimal = 1000000
        Dim ij As New FrmTransaksiStockInAdj_QuickInput

        Dim BatchNumber As String = ""
        Dim PrimaryLocation As Double = -1
        Dim Expiry As Date = #1/1/1950#
        Dim WarehouseCode, WarehouseCodeRef, LocationCode, LocationCodeRef As String

        Dim DTALocation As New dseTransaksiStockInOutTableAdapters.tblmaster_storage_areaTableAdapter
        Dim TBLLocation As New dseTransaksiStockInOut.tblmaster_storage_areaDataTable
        Dim rowLocation As dseTransaksiStockInOut.tblmaster_storage_areaRow
        Dim ListProductTidakDiInput As New List(Of String)
        Dim obj As New clsMasterStorageArea
        Dim Level As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If Me.dgvItem.Rows.Count > 0 Then
                MsgBox("Please make sure dgv is empty.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If MsgBox("Start quick input?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            ij.DefQty = QtyRef
            If ij.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Exit Sub
            End If

            '2025 IM, tambah pengecekan count, agar membedakan
            ij.dgvItemProductGroup.AllowUserToAddRows = False
            If ij.dgvItemProductGroup.Rows.Count > 0 Then
                If ij.chkImportFromDB.Checked = True Then
                    Me.ImportQtyFromDB(ij, DTA, TBL)
                ElseIf ij.ChkMirror.Checked = True Then
                    If MsgBox("We will mirror stock data as off " + Format(Me.dtpDate.Value, "dd MMM yyyy") + ". Continue process?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ImportMirror(ij)
                    End If
                ElseIf ij.chkCoverMinus.Checked = True Then
                    Me.CoverMinus(ij.dtpCoverMinus_RefDate.Value.Date - TimeSpan.FromDays(1))
                Else
                    QtyRef = CDec(ij.TxtQty.Text)

                    Windows.Forms.Cursor.Current = Cursors.WaitCursor

                    DTA.Connection.ConnectionString = Core_Module.GetConstring
                    DTALocation.Connection.ConnectionString = DTA.Connection.ConnectionString

                    ij.dgvItemProductGroup.AllowUserToAddRows = False
                    DTALocation.Fill(TBLLocation)

                    'get kode lokasi
                    If TBLLocation.Rows.Count > 0 Then
                        'rowLocation = TBLLocation.Rows(0)
                        'PrimaryLocationRef = rowLocation.Primary
                        'WarehouseCodeRef = obj.GetWarehouseCode(PrimaryLocationRef)
                        'LocationCodeRef = obj.GetStorageAreaCode(PrimaryLocationRef)
                        Dim ijk As New Core_Module.FrmCariMasterStorageArea
                        ijk.ShowDialog()
                        If ijk.PrimaryStorageArea <> -1 Then
                            PrimaryLocationRef = ijk.PrimaryStorageArea
                            WarehouseCodeRef = obj.GetWarehouseCode(PrimaryLocationRef)
                            LocationCodeRef = obj.GetStorageAreaCode(PrimaryLocationRef)
                        Else
                            PrimaryLocationRef = -1
                            WarehouseCodeRef = ""
                            LocationCodeRef = ""
                        End If
                    Else
                        PrimaryLocationRef = -1
                        WarehouseCodeRef = ""
                        LocationCodeRef = ""
                    End If

                    For Each rowDGV1 In ij.dgvItemProductGroup.Rows
                        TBL.Rows.Clear()
                        Level = Me.objItem.GetProductLevel(rowDGV1.Cells(ij.ColProductGroup_Level.Index).Value)
                        If Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                            DTA.FillBy_PrimaryProductGroup(TBL, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                        ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                            DTA.FillBy_PrimaryBrand(TBL, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                        ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                            DTA.FillBy_PrimarySubBrand(TBL, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                        ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                            DTA.FillBy_PrimaryCategory(TBL, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                        Else
                            DTA.FillBy_PrimaryCategory(TBL, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                        End If

                        For Each row In TBL.Rows
                            If Me.ApakahSudahAda(row.Primary) = False Then
                                If PrimaryLocationRef = -1 And row.Location_Controlled = 1 Then
                                    'kalo sampe ada product yang kontrol lokasi tetapi tidak ada master lokasi
                                    ListProductTidakDiInput.Add(row.Product_Code)
                                Else
                                    If row.Batch_No_Controlled = 0 Then
                                        BatchNumber = ""
                                        Expiry = #1/1/1950#
                                    Else
                                        BatchNumber = BatchRef
                                        If row.Expiry_Date_Controlled = True Then
                                            Expiry = ExpiryRef
                                        Else
                                            Expiry = #1/1/1950#
                                        End If
                                    End If
                                    If row.Location_Controlled = 1 Then
                                        PrimaryLocation = PrimaryLocationRef
                                        LocationCode = LocationCodeRef
                                        WarehouseCode = WarehouseCodeRef
                                    Else
                                        PrimaryLocation = -1
                                        LocationCode = ""
                                        WarehouseCode = ""
                                    End If

                                    Me.dgvItem.Rows.Add()
                                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                                    With rowDGV
                                        .Cells(Me.ColAmount.Index).Value = 0
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
                                        .Cells(Me.ColPONo.Index).Value = ""
                                        .Cells(Me.ColPrimaryCodeNo.Index).Value = row.Primary
                                        .Cells(Me.ColPrimaryLocation.Index).Value = PrimaryLocation
                                        .Cells(Me.ColPrimaryPdO.Index).Value = -1
                                        .Cells(Me.ColPrimaryPOno.Index).Value = -1
                                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                                        .Cells(Me.ColPrimaryUnits.Index).Value = row.Primary_Unit
                                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(QtyRef, row._Decimal)
                                        .Cells(Me.ColRemark.Index).Value = ""
                                        .Cells(Me.ColSecondaryQty.Index).Value = -1
                                        .Cells(Me.ColSecondaryUnit.Index).Value = ""
                                        .Cells(Me.ColSuppDONo.Index).Value = ""
                                        .Cells(Me.ColUnit.Index).Value = row.Unit_Code
                                        .Cells(Me.ColWareHouse.Index).Value = WarehouseCode
                                    End With
                                End If
                            End If
                        Next
                    Next
                End If
            ElseIf ij.dgvImport.Rows.Count > 0 Then
                '2025 IM, tambah rowdgv untuk import
                Dim rowDGVImport As DataGridViewRow
                'pindahkan dgv import ke dgv item (karena kolom sudah sama)
                For Each rowDGVImport In ij.dgvImport.Rows
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColAmount.Index).Value = rowDGVImport.Cells(ij.Col_Import_Amount.Index).Value
                        .Cells(Me.ColBatch.Index).Value = rowDGVImport.Cells(ij.Col_Import_Batch.Index).Value
                        .Cells(Me.ColCodeNo.Index).Value = rowDGVImport.Cells(ij.Col_Import_CodeNo.Index).Value
                        .Cells(Me.ColDescription.Index).Value = rowDGVImport.Cells(ij.Col_Import_Description.Index).Value
                        If rowDGVImport.Cells(ij.Col_Import_Description.Index).Value = "-" Or String.IsNullOrEmpty(rowDGVImport.Cells(ij.Col_Import_Description.Index).Value) Then
                            .Cells(Me.ColExpiredDate.Index).Value = "-"
                        Else
                            .Cells(Me.ColExpiredDate.Index).Value = Format(rowDGVImport.Cells(ij.Col_Import_Description.Index).Value, "dd MMM yyyy")
                        End If
                        .Cells(Me.ColLocation.Index).Value = rowDGVImport.Cells(ij.Col_Import_Location.Index).Value
                        .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                        .Cells(Me.ColPONo.Index).Value = rowDGVImport.Cells(ij.Col_Import_PONo.Index).Value
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDGVImport.Cells(ij.Col_Import_PrimaryCodeNo.Index).Value
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowDGVImport.Cells(ij.Col_Import_PrimaryLocation.Index).Value
                        .Cells(Me.ColPrimaryPdO.Index).Value = rowDGVImport.Cells(ij.Col_Import_PrimaryPdO.Index).Value
                        .Cells(Me.ColPrimaryPOno.Index).Value = rowDGVImport.Cells(ij.Col_Import_PrimaryPOno.Index).Value
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = rowDGVImport.Cells(ij.Col_Import_PrimarySecondaryUnits.Index).Value
                        .Cells(Me.ColPrimaryUnits.Index).Value = rowDGVImport.Cells(ij.Col_Import_PrimaryUnits.Index).Value
                        .Cells(Me.ColQty.Index).Value = rowDGVImport.Cells(ij.Col_Import_Qty.Index).Value
                        .Cells(Me.ColRemark.Index).Value = rowDGVImport.Cells(ij.col_Import_Remarks.Index).Value
                        .Cells(Me.ColSecondaryQty.Index).Value = rowDGVImport.Cells(ij.Col_Import_SecondaryQty.Index).Value
                        .Cells(Me.ColSecondaryUnit.Index).Value = rowDGVImport.Cells(ij.Col_Import_SecondaryUnit.Index).Value
                        .Cells(Me.ColSuppDONo.Index).Value = rowDGVImport.Cells(ij.Col_Import_SuppDONo.Index).Value
                        .Cells(Me.ColUnit.Index).Value = rowDGVImport.Cells(ij.Col_Import_Unit.Index).Value
                        .Cells(Me.ColUnitCost.Index).Value = rowDGVImport.Cells(ij.Col_Import_UnitCost.Index).Value
                        .Cells(Me.ColWareHouse.Index).Value = rowDGVImport.Cells(ij.Col_Import_WareHouse.Index).Value
                    End With
                Next
            Else
                Exit Sub
            End If

            FormatDGV()

            If ListProductTidakDiInput.Count > 0 Then
                'kabari ke user soal masalah product2 yang location control tetapi tidak ada master lokasinya
                MsgBox("")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTALocation.Dispose()
            TBLLocation.Dispose()
            ListProductTidakDiInput = Nothing
            obj = Nothing
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            ij = Nothing
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub CoverMinus(ByVal vRefDate As Date)
        Dim ij As New FrmTransaksiStockInAdj_CoverMinus
        Dim TBLQty As New dseTransaksiStockInOut_CoverMinus.tblmaster_productqtyDataTable
        Dim DTAProd As New dseTransaksiStockInOut_CoverMinusTableAdapters.tblmaster_productTableAdapter
        Dim TBLProd As New dseTransaksiStockInOut_CoverMinus.tblmaster_productDataTable
        Dim rowProd As dseTransaksiStockInOut_CoverMinus.tblmaster_productRow
        Dim rowQty As dseTransaksiStockInOut_CoverMinus.tblmaster_productqtyRow
        Dim ExpiredDateStr As String = ""
        Dim DTALocation As New dseTransaksiStockInOut_CoverMinusTableAdapters.tblmaster_locationTableAdapter
        Dim TBLLocation As New dseTransaksiStockInOut_CoverMinus.tblmaster_locationDataTable
        Dim rowLocation As dseTransaksiStockInOut_CoverMinus.tblmaster_locationRow

        Try
            DTAProd.Connection.ConnectionString = Core_Module.GetConstring
            DTALocation.Connection.ConnectionString = DTAProd.Connection.ConnectionString

            'berikan msg peringatan bahwa harus user yang berdiri sendiri
            MsgBox("Make sure there is no other user online.", MsgBoxStyle.Information, "Information")
            'kosongkan dgv
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()
            ij.TblQty = TBLQty
            ij.RefDate = vRefDate
            ij.TblQty.DefaultView.RowFilter = ""
            ij.TblQty.DefaultView.Sort = "[Primary Product]"
            DTAProd.Fill(TBLProd)
            rowProd = TBLProd.Rows(0)
            DTALocation.Fill(TBLLocation)
            rowLocation = TBLLocation.Rows(0)

            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                ij.TblQty.DefaultView.RowFilter = ""
                ij.TblQty.DefaultView.Sort = "[Primary Product]"

                'masukkan tabel qty ke dalam dgv
                For Each DV As Data.DataRowView In ij.TblQty.DefaultView
                    rowQty = DV.Row
                    If rowProd.Primary <> rowQty.Primary_Product Then
                        TBLProd.DefaultView.RowFilter = "Primary = " + rowQty.Primary_Product.ToString
                        rowProd = TBLProd.DefaultView(0).Row
                    End If
                    If rowLocation.Primary <> rowQty.Primary_Storage_Area Then
                        If rowQty.Primary_Storage_Area <> -1 Then
                            TBLLocation.DefaultView.RowFilter = "Primary = " + rowQty.Primary_Storage_Area.ToString
                            rowLocation = TBLLocation.DefaultView(0).Row
                        Else
                            rowLocation = TBLLocation.Rows(0)
                        End If
                    Else
                        rowLocation = TBLLocation.Rows(0)
                    End If

                    If rowQty.Expiry_Date.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If

                    If rowQty.Primary_Storage_Area <> -1 Then
                        Me.dgvItem.Rows.Add("", "", "", rowProd.Product_Code, rowProd.Description, Me.ObjCD.DigitGrouping(Math.Abs(rowQty.Qty), rowProd._Decimal), _
                                            rowProd.Unit_Code, -1, "", rowQty.Lot_No, ExpiredDateStr, rowLocation.Storage_Area_Code, rowLocation.Warehouse_Code, "", _
                                            -1, rowQty.Primary_Product, rowQty.Primary_Storage_Area, rowProd.Primary_Unit, _
                                            -1, -1, 0)
                    Else
                        Me.dgvItem.Rows.Add("", "", "", rowProd.Product_Code, rowProd.Description, Me.ObjCD.DigitGrouping(Math.Abs(rowQty.Qty), rowProd._Decimal), _
                                            rowProd.Unit_Code, -1, "", rowQty.Lot_No, ExpiredDateStr, "", "", "", _
                                            -1, rowQty.Primary_Product, rowQty.Primary_Storage_Area, rowProd.Primary_Unit, _
                                            -1, -1, 0)
                    End If
                Next

                'set ulang nomor
                Me.ObjCD.SetNumber(Me.dgvItem)
                Me.FormatDGV()
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
            TBLQty.Dispose()
            DTAProd.Dispose()
            TBLProd.Dispose()
            DTALocation.Dispose()
            TBLLocation.Dispose()
        End Try
    End Sub

    Private Sub ImportQtyFromDB(ByVal ij As FrmTransaksiStockInAdj_QuickInput, ByVal DTA As dseTransaksiStockInOutTableAdapters.tblmaster_product1TableAdapter, _
   ByVal TBLDummy As dseTransaksiStockInOut.tblmaster_product1DataTable)
        Dim Level As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim rowDGV1 As DataGridViewRow
        Dim Tipe As clsMasterProduct.EnumTipeProduct = clsMasterProduct.EnumTipeProduct.Goods
        Dim DTAImport As New dseTransaksiStockInOutTableAdapters.Import_StockTableAdapter
        Dim TBLImport As New dseTransaksiStockInOut.Import_StockDataTable
        Dim rowImport As dseTransaksiStockInOut.Import_StockRow
        Dim TBL As New dseTransaksiStockInOut.tblmaster_product1DataTable
        Dim row, rowDummy As dseTransaksiStockInOut.tblmaster_product1Row
        Dim rowDGV As DataGridViewRow
        Dim DTALocation As New dseTransaksiStockInOutTableAdapters.TBLMasterLocationTableAdapter
        Dim TBLLocation As New dseTransaksiStockInOut.TBLMasterLocationDataTable
        Dim rowLocatio As dseTransaksiStockInOut.TBLMasterLocationRow
        Dim Lanjut As Boolean = False

        Try
            'DTAImport.Connection.ConnectionString = Core_Module.GetConstring
            DTALocation.Connection.ConnectionString = Core_Module.GetConstring

            ij.dgvItemProductGroup.AllowUserToAddRows = False

            DTAImport.Fill(TBLImport)
            DTALocation.Fill(TBLLocation)

            For Each rowDGV1 In ij.dgvItemProductGroup.Rows
                Level = Me.objItem.GetProductLevel(rowDGV1.Cells(ij.ColProductGroup_Level.Index).Value)
                If Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                    DTA.FillBy_PrimaryProductGroup(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                    DTA.FillBy_PrimaryBrand(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                    DTA.FillBy_PrimarySubBrand(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                    DTA.FillBy_PrimaryCategory(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                Else
                    DTA.FillBy_PrimaryCategory(TBLDummy, CInt(Tipe), CDec(rowDGV1.Cells(ij.ColProductGroup_PrimaryAcuan.Index).Value))
                End If

                For Each rowDummy In TBLDummy.Rows
                    TBL.DefaultView.RowFilter = "Primary = " + rowDummy.Primary.ToString
                    If TBL.DefaultView.Count = 0 Then
                        row = TBL.Newtblmaster_product1Row
                        For i As Integer = 0 To TBL.Columns.Count - 1
                            row.Item(i) = rowDummy.Item(i)
                        Next
                        TBL.Rows.Add(row)
                    End If
                Next
                TBLDummy.Rows.Clear()
            Next

            For Each rowImport In TBLImport.Rows
                Lanjut = True

                '1. cek kode barang barus ada
                TBL.DefaultView.RowFilter = "[Product Code] = '" + rowImport.CODE_BARANG.Trim + "'"
                If TBL.DefaultView.Count > 0 Then
                    row = TBL.DefaultView(0).Row
                    'cek profile kontrolnya sesuai
                    'location
                    If row.Location_Controlled = 1 Then
                        'cek lokasi ada
                        If rowImport.IsLOCATIONNull = False Then
                            TBLLocation.DefaultView.RowFilter = "[Storage Area Code] = '" + rowImport.LOCATION + "'"
                            If TBLLocation.DefaultView.Count <= 0 Then
                                Lanjut = False
                            End If
                        Else
                            Lanjut = False
                        End If
                    Else
                        'cek lokasi tidak ada
                        If rowImport.IsLOCATIONNull = False Then
                            Lanjut = False
                        End If
                    End If
                    'bacth no
                    If row.Batch_No_Controlled = 0 Then
                        If rowImport.IsBATCH_NONull = True Then
                            'ok
                        Else
                            Lanjut = False
                        End If
                    Else
                        If rowImport.IsBATCH_NONull = True Then
                            Lanjut = False
                        Else
                            'ok
                            If rowImport.BATCH_NO.Trim = "" Then
                                Lanjut = False
                            Else
                                If row.Expiry_Date_Controlled = True Then
                                    If rowImport.IsEXPIRYNull = True Then
                                        Lanjut = False
                                    End If
                                Else
                                    If rowImport.EXPIRY <> #1/1/1950# Then
                                        Lanjut = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                    'qty
                    If rowImport.SALDO <= 0 Then
                        Lanjut = False
                    End If

                    If Lanjut = True Then
                        rowLocatio = TBLLocation.DefaultView(0).Row

                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        With rowDGV
                            .Cells(Me.ColAmount.Index).Value = 0
                            If rowImport.IsBATCH_NONull = True Then
                                .Cells(Me.ColBatch.Index).Value = ""
                                .Cells(Me.ColExpiredDate.Index).Value = "-"
                            Else
                                .Cells(Me.ColBatch.Index).Value = rowImport.BATCH_NO
                                If rowImport.EXPIRY.Date = #1/1/1950# Then
                                    .Cells(Me.ColExpiredDate.Index).Value = "-"
                                Else
                                    .Cells(Me.ColExpiredDate.Index).Value = Format(rowImport.EXPIRY.Date, "dd MMM yyyy")
                                End If
                            End If
                            .Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                            .Cells(Me.ColDescription.Index).Value = row.Description
                            .Cells(Me.ColLocation.Index).Value = rowImport.LOCATION
                            .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                            .Cells(Me.ColPONo.Index).Value = ""
                            .Cells(Me.ColPrimaryCodeNo.Index).Value = row.Primary
                            .Cells(Me.ColPrimaryLocation.Index).Value = rowLocatio.Primary_Location
                            .Cells(Me.ColPrimaryPdO.Index).Value = -1
                            .Cells(Me.ColPrimaryPOno.Index).Value = -1
                            .Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                            .Cells(Me.ColPrimaryUnits.Index).Value = row.Primary_Unit
                            .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowImport.SALDO, row._Decimal)
                            .Cells(Me.ColRemark.Index).Value = ""
                            .Cells(Me.ColSecondaryQty.Index).Value = -1
                            .Cells(Me.ColSecondaryUnit.Index).Value = ""
                            .Cells(Me.ColSuppDONo.Index).Value = ""
                            .Cells(Me.ColUnit.Index).Value = row.Unit_Code
                            .Cells(Me.ColWareHouse.Index).Value = rowLocatio.Warehouse_Code
                        End With
                    Else
                        'MsgBox("Skip")
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAImport.Dispose()
            TBLImport.Dispose()
            DTALocation.Dispose()
            TBLLocation.Dispose()
        End Try
    End Sub

    Private Sub ImportMirror(ByVal ij As FrmTransaksiStockInAdj_QuickInput)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tblmaster_product1TableAdapter
        Dim TBL_A As New dseTransaksiStockInOut.tblmaster_product1DataTable
        Dim TBLDummy As New dseTransaksiStockInOut.tblmaster_product1DataTable
        Dim row, rowDummy As dseTransaksiStockInOut.tblmaster_product1Row
        Dim Tipe As clsMasterProduct.EnumTipeProduct = clsMasterProduct.EnumTipeProduct.Goods
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim ExpiryRef As Date = #1/1/2500#

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
                    .Cells(Me.ColAmount.Index).Value = 0
                    .Cells(Me.ColBatch.Index).Value = BatchNumber
                    If Expiry = #1/1/1950# Then
                        .Cells(Me.ColExpiredDate.Index).Value = "-"
                    Else
                        .Cells(Me.ColExpiredDate.Index).Value = Format(Expiry, "dd MMM yyyy")
                    End If
                    .Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                    .Cells(Me.ColDescription.Index).Value = row.Description
                    .Cells(Me.ColLocation.Index).Value = LocationCode
                    .Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColPONo.Index).Value = ""
                    .Cells(Me.ColPrimaryCodeNo.Index).Value = rowQty.Primary_Product
                    .Cells(Me.ColPrimaryLocation.Index).Value = PrimaryLocation
                    .Cells(Me.ColPrimaryPdO.Index).Value = -1
                    .Cells(Me.ColPrimaryPOno.Index).Value = -1
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
                    'rowInOut = TBL.Rows(TBL.Rows.Count - 1)
                    filter = Me.GetPrimaryMax(TBL)
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

    Private Function GetPrimaryMax(ByVal TBL As dseLaporanStockTaking.tblmaster_productqtyDataTable) As Double
        Dim rowInOut As dseLaporanStockTaking.tblmaster_productqtyRow
        Dim RetVal As Double = 0

        Try
            For Each rowInOut In TBL.Rows
                If RetVal < rowInOut.Primary Then
                    RetVal = rowInOut.Primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LinkLabel1_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSetAmount.LinkClicked
        Dim rowDGV As DataGridViewRow
        Dim RefVal As String = ""

        Try
            'copy nilai amountnya dari baris pertama
            If MsgBox("Copy amount data from first row?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                rowDGV = Me.dgvItem.Rows(0)
                RefVal = rowDGV.Cells(Me.ColAmount.Index).Value
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each rowDGV In Me.dgvItem.Rows
                    If rowDGV.Index <> 0 Then
                        rowDGV.Cells(Me.ColAmount.Index).Value = RefVal
                    End If
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
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
                i.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInAdjust
                i.ShowDialog()

                i = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub EditTanggal(ByVal vTglTujuan As Date)
        'Dim hasil As MsgBoxResult
        Try
            If Me.TxtStockInNo.Text <> "" Then
                If Me.objStockIn.ApakahBolehEdit(Me.UserCode) = True Then
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
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub lblQuickInput2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblQuickInput2.LinkClicked
        Dim hasil As MsgBoxResult
        Try
            'konfirmasi proses atau tidak
            hasil = MsgBox("Quick input?", MsgBoxStyle.YesNoCancel, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                objStockIn.GetQtySelisihInAdjust(#2/29/2016#, Me.dgvItem)
            ElseIf hasil = MsgBoxResult.No Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.IsiDGVQuickIput2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LlbUpdDecProd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUpdDecProd.LinkClicked
        Try
            Dim hasil As MsgBoxResult
            'konfirmasi proses atau tidak
            hasil = MsgBox("Update Product Qty (Decimal)?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If objStockIn.UpdateDecProductQty() = True Then
                    MsgBox("Update success", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub IsiDGVQuickIput2()
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
        Dim BatchNoRef As String = ""
        Dim ExpiryRef As Date = #1/1/0950#
        Dim PrimaryLocationRef As Double = -1
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

                    If decSelisih > totalStok Then
                        ListX.Add(.ColCodeNo)
                        ' decDetStok = decSelisih
                        If tblStok.DefaultView.Count > 0 Then
                            rowStok = tblStok.DefaultView(0).Row
                            '    vDgv.Rows.Add("", "", "", .Product_Code, .Description, Qty_Akhir, .Unit_Code, -1, "", .Lot_No, _
                            'ExpiredDateStr, .Storage_Area_Code, .Warehouse_Code, "", -1, .Primary_Product, .Primary_Storage_Area, _
                            '.Primary_Unit, -1, -1, 0)
                            If rowStok.Expiry_Date = #1/1/1950# Then
                                Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                                                                                               .ColDescription, Me.ObjCD.DigitGrouping(CDec(decSelisih - totalStok), .DefPrec), _
                                                                                               .ColUnit, -1, "", rowStok.Lot_No, _
                                                                                               "-", rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                                                                                               -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                                                                                               -1, -1, 0)
                            Else
                                Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                                                               .ColDescription, Me.ObjCD.DigitGrouping(CDec(decSelisih - totalStok), .DefPrec), _
                                                               .ColUnit, -1, "", rowStok.Lot_No, _
                                                               Format(rowStok.Expiry_Date, "dd MMM yyyy"), rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                                                               -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                                                               -1, -1, 0)
                            End If
                        Else
                            If Me.ObjProduct.ApakahBatchNoControlled(.ColPrimaryCodeNo) = True Then
                                BatchNoRef = "AWAL"
                            Else
                                BatchNoRef = ""
                            End If
                            If Me.ObjProduct.ApakahExpiryDateControlled(.ColPrimaryCodeNo) = True Then
                                ExpiryRef = #1/1/2016#
                            Else
                                ExpiryRef = #1/1/1950#
                            End If
                            If Me.ObjProduct.ApakahLocationControlled(.ColPrimaryCodeNo) = True Then
                                PrimaryLocationRef = 1
                            Else
                                PrimaryLocationRef = -1
                            End If
                            Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                                                               .ColDescription, Me.ObjCD.DigitGrouping(CDec(decSelisih - totalStok), .DefPrec), _
                                                               .ColUnit, -1, "", BatchNoRef, _
                                                               ExpiryRef, Me.ObjLocation.GetStorageAreaCode(PrimaryLocationRef), Me.ObjLocation.GetWarehouseCode(PrimaryLocationRef), .ColRemark, _
                                                               -1, .ColPrimaryCodeNo, PrimaryLocationRef, .ColPrimaryUnits, _
                                                               -1, -1, 0)
                        End If

                    Else
                        'Dim intA As Integer = 1
                        'If tblStok.DefaultView.Count > 0 Then
                        '    While decSelisih <> decDetStok
                        '        rowStok = tblStok.DefaultView(tblStok.DefaultView.Count - intA).Row
                        '        If (decSelisih - decDetStok) <= CDec(rowStok.Qty) Then
                        '            Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                        '            .ColDescription, Me.ObjCD.DigitGrouping((decSelisih - decDetStok), .DefPrec), _
                        '             .ColUnit, -1, "", rowStok.Lot_No, _
                        '             rowStok.Expiry_Date, rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                        '             -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                        '             -1, -1)
                        '            decDetStok += (decSelisih - decDetStok)
                        '        ElseIf (decSelisih - decDetStok) > CDec(rowStok.Qty) Then
                        '            Me.dgvItem.Rows.Add(Me.dgvItem.RowCount + 1, "", "", .ColCodeNo, _
                        '            .ColDescription, Me.ObjCD.DigitGrouping(CDec(rowStok.Qty), .DefPrec), _
                        '            .ColUnit, -1, "", rowStok.Lot_No, _
                        '            rowStok.Expiry_Date, rowStok.Storage_Area_Code, rowStok.Warehouse_Code, .ColRemark, _
                        '            -1, .ColPrimaryCodeNo, rowStok.Primary_Storage_Area, .ColPrimaryUnits, _
                        '            -1, -1)
                        '            decDetStok += CDec(rowStok.Qty)
                        '        End If
                        '        If tblStok.DefaultView.Count - intA > 0 Then
                        '            intA += 1
                        '        End If
                        '    End While
                        'End If


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

            'If ListX.Count > 0 Then
            '    Dim msg As String = ""
            '    For intA As Integer = 0 To ListX.Count - 1
            '        If msg = "" Then
            '            msg &= ListX.Item(intA)
            '        Else
            '            msg &= ", " & ListX.Item(intA)
            '        End If
            '    Next
            '    If msg <> "" Then
            '        MsgBox("There are list of code no smess with condition stok for now > stok 29 feb 2016 : " & vbNewLine & msg, MsgBoxStyle.Information, "Information")
            '    End If
            'End If

            Me.FormatDGV()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_BeginningBalance(ByVal vUserCode As String, ByVal rowQty As DseLaporanStockDetail2.TblCetakRow, _
      ByVal vEndDateCutOff As Date, ByVal vBaru As Boolean, ByVal vSimpan As Boolean) As Boolean
        Dim rowDgv As DataGridViewRow

        Try
            'Buat baru
            If vBaru = True Then
                Me.BeginingBalance = True
                Me.UserCode = vUserCode

                Me.baru = True
                Me.objStockIn.ResetObject(Me.dgvItem)
                Me.objStockIn.TanggalTransaksi = vEndDateCutOff
                Me.objStockIn.NoTransaksi = Me.objStockIn.GetNoBaru(Me.TipeTransaksi, vEndDateCutOff)
            End If

            'Isikan item
            With rowQty
                If .Lot_No = "-" Then
                    .Lot_No = ""
                End If
                rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Add())
                rowDgv.Cells(Me.ColQty.Index).Value = .QtyDec
                rowDgv.Cells(Me.ColSecondaryQty.Index).Value = -1
                rowDgv.Cells(Me.ColBatch.Index).Value = .Lot_No
                rowDgv.Cells(Me.ColExpiredDate.Index).Value = .Expiry_Date
                rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                rowDgv.Cells(Me.ColPrimaryLocation.Index).Value = Me.ObjLocation.GetPrimary(.Storage_Area_Code)
                rowDgv.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Unit
                rowDgv.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                rowDgv.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(.Item("Amount Final"), 0)
                rowDgv.Cells(Me.ColPrimaryPOno.Index).Value = -1
                rowDgv.Cells(Me.ColPrimaryPdO.Index).Value = -1
                rowDgv.Cells(Me.ColRemark.Index).Value = ""
                rowDgv.Cells(Me.ColSuppDONo.Index).Value = ""
            End With

            'Simpan StockIn
            If vSimpan = True Then
                If Me.objStockIn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, -1, False, False, True) = False Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Sub FromCutOff_InAdjust_Create(ByVal vPrimaryTransaksi As Double)
        Try
            Me.baru = False
            Me.BeginingBalance = False
            Me.objStockIn.ResetObject(Me.dgvItem)
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust

            Me.objStockIn.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.TblItemCustom)
            Me.UserCode = Me.objStockIn.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_InAdjust_Save() As Double
        Try
            Me.baru = True
            Me.objStockIn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, -1, False, False, True)
            Return Me.objStockIn.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub llbQuickInputBundling_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbQuickInputBundling.LinkClicked
        Dim ij As New FrmCariTransaksiStockBundling
        Dim ExpiredDateStr As String = ""

        Try
            'find bundling yang dimaksud
            ij.TipeBundling = Me.ObjStockBundling.TipeBundling
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If Not ij.PrimaryTransaksiStockBundling = -1 Then
                'fill item after dari bundling tersebut
                Me.TBLItemAft.Rows.Clear()
                Me.DTAItemAft.Fill(Me.TBLItemAft, ij.PrimaryTransaksiStockBundling)

                'masukkan ke dgv
                If Me.TBLItemAft.Rows.Count > 0 Then
                    Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                    For Each Me.RowItemAft In Me.TBLItemAft.Rows
                        With Me.RowItemAft
                            If .Expiry_Date.Date = #1/1/1950# Then
                                ExpiredDateStr = "-"
                            Else
                                ExpiredDateStr = Format(.Expiry_Date.Date, "dd MMM yyyy")
                            End If

                            Me.dgvItem.Rows.Add(Me.dgvItem.Rows.Count + 1, "", "", .Product_Code, .Description, Me.ObjCD.DigitGrouping(.Qty, ._decimal), _
                            .Unit_Code, -1, "", .Batch_No, ExpiredDateStr, .StorageAreaCode, .Warehouse_Code, "", _
                            -1, .Primary_Product, .Primary_Storage_Area, .Primary_Units, -1, -1, 0)
                        End With
                    Next

                    ' Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                End If
            End If
            If Me.dgvItem.Rows.Count > 0 Then
                Me.FormatDGV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

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
            'ij.ParentForm1 = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_In_Adjust
            'ij.ShowDialog()
            Obj.GetDataInput(TBLData, InputType, False, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_In_Adjust, Me.dgvItem, Me.TxtPrimaryStockIn.Text, Me.UserCode)
            'Obj.GetDataInput(TBLData, InputType, False, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Stock_In_Adjust)
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
        Dim rowData As Core_Module.DseBarcodeInput.DataTable1Row
        Dim InputType As CustomSMESS.ClsCustomBarcode.EnumInputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add
        Dim Index As Integer = -1

        Try
            Obj.GetDataInput(TBLData, InputType, False, CustomSMESS.ClsCustomBarcode.EnumTransaksiAcuan.InAdjust)

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

    Private Sub tsm_QuickInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsm_QuickInput.Click
        Try
            Me.QuickInput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cms1_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cms1.Opening
        Try
            Me.tsm_QuickInput.Enabled = Me.QuickAddByProdLevel
            Me.tsm_AddByRefDoc.Enabled = Me.QuickAddByRefDoc
            Me.tsm_AddByBarcode.Enabled = Me.QuickAddByBarcode
            Me.tsm_QuickInput.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex2(ByVal vPrimaryProduct As Double, ByVal BatchNo As String, _
  ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal vPrimaryLocation As Double, ByVal vLocationControl As Integer) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If vLocationControl = 1 Then
                    If IsNumeric(row.Cells(7).Value) = False Then
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo Then
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
                        row.Cells(9).Value = BatchNo Then
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

#Region "tambahan ahdar"
    'sisanya dilain2 event/function
    Private Sub AturHakAksesFormCustom()
        Dim ObjUser As New Core_Module.ClsUser
        Dim Col1 As DataGridViewColumn

        Try
            Me.ColSecondaryQty.Visible = False
            Me.ColSecondaryUnit.Visible = False

            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockInAdj_AddItembyCustom") = True Then
                For Each Col1 In Me.dgvItem.Columns
                    Col1.ReadOnly = True
                Next
                Me.dgvItem.AllowUserToDeleteRows = False
                Me.TxtCodeNo.Enabled = False
                Me.BtnFindCodeNo.Enabled = False
                Me.txtDesc.Enabled = False
                Me.txtBatchNo.Enabled = False
                Me.DtpExpired.Enabled = False
                Me.TxtWarehouse.Enabled = False
                Me.TxtQtyperBatch.Enabled = False
                Me.TxtRemarkItem.Enabled = False
                Me.BtnInput.Enabled = False
                Me.btnFindLocation.Enabled = False
            Else
                For Each Col1 In Me.dgvItem.Columns
                    Col1.ReadOnly = False
                Next
                Me.cms1.Visible = False
                Me.llbAdvance.Visible = False
                Me.dgvItem.AllowUserToDeleteRows = True
                Me.TxtCodeNo.Enabled = True
                Me.BtnFindCodeNo.Enabled = True
                Me.txtDesc.Enabled = True
                Me.txtBatchNo.Enabled = True
                Me.DtpExpired.Enabled = True
                Me.TxtLocation.Enabled = True
                Me.TxtWarehouse.Enabled = True
                Me.TxtQtyperBatch.Enabled = True
                Me.TxtRemarkItem.Enabled = True
                Me.btnFindLocation.Enabled = True
            End If

            Me.llbAdvance.Enabled = True
            Me.ColRemark.ReadOnly = False
            Me.ColAmount.ReadOnly = False
            Me.ColUnitCost.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiStockInAdj_AddItembyCustom
        Try
            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TanggalTransaksi = Me.dtpDate.Value
            i.PrimaryStockIn = Me.TxtPrimaryStockIn.Text
            i.TblItemCustom = Me.TblItemCustom
            i.Baru = Me.baru
            If Me.TxtRemark.Enabled = False Then
                i.AktifForm = False
            Else
                i.AktifForm = True
            End If

            If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblItemCustom = i.TblItemCustom
                Me.IsikanDgvFromTblDummy()
            End If
            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsikanDgvFromTblDummy() As Boolean
        Dim rowDGV As New DataGridViewRow
        Dim Row As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemRow
        Dim DefPrecCur As Integer = 0

        Try
            'isikan nilai amount dan unit cost pada tbl dummy
            'MsgBox("belum coding untuk reserve amount dan unit cost")
            DefPrecCur = Me.ObjCur.GetDecimal_DefaultCurrency

            For Each rowDGV In Me.dgvItem.Rows
                For Each Row In Me.TblItemCustom.Rows
                    With Row
                        If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product And _
                        rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area And _
                        rowDGV.Cells(Me.ColBatch.Index).Value = .Batch_No Then
                            If CDec(rowDGV.Cells(Me.ColQty.Index).Value) = CDec(Row.Qty) Then
                                .Unit_Cost = CDec(rowDGV.Cells(Me.ColUnitCost.Index).Value)
                                .Amount_1 = CDec(rowDGV.Cells(Me.ColAmount.Index).Value)
                            Else
                                .Unit_Cost = rowDGV.Cells(Me.ColUnitCost.Index).Value
                                .Amount_1 = Me.ObjCD.DigitGrouping(Row.Qty * .Unit_Cost, DefPrecCur)
                            End If

                        End If
                    End With
                Next
            Next

            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each Row In Me.TblItemCustom.Rows
                Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(Row.Primary_Product)

                Dim Index As Integer = Me.GetIndexCustom(Row.Primary_Product, Row.Primary_Storage_Area, Row.Batch_No, _
                                           -1, Row.Primary_PdO, Row.Expiry_Date)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim rowD As DataGridViewRow = Me.dgvItem.Rows(Index)
                    rowD.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowD.Cells(Me.ColQty.Index).Value) + Row.Qty, DefPrecDefault)
                    rowD.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(rowD.Cells(Me.ColAmount.Index).Value + Row.Amount_1, DefPrecCur)
                    rowD.Cells(Me.ColUnitCost.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowD.Cells(Me.ColAmount.Index).Value) / CDec(rowD.Cells(Me.ColQty.Index).Value), DefPrecCur)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Row.Expiry_Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Row.Expiry_Date, "dd MMM yyyy")
                    End If

                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPONo.Index).Value = ""
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = Row.Product_Code
                        .Cells(Me.ColDescription.Index).Value = Row.Description
                        .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty, DefPrecDefault)
                        .Cells(Me.ColUnit.Index).Value = Row.unit

                        .Cells(Me.ColSecondaryQty.Index).Value = "-1"
                        .Cells(Me.ColSecondaryUnit.Index).Value = ""

                        'If Row.Secondary_Qty = -1 Then
                        '    .Cells(Me.ColSecondaryQty.Index).Value = ""
                        '    .Cells(Me.ColSecondaryUnit.Index).Value = ""
                        'Else
                        '    .Cells(Me.ColSecondaryQty.Index).Value = Me.ObjCD.DigitGrouping(Row.Secondary_Qty, DefPrecDefault)
                        '    .Cells(Me.ColSecondaryUnit.Index).Value = Row.unit
                        'End If

                        .Cells(Me.ColBatch.Index).Value = Row.Batch_No
                        .Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColLocation.Index).Value = Row.Storage_Area_Code
                        .Cells(Me.ColWareHouse.Index).Value = Row.Warehouse_Code
                        .Cells(Me.ColRemark.Index).Value = Row.Remarks
                        .Cells(Me.ColPrimaryPOno.Index).Value = "-1"
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Row.Primary_Product
                        .Cells(Me.ColPrimaryLocation.Index).Value = Row.Primary_Storage_Area
                        .Cells(Me.ColPrimaryUnits.Index).Value = Row.Primary_Units
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = Row.Primary_Secondary_Units
                        .Cells(Me.ColPrimaryPdO.Index).Value = Row.Primary_PdO
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Row.Amount_1, DefPrecCur)

                        If Row.IsUnit_CostNull = False Then
                            .Cells(Me.ColUnitCost.Index).Value = Me.ObjCD.DigitGrouping(Row.Unit_Cost, DefPrecCur)
                        Else
                            .Cells(Me.ColUnitCost.Index).Value = ""
                        End If
                    End With
                End If
            Next
            'set ulang nomor
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.FormatDGV()

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.ColSecondaryUnit.Visible = False
            Me.ColSecondaryQty.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexCustom(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(10).Value) = False Then
                    If IsNumeric(row.Cells(7).Value) = False Then
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And row.Cells(19).Value = PrimaryPdO _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And row.Cells(19).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(7).Value) = False Then
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And row.Cells(19).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(10).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(15).Value = vPrimaryProduct And row.Cells(16).Value = vPrimaryLocation And _
                        row.Cells(9).Value = BatchNo And row.Cells(19).Value = PrimaryPdO And _
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
#End Region
End Class