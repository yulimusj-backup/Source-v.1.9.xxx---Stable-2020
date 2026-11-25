Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockInSalesReturn3
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
    Dim baru As Boolean
    Dim objStockIn As New clsTransaksiStockInOut
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjStockInOutItem As New clsTransaksiStockInOutItem
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjItemInOut As New clsTransaksiStockInOut
    Dim ObjItemBundling As New clsTransaksiStockBundlingItem
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjDO As New ClsTransaksiDO
    Dim ObjDOItem As New ClsTransaksiDOItem
    Dim ObjLS As New clsMasterListedSoftware
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim buffer As String = ""

    Dim TBLSO As New DseTransaksiStockInSalesReturn_Balance.tbltrans_do_itemDataTable
    Dim DTASO As New DseTransaksiStockInSalesReturn_BalanceTableAdapters.tbltrans_do_itemTableAdapter
    Dim TBLSO_StockInSalesReturn As New DseTransaksiStockInSalesReturn_Balance.tbltransaksi_stockinout_itemDataTable
    Dim DTASO_StockInSalesReturn As New DseTransaksiStockInSalesReturn_BalanceTableAdapters.tbltransaksi_stockinout_itemTableAdapter

    Dim TBLCodeNo As New DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item1DataTable
    Dim DTACodeNo As New DseTransaksiStockInSalesReturn_BalanceTableAdapters.tbltrans_do_item1TableAdapter
    Dim TBLCodeNo_InSalesReturn As New DseTransaksiStockInSalesReturn_Balance.tbltransaksi_stockinout_item1DataTable
    Dim DTACodeNo_InSalesReturn As New DseTransaksiStockInSalesReturn_BalanceTableAdapters.tbltransaksi_stockinout_item1TableAdapter

    Dim TBLBatchNo As New DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item2DataTable
    Dim DTABatchNo As New DseTransaksiStockInSalesReturn_BalanceTableAdapters.tbltrans_do_item2TableAdapter
    Dim TBLBacthNo_InSalesReturn As New DseTransaksiStockInSalesReturn_Balance.tbltransaksi_stockinout_item2DataTable
    Dim DTABacthNo_InSalesReturn As New DseTransaksiStockInSalesReturn_BalanceTableAdapters.tbltransaksi_stockinout_item2TableAdapter

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)

            Me.cboSO_Number.DataSource = Nothing
            Me.cboSO_Number.Items.Clear()
            Me.cboCodeNo.DataSource = Nothing
            Me.cboCodeNo.Items.Clear()
            Me.CboBatchNo.DataSource = Nothing
            Me.CboBatchNo.Items.Clear()
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Items.Clear()
            Me.cboSecQty.DataSource = Nothing
            Me.cboSecQty.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            Me.TxtPrimaryStockIn.Text = "-1"
            Me.TxtPrimaryPihakAcuan.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()

            Me.cboSO_Number.DataSource = Nothing
            Me.cboSO_Number.Items.Clear()
            Me.cboCodeNo.DataSource = Nothing
            Me.cboCodeNo.Items.Clear()
            Me.CboBatchNo.DataSource = Nothing
            Me.CboBatchNo.Items.Clear()
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Items.Clear()
            Me.cboSecQty.DataSource = Nothing
            Me.cboSecQty.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimarySO.Text = "-1"
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

            '9. Bila setelah diinput jumlah product menjadi minus, kalo input baru ga mungkin
            '10. Bila transaksi berkaitan dengan sales dan terinstall software sales maka ada batasan 
            '    berkaitan dengan jumlah item DO. Berkaitan dengan jumlah item DO secara global, 
            '    hanya memperdulikan primary SO dan primary product
            Dim JmlDo As Double = Me.ObjDO.GetJml3(Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, _
            Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text, Me.txtSecondaryQty.Text)
            Dim JmlDGV As Double = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtBatchNo.Text, _
            Me.txtPrimarySO.Text)
            Dim JmlObj, JmlReturn As Double
            JmlReturn = Me.ObjStockInOutItem.GetJmlReturn(Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, _
            Me.txtPrimaryDO.Text, Me.txtBatchNo.Text, Me.txtSecondaryQty.Text)
            If Me.baru = True Then
                JmlObj = 0
            Else
                JmlObj = Me.ObjStockInOutItem.GetJmlReturn(Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.txtPrimaryDO.Text, Me.TxtPrimaryStockIn.Text, Me.txtBatchNo.Text, Me.txtSecondaryQty.Text)
            End If

            If JmlDo - JmlDGV - CDbl(Me.TxtQtyperBatch.Text) - JmlReturn + JmlObj < 0 Then
                Dim Msg As String = "Maximum value you can enter for:" + Me.TxtCodeNo.Text
                If Me.txtBatchNo.Text <> "" Then
                    Msg = Msg + ", batch no: " + Me.txtBatchNo.Text
                End If
                If Me.txtSecondaryQty.Text <> "-1" And Me.txtSecondaryQty.Text <> "" Then
                    Msg = Msg + ", secondary qty: " + Me.txtSecondaryQty.Text + " " + _
                    Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
                MsgBox(Msg + " and SO number: " + _
                Me.txtSalesOrderNo.Text + " is " + CStr(JmlDo - JmlDGV - JmlReturn + JmlObj) + " " + _
                CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatch.Focus()
                Return False
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

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vBatchNo As String, ByVal vPrimarySO As Double) As Double
        Try
            Dim row As DataGridViewRow
            Dim Jml As Double = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(15).Value = vPrimaryCodeNo And row.Cells(14).Value = vPrimarySO And _
                row.Cells(9).Value = vBatchNo Then
                    If IsNumeric(row.Cells(7).Value) = True Then
                        If row.Cells(7).Value = Secondary_Qty Then
                            Jml += row.Cells(5).Value
                        End If
                    Else
                        If Secondary_Qty = -1 Then
                            Jml += 1
                        End If
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
                MsgBox("Please specify customer's code", MsgBoxStyle.Information, "Information")
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
            With Me.objStockIn
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
            With Me.objStockIn
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryStockIn.Text = .Primary
                Me.txtPrimaryDO.Text = .PrimaryDOSales
                Me.txtDONo.Text = Me.ObjDO.GetDoNo(Me.txtPrimaryDO.Text)
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
                Me.TxtPrimaryPihakAcuan.Text = .PrimaryCustomer
                Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
            End With
            Me.FormatDGV()
            'isikan data DO
            Me.ObjDOItem.IsiDGV(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
            Me.FormatDGVDO()
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
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemDO2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.KosongkanField()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
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
                        If Me.objStockIn.SimpanBaru(Me.dgvItem, Me.UserCode) Then
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
                        If Me.objStockIn.SimpanPerubahan(Me.dgvItem, Me.UserCode) Then
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
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objStockIn.CariDetail(Me.TxtPrimaryStockIn.Text, Me.dgvItem)
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
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
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
        Dim hj As New FrmCariTransaksiStockInOut
        Try
            hj.JenisTransaksi = Me.TipeTransaksi
            hj.ShowDialog()
            If Not hj.PrimaryTransaksiStockInOut = -1 Then
                If Me.objStockIn.CariDetail(hj.PrimaryTransaksiStockInOut, Me.dgvItem) = True Then
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
                If Me.objStockIn.ApakahBolehEdit(Me.UserCode, True) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()

                    'isikan cbo so
                    Me.IsiCBOSO()
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
                If Me.objStockIn.ApakahBolehHapus(Me.UserCode) = False Then
                    MsgBox("You cannot delete this data, because this will cause negative quantity data or there is sales return refers to this transaction.", _
                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockIn.Hapus(Me.TxtPrimaryStockIn.Text, Me.UserCode) = True Then
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
                            If Me.objStockIn.SimpanBaru(Me.dgvItem, Me.UserCode) = True Then
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
                            If Me.objStockIn.SimpanPerubahan(Me.dgvItem, Me.UserCode) = True Then
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

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Double
        Try
            Dim row As DataGridViewRow
            Dim Jml As Double = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(15).Value = vPrimaryCodeNo And vPrimarySO = row.Cells(14).Value Then
                    Jml = Jml + CDbl(row.Cells(5).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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
                    JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, rowSONO.Primary_SO)
                    JmlDO = Me.ObjDO.GetJml2(Me.txtPrimaryDO.Text, rowSONO.Primary_SO, Me.TxtPrimaryCodeNo.Text)
                    If JmlDO - JmlDGV <> 0 Then
                        'ganti no sales order
                        Me.txtPrimarySO.Text = rowSONO.Primary_SO
                        Me.txtSalesOrderNo.Text = rowSONO.So_No
                        Exit For
                    End If
                Next
            Else
                'hitung ulang jumlah so, jangan sampe minus
                JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
                JmlDO = Me.ObjDO.GetJml2(Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)
                If JmlDO - JmlDGV = 0 Then
                    'ganti no sales order
                    Me.ObjDOItem.IsiListNoSO(Me.txtPrimaryDO.Text, Me.TxtPrimaryCodeNo.Text, TBLSO)
                    For Each rowSONO In TBLSO.Rows
                        If rowSONO.Primary_SO <> Me.txtPrimarySO.Text Then
                            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, rowSONO.Primary_SO)
                            JmlDO = Me.ObjDO.GetJml2(Me.txtPrimaryDO.Text, rowSONO.Primary_SO, Me.TxtPrimaryCodeNo.Text)
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

    Private Sub SetGboItem()
        Try
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
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            Me.TxtQtyperBatch.Focus()
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
                    If Me.txtBatchNo.Visible = True Then
                        Me.txtBatchNo.Focus()
                    ElseIf Me.TxtLocation.Visible = True Then
                        Me.TxtLocation.Focus()
                    Else
                        Me.TxtQtyperBatch.Focus()
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
                        If Me.txtBatchNo.Visible = True Then
                            Me.txtBatchNo.Focus()
                        ElseIf Me.TxtLocation.Visible = True Then
                            Me.TxtLocation.Focus()
                        Else
                            Me.TxtQtyperBatch.Focus()
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
    Private Function GetIndex1(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
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
                Dim Index As Integer = Me.GetIndex1(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(5).Value = CDbl(row.Cells(5).Value) + CDbl(Me.TxtQtyperBatch.Text)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItem.Rows.Add("", Me.txtSalesOrderNo.Text, "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatch.Text, _
                    Me.lblUnits.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.txtBatchNo.Text, _
                    ExpiredDateStr, Me.TxtLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnits.Text, _
                    Me.txtPrimarySecondaryUnits.Text, Me.txtPrimaryPDO.Text)
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
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 Or e.ColumnIndex = 11 _
                Or e.ColumnIndex = 13 Or e.ColumnIndex = 10 Then
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
                        ElseIf e.ColumnIndex = 10 Then
                            Me.EndEditExpiryDate(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
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
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.ObjProduct.ApakahGoods(.Cells(15).Value) = True Then
                    JmlDGV = .Cells(5).Value
                    If Me.ApakahEditSah(row, .Cells(16).Value, .Cells(19).Value, .Cells(15).Value, _
                    .Cells(9).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If

                'cek data sales order balance
                Dim JmlDo As Decimal = Me.ObjDO.GetJml3(Me.txtPrimaryDO.Text, .Cells(14).Value, _
                .Cells(15).Value, .Cells(9).Value, .Cells(7).Value)
                JmlDGV = Me.GetJmlDGV(.Cells(15).Value, .Cells(7).Value, .Cells(9).Value, .Cells(14).Value)
                Dim JmlDGVSendiri As Decimal = .Cells(5).Value
                Dim JmlObj, JmlReturn As Decimal
                JmlReturn = Me.ObjStockInOutItem.GetJmlReturn(.Cells(14).Value, .Cells(15).Value, _
                Me.txtPrimaryDO.Text, .Cells(9).Value, .Cells(7).Value)
                If Me.baru = True Then
                    JmlObj = 0
                Else
                    JmlObj = Me.ObjStockInOutItem.GetJmlReturn(.Cells(14).Value, .Cells(15).Value, _
                    Me.txtPrimaryDO.Text, Me.TxtPrimaryStockIn.Text, .Cells(9).Value, .Cells(7).Value)
                End If

                If JmlDo - JmlDGV - JmlReturn + JmlObj < 0 Then
                    Dim Msg As String = "Maximum value you can enter for:" + CStr(.Cells(3).Value)
                    If .Cells(9).Value <> "" Then
                        Msg = Msg + ", batch no: " + .Cells(9).Value
                    End If
                    If .Cells(7).Value <> "-1" Then
                        Msg = Msg + ", secondary qty: " + .Cells(7).Value + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(.Cells(15).Value)
                    End If

                    MsgBox(Msg + " and SO number: " + _
                    CStr(.Cells(1).Value) + " is " + CStr(JmlDo - JmlDGV + JmlDGVSendiri - JmlReturn + JmlObj) + " " + _
                    CStr(Me.ObjProduct.GetUnitCode(.Cells(15).Value)) + ".", MsgBoxStyle.Information, _
                    "Information")

                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value)
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
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
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
                            'kembalikan data gudang
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
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal As Double
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.ObjStockInOutItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Minimum value you must can enter is " + CStr(JmlObjLama - JmlGudang) + _
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
            Me.IsiCBOSO()
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
            'yang dicek hanya yang goods saja
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
            Me.TxtRemark.BackColor = BackColorAll
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
            Me.TxtCodeNo.BackColor = BackColorAll
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
            Me.txtBatchNo.BackColor = BackColorAll
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
            Me.TxtLocation.BackColor = BackColorAll
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
            Me.TxtQtyperBatch.BackColor = BackColorAll
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
            Me.txtSecondaryQty.BackColor = BackColorAll
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
            Me.TxtRemarkItem.BackColor = BackColorAll
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

    Private Sub txtSalesOrderNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesOrderNo.GotFocus
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

    Private Sub txtSalesOrderNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSalesOrderNo.KeyDown
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

    Private Sub txtSalesOrderNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesOrderNo.Leave
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

    Private Sub btnFindSalesOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesOrder.Click
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

    Private Sub txtSalesOrderNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesOrderNo.LostFocus
        Try
            If Me.txtSalesOrderNo.Enabled = True And Me.txtSalesOrderNo.ReadOnly = False Then
                Me.txtSalesOrderNo.BackColor = Me.txtPrimaryDO.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDONo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDONo.KeyDown
        Dim i As New FrmCariTransaksiDO
        Dim Jml As Boolean = False

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjDO.ApakahSudahAda(Me.txtDONo.Text)
                If Jml = True Then
                    Me.txtPrimaryDO.Text = Me.ObjDO.GetPrimary(Me.txtDONo.Text)
                    'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                    Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(Me.txtPrimaryDO.Text)
                    If PrimaryStockOut = -1 Then
                        MsgBox("This delivery order is on progress for shipment, you cannot return the good from this delivery order." + _
                        CStr(Me.objStockIn.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
                        "Information")
                        Me.txtPrimaryDO.Text = "-1"
                        Me.txtDONo.Text = ""
                        Exit Sub
                    End If
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjDO.GetCustomerPrimary(Me.txtPrimaryDO.Text)
                    Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtRemark.Focus()
                    Me.ObjDOItem.IsiDGV(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
                    Me.FormatDGVDO()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtDONo.Text <> "" Then
                        i.txtCari.Text = Me.txtDONo.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryTransaksi > 0 Then
                        'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                        Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(i.PrimaryTransaksi)
                        If PrimaryStockOut = -1 Then
                            MsgBox("This delivery order is on progress for shipment, you cannot return the good from this delivery order." + _
                            CStr(Me.objStockIn.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
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
                        Me.ObjDOItem.IsiDGV(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
                        Me.FormatDGVDO()
                    Else
                        Me.txtDONo.Clear()
                        Me.txtPrimaryDO.Text = "-1"
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.TxtKodePihakAcuan.Text = ""
                        Me.TxtPihakAcuan.Text = ""
                        Me.txtDONo.Focus()
                        Me.dgvItemDO2.Rows.Clear()
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

        Try
            If Me.txtDONo.Text <> "" Then
                If Me.ObjDO.ApakahSudahAda(Me.txtDONo.Text) = True Then
                    Me.txtPrimaryDO.Text = Me.ObjDO.GetPrimary(Me.txtDONo.Text)
                    'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                    Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(Me.txtPrimaryDO.Text)
                    If PrimaryStockOut = -1 Then
                        MsgBox("This delivery order is on progress for shipment, you cannot return the good from this delivery order." + _
                        CStr(Me.objStockIn.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
                        "Information")
                        Me.txtPrimaryDO.Text = "-1"
                        Me.txtDONo.Text = ""
                        Exit Sub
                    End If
                    Me.TxtPrimaryPihakAcuan.Text = Me.ObjDO.GetCustomerPrimary(Me.txtPrimaryDO.Text)
                    Me.TxtKodePihakAcuan.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtPihakAcuan.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryPihakAcuan.Text)
                    Me.TxtRemark.Focus()
                    Me.ObjDOItem.IsiDGV(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
                    Me.FormatDGVDO()
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.txtDONo.Text <> "" Then
                        i.txtCari.Text = Me.txtDONo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                        Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(i.PrimaryTransaksi)
                        If PrimaryStockOut = -1 Then
                            MsgBox("This delivery order is on progress for shipment, you cannot return the good from this delivery order." + _
                            CStr(Me.objStockIn.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
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
                        Me.ObjDOItem.IsiDGV(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
                        Me.FormatDGVDO()
                    Else
                        Me.txtDONo.Clear()
                        Me.txtPrimaryDO.Text = "-1"
                        Me.TxtPrimaryPihakAcuan.Text = "-1"
                        Me.TxtKodePihakAcuan.Text = ""
                        Me.TxtPihakAcuan.Text = ""
                        Me.txtDONo.Focus()
                        Me.dgvItemDO2.Rows.Clear()
                    End If
                End If
            Else
                Me.txtPrimaryDO.Text = "-1"
                Me.TxtPrimaryPihakAcuan.Text = "-1"
                Me.TxtKodePihakAcuan.Text = ""
                Me.TxtPihakAcuan.Text = ""
                Me.dgvItemDO2.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindDO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDO.Click
        Dim i As New FrmCariTransaksiDO

        Try
            i.BackColorAll = Me.BackColorAll

            If Me.txtDONo.Text <> "" Then
                i.txtCari.Text = Me.txtDONo.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                'cek apakah sudah ada stock out yang merujuk ke transaksi sales tersebut
                Dim PrimaryStockOut As Double = Me.ObjDO.GetPrimaryStockOut(i.PrimaryTransaksi)
                If PrimaryStockOut = -1 Then
                    MsgBox("This delivery order is on progress for shipment, you cannot return the good from this delivery order." + _
                    CStr(Me.objStockIn.GetNoTransaksi(PrimaryStockOut)), MsgBoxStyle.Information, _
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
                Me.ObjDOItem.IsiDGV(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
                Me.FormatDGVDO()
                'isikan data cbo so
                Me.IsiCBOSO()
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

    Private Sub FormatDGVDO()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItemDO2.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(7).Value) = False Then
                    row.Cells(7).Value = "-"
                Else
                    If CDate(row.Cells(7).Value).Date = #1/1/1950# Then
                        row.Cells(7).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(7).Value = Format(CDate(row.Cells(7).Value), "dd MMM yyyy")
                    End If
                End If

                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If IsNumeric(row.Cells(14).Value) = True Then
                    If row.Cells(14).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                If row.Cells(14).Value = "-1" Then
                    row.Cells(14).Value = "-"
                    row.Cells(15).Value = "-"
                End If

                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(6).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvItemDO2.Columns(14).Visible = False
                Me.dgvItemDO2.Columns(15).Visible = False
            Else
                Me.dgvItemDO2.Columns(14).Visible = True
                Me.dgvItemDO2.Columns(15).Visible = True
            End If
            '3. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItemDO2.Columns(6).Visible = False
            Else
                Me.dgvItemDO2.Columns(6).Visible = True
            End If
            '4. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItemDO2.Columns(7).Visible = False
            Else
                Me.dgvItemDO2.Columns(7).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub IsiCBOSO()
        Dim rowSO As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_itemRow
        Dim rowStockIn As DseTransaksiStockInSalesReturn_Balance.tbltransaksi_stockinout_itemRow
        Dim Filter As String = ""

        Try
            'lepaskan hubungan dengan cbo
            Me.cboSO_Number.DataSource = Nothing
            Me.cboSO_Number.Items.Clear()

            'isikan item DO
            Me.DTASO.FillBy_PrimaryDO(Me.TBLSO, Me.txtPrimaryDO.Text)
            Me.DTASO_StockInSalesReturn.Fill(Me.TBLSO_StockInSalesReturn, Me.TxtPrimaryStockIn.Text)

            'tambahkan dengan yang sudah disimpan di basis data
            For Each rowStockIn In Me.TBLSO_StockInSalesReturn.Rows
                Filter = "[Primary SO] =" + rowStockIn.Primary_SO.ToString + ""
                Me.TBLSO.DefaultView.RowFilter = Filter
                If Me.TBLSO.DefaultView.Count > 0 Then
                    rowSO = Me.TBLSO.DefaultView(0).Row
                    rowSO.Total_Qty += rowStockIn.Qty
                End If
            Next

            'kurangkan dengan yang ditampilkan pada DGV
            Me.KurangkanItemDGV_SO()

            'hubungkan cbo
            Me.TBLSO.DefaultView.RowFilter = ""
            Me.cboSO_Number.DataSource = Me.TBLSO.DefaultView
            Me.cboSO_Number.DisplayMember = "SO_Number"
            If Me.TBLSO.DefaultView.Count > 0 Then
                Me.cboSO_Number.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangkanItemDGV_SO()
        Dim rowDGV As DataGridViewRow
        Dim rowSO As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_itemRow
        Dim Filter As String = ""

        Try
            For Each rowDGV In Me.dgvItem.Rows
                Filter = "[Primary SO] = " + rowDGV.Cells(Me.ColPrimaryPOnoDoNo.Index).Value.ToString
                Me.TBLSO.DefaultView.RowFilter = Filter
                If Me.TBLSO.DefaultView.Count > 0 Then
                    rowSO = Me.TBLSO.DefaultView(0).Row
                    rowSO.Total_Qty -= CDbl(rowDGV.Cells(Me.ColQty.Index).Value)

                    If rowSO.Total_Qty = 0 Then
                        Me.TBLSO.Rows.Remove(rowSO)
                    End If
                Else
                    MsgBox("Sorry, we have a little problem here. Please contact your vendor if you see this message.", _
                    MsgBoxStyle.Information, "Information")
                End If
            Next

            'kembalikan filter ke kondisi 'no filter'
            Me.TBLSO.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOCodeNo()
        Dim rowSO As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item1Row
        Dim rowStockIn As DseTransaksiStockInSalesReturn_Balance.tbltransaksi_stockinout_item1Row
        Dim Filter As String = ""

        Try
            'lepaskan hubungan dengan cbo
            Me.cboCodeNo.DataSource = Nothing
            Me.cboCodeNo.Items.Clear()

            'isikan item DO
            Me.DTACodeNo.FillBy_PrimarySO(Me.TBLCodeNo, Me.txtPrimarySO.Text, Me.txtPrimaryDO.Text)
            Me.DTACodeNo_InSalesReturn.Fill(Me.TBLCodeNo_InSalesReturn, Me.TxtPrimaryStockIn.Text, Me.txtPrimarySO.Text)

            'tambahkan dengan yang sudah disimpan di basis data
            For Each rowStockIn In Me.TBLCodeNo_InSalesReturn.Rows
                Filter = "[Primary SO] = " + rowStockIn.Primary_SO.ToString + _
                " AND [Primary Product] = " + rowStockIn.Primary_Product.ToString

                Me.TBLCodeNo.DefaultView.RowFilter = Filter
                If Me.TBLCodeNo.DefaultView.Count > 0 Then
                    rowSO = Me.TBLCodeNo.DefaultView(0).Row
                    rowSO.Total_Qty += rowStockIn.Qty
                End If
            Next

            'kurangkan dengan yang ditampilkan pada DGV
            Me.KurangkanDGVCodeNo()

            'hubungkan cbo
            Me.TBLCodeNo.DefaultView.RowFilter = ""
            Me.cboCodeNo.DataSource = Me.TBLCodeNo.DefaultView
            Me.cboCodeNo.DisplayMember = "Product Code"
            If Me.TBLCodeNo.DefaultView.Count > 0 Then
                Me.cboCodeNo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangkanDGVCodeNo()
        Dim rowDGV As DataGridViewRow
        Dim rowCodeNo As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item1Row
        Dim Filter As String = ""

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = Me.txtPrimarySO.Text Then
                    Filter = "[Primary SO] = " + rowDGV.Cells(Me.ColPrimaryPOnoDoNo.Index).Value.ToString + _
                    " AND [Primary Product] =" + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString

                    Me.TBLCodeNo.DefaultView.RowFilter = Filter
                    If Me.TBLCodeNo.DefaultView.Count > 0 Then
                        rowCodeNo = Me.TBLCodeNo.DefaultView(0).Row
                        rowCodeNo.Total_Qty -= CDbl(rowDGV.Cells(Me.ColQty.Index).Value)

                        If rowCodeNo.Total_Qty = 0 Then
                            Me.TBLSO.Rows.Remove(rowCodeNo)
                        End If
                    Else
                        MsgBox("Sorry, we have a little problem here. Please contact your vendor if you see this message.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                End If
            Next

            'kembalikan filter ke kondisi 'no filter'
            Me.TBLCodeNo.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOBatchNo()
        Dim rowBatchNo As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item2Row
        Dim rowStockIn As DseTransaksiStockInSalesReturn_Balance.tbltransaksi_stockinout_item2Row
        Dim Filter As String = ""

        Try
            'lepaskan hubungan dengan cbo
            Me.CboBatchNo.DataSource = Nothing
            Me.CboBatchNo.Items.Clear()

            'isikan item DO
            Me.DTABatchNo.FillBy(Me.TBLBatchNo, Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)
            Me.DTABacthNo_InSalesReturn.Fill(Me.TBLBacthNo_InSalesReturn, Me.TxtPrimaryStockIn.Text, _
            Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)

            'tambahkan dengan yang sudah disimpan di basis data
            For Each rowStockIn In Me.TBLBacthNo_InSalesReturn.Rows
                Filter = "[Primary SO] = " + rowStockIn.Primary_SO.ToString + _
                " AND [Primary Product] = " + rowStockIn.Primary_Product.ToString + _
                " AND [Batch No] = '" + rowStockIn.Batch_No + "'"

                Me.TBLBatchNo.DefaultView.RowFilter = Filter
                If Me.TBLBatchNo.DefaultView.Count > 0 Then
                    rowBatchNo = Me.TBLBatchNo.DefaultView(0).Row
                    rowBatchNo.Qty += rowStockIn.Qty
                End If
            Next

            'kurangkan dengan yang ditampilkan pada DGV
            Me.KurangkanDGVBatchNo()

            'hubungkan cbo
            Me.TBLBatchNo.DefaultView.RowFilter = ""
            Me.CboBatchNo.DataSource = Me.TBLBatchNo.DefaultView
            Me.CboBatchNo.DisplayMember = "Batch No"
            If Me.TBLBatchNo.DefaultView.Count > 0 Then
                Me.CboBatchNo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KurangkanDGVBatchNo()
        Dim rowDGV As DataGridViewRow
        Dim rowBatch As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item2Row
        Dim Filter As String = ""

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryPOnoDoNo.Index).Value = Me.txtPrimarySO.Text And _
                rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text Then
                    Filter = "[Primary SO] = " + rowDGV.Cells(Me.ColPrimaryPOnoDoNo.Index).Value.ToString + _
                    " AND [Primary Product] =" + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + _
                    " [Batch No] ='" + rowDGV.Cells(Me.ColBatchNo.Index).Value + "'"

                    Me.TBLBatchNo.DefaultView.RowFilter = Filter
                    If Me.TBLBatchNo.DefaultView.Count > 0 Then
                        rowBatch = Me.TBLBatchNo.DefaultView(0).Row
                        rowBatch.Qty -= CDbl(rowDGV.Cells(Me.ColQty.Index).Value)

                        If rowBatch.Qty = 0 Then
                            Me.TBLSO.Rows.Remove(rowBatch)
                        End If
                    Else
                        MsgBox("Sorry, we have a little problem here. Please contact your vendor if you see this message.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                End If
            Next

            'kembalikan filter ke kondisi 'no filter'
            Me.TBLBatchNo.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOLocation()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOSecQty()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboSO_Number_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSO_Number.SelectedIndexChanged
        Dim rowSO As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_itemRow

        Try
            If Me.cboSO_Number.SelectedIndex = -1 Then
                Me.cboCodeNo.DataSource = Nothing
                Me.cboCodeNo.Items.Clear()
                Me.txtPrimarySO.Text = "-1"
            Else
                rowSO = Me.TBLSO.DefaultView(Me.cboSO_Number.SelectedIndex).Row
                Me.txtPrimarySO.Text = rowSO.Primary_SO
                Me.IsiCBOCodeNo()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCodeNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedIndexChanged
        Dim rowCodeNo As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item1Row

        Try
            If Me.cboCodeNo.SelectedIndex = -1 Then
                Me.cboCodeNo.DataSource = Nothing
                Me.cboCodeNo.Items.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
            Else
                rowCodeNo = Me.TBLCodeNo.DefaultView(Me.cboCodeNo.SelectedIndex).Row
                Me.TxtPrimaryCodeNo.Text = rowCodeNo.Primary_Product
                If Me.chkBatchNoControlled.Checked = True Then
                    Me.IsiCBOBatchNo()
                ElseIf Me.chkLocationControlled.Checked = True Then
                    Me.IsiCBOLocation()
                ElseIf Me.chkUseSecondaryUnit.Checked = True Then
                    Me.IsiCBOSecQty()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboBatchNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboBatchNo.SelectedIndexChanged
        Dim rowBatch As DseTransaksiStockInSalesReturn_Balance.tbltrans_do_item2Row

        Try
            If Me.CboBatchNo.SelectedIndex = -1 Then
                Me.CboBatchNo.DataSource = Nothing
                Me.CboBatchNo.Items.Clear()
                Me.txtBatchNo.Text = ""
            Else
                rowBatch = Me.TBLCodeNo.DefaultView(Me.CboBatchNo.SelectedIndex).Row
                Me.txtBatchNo.Text = Me.CboBatchNo.Text
                If Me.chkExpiryDateControlled.Checked = True Then
                    Me.DtpExpired.Value = rowBatch.Expiry_Date
                End If

                If Me.chkLocationControlled.Checked = True Then
                    Me.IsiCBOLocation()
                ElseIf Me.chkUseSecondaryUnit.Checked = True Then
                    Me.IsiCBOSecQty()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.Leave
        Try
            Me.txtBatchNo.Text = Me.txtBatchNo.Text.Trim
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class