Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiBookingInv
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0

    Dim baru As Boolean
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSOItem As New ClsTransaksiSalesOrderItem
    Dim objDOItem As New ClsTransaksiDOItem
    Dim ObjBookingInv As New ClsTransaksiBookingInv
    Dim ObjBookingInvItem As New ClsTransaksiBookingInvItem
    Dim ObjBookingGITItem As New ClsTransaksiBookingGITitem
    Dim ObjLicense As New ClsLicense

    Dim IsInventoryInstalled As Boolean = False

    Dim TBL As New DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable
    Dim TBL2 As New DseTransaksiBookingInv.DataTable22DataTable
    Dim TBL5 As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL7 As New dseMasterProduct.tblmaster_productqty3DataTable
    Dim TBL8 As New DseTransaksiBookingInv.tbltrans_do_item2DataTable
    Dim TBL9 As New DseTransaksiBookingInv.tbltrans_so_itemDataTable

    Dim PersenPajakAcuan, PrimaryCurrencySO As Decimal
    Dim JenisFakturPajakAcuan As ClsTransaksiSalesOrder.EnumJenisFakturPajak = ClsTransaksiSalesOrder.EnumJenisFakturPajak.FakturPajakSederhana
    Dim SamaSO As Boolean = False
    Dim SamaCodeNo As Boolean = False
    Dim buffer As String = ""
    Dim bufferProduct As String = ""
    Dim bufferBatchNo As String = ""
    Dim bufferPrimaryLocation As String = ""

    'elemen yang berkaitan dengan perhitungan free qty, available to book, dll
    Dim DTA_Prod_qty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

    'eelemen yang berkaitan dengan pengisian cbo secara otomatis-ini dicoding berkaitan dengan penambahan kontrol 
    'lokasi pada inventory
    Dim TBL_1 As New Core_Module.dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL_2 As New DseTransaksiBookingInv.tbltrans_booking_invitem3DataTable
    Dim TBL_3 As New Core_Module.dseMasterProduct.stock_location_viewDataTable
    Dim TBL_4 As New DseTransaksiBookingInv.DataTable5DataTable
    Dim TBL_5 As New Core_Module.dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL_6 As New DseTransaksiBookingInv.tbltrans_booking_invitem4DataTable
    Dim TBL3 As New DseTransaksiSO.laporan_sales_orderbalance_view9dDataTable
    Dim TBL4 As New DseTransaksiBookingInv.DataTable22DataTable

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryBooking.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.txtPrimaryBooking.BackColor, False)
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
            Me.dgvItem.Columns(Me.ColBatch.Index).ReadOnly = True
            Me.DtpExpired.Enabled = False
            Me.llbDetailStock.Enabled = True

            If Me.dgvItem.RowCount > 0 Then
                Me.TxtSONumber.Enabled = False
                Me.BtnFindSONumber.Enabled = False
            End If

            If Me.TxtPrimaryCustomer.Text <> "-1" And Me.TxtPrimaryCustomer.Text <> "" Then
                Me.SetGboItem()
            End If
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
            Me.llbDetailStock.Enabled = False

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
            Me.TxtPrimaryCustomer.Text = "-1"
            Me.TxtPrimarySONumber.Text = "-1"
            Me.txtPrimaryBooking.Text = "-1"
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
            Me.txtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnits.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyperBatch.Text = "0"
            Me.txtSecondaryQty.Text = "-1"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.lblSecondaryUnits.Visible = False
            Me.chkLocationControlled.Checked = False
            Me.lblUnits.Left = Me.lblSecondaryUnits.Left
            Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
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
            Me.cboCodeNo.DataSource = Nothing
            Me.lblLocation.Visible = False
            Me.lblWarehouse.Visible = False
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Visible = True
            Me.cboLocation.Items.Clear()
            Me.TxtWarehouse.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim DefPrec As Decimal = 0
        Try
            'data item salah bila:
            '1. Jika product code tidak ada
            If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                MsgBox("Please specify code no", MsgBoxStyle.Information, "Information")
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
            If Me.chkExpiryDateControlled.Checked = True Then
                If Me.DtpExpired.Value.Date < Me.DtpCreatedDate.Value.Date Then
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
                    If row.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text And Me.cboBatchNo.Text <> "" And row.Cells("ColPrimaryCodeNo").Value = _
                    Me.txtPrimaryCodeNo.Text Then
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

            '9. Bila setelah diinput jumlah sales order balance menjadi minus
            Dim JmlObjLama, JmlSOBBaru, JmlDGV, JmlSOB, JmlPendingDO, JmlBookingInv, JmlBookingGIT As Decimal
            JmlDGV = Me.GetJmlDGV(Me.txtPrimaryCodeNo.Text, Me.TxtPrimarySONumber.Text)
            JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.TxtPrimarySONumber.Text, Me.txtPrimaryCodeNo.Text)
            JmlObjLama = Me.ObjBookingInvItem.GetJmlByPrimaryProductNBooking(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text)
            'JmlPendingDO = Me.objDOItem.GetJmlPendingDO(Me.txtPrimaryCodeNo.Text, Me.TxtPrimarySONumber.Text)
            JmlBookingInv = Me.ObjBookingInvItem.GetJmlByPrimarySO_n_PrimaryProduct(Me.TxtPrimarySONumber.Text, Me.txtPrimaryCodeNo.Text)
            JmlBookingGIT = Me.ObjBookingGITItem.GetQtyByPrimarySO_n_Product_Multi(Me.TxtPrimarySONumber.Text, Me.txtPrimaryCodeNo.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlSOBBaru = JmlSOB + JmlObjLama - JmlDGV - JmlPendingDO - _
            CDec(Me.TxtQtyperBatch.Text) - JmlBookingInv - JmlBookingGIT
            If JmlSOBBaru < 0 Then
                DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
                If CDec(Me.txtAvailable.Text) < CDec(Me.txtQtyOrderBalance.Text) Then
                    MsgBox("Maximum value you can enter is " + Me.txtAvailable.Text + " " + _
                    CStr(Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                    "Information")
                Else
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlSOB + JmlObjLama - _
                    JmlDGV - JmlPendingDO - JmlBookingInv - JmlBookingGIT), DefPrec) + " " + _
                    CStr(Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                    "Information")
                End If
                Me.TxtQtyperBatch.Focus()
                Return False
            End If

            '9.b. Cek stock pada tanggal yang bersangkutan
            'msgbox("belum"), harus dipertimbangkan kembali

            '9.c Jika inventory diinstall cek kembali berkaitan dengan stock digudang,
            '    jangan sampe malahan bikin minus
            If CDec(Me.txtAvailable.Text) < CDec(Me.TxtQtyperBatch.Text) Then
                MsgBox("Maximum value you can enter is " + Me.txtAvailable.Text + " " + _
                CStr(Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatch.Focus()
                Return False
            End If

            '10. Dalam 1 do bisa ada beberapa SO dengan syarat:
            If Me.dgvItem.RowCount > 0 Then
                Dim mTaxNew, mprimaryCurrency As Decimal
                Dim mJenisFaktur As ClsTransaksiSalesOrder.EnumJenisFakturPajak
                Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.TxtPrimarySONumber.Text, _
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

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Decimal, ByVal vBatchNo As String, _
    ByVal vSecQty As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    If row.Cells("ColPrimaryCodeNo").Value = vPrimaryCodeNo And row.Cells(Me.ColBatch.Index).Value = vBatchNo _
                    And row.Cells(Me.ColSecondaryQty.Index).Value = vSecQty Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                    End If
                Else
                    If row.Cells("ColPrimaryCodeNo").Value = vPrimaryCodeNo And row.Cells(Me.ColBatch.Index).Value = vBatchNo Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV4(ByVal vPrimaryCodeNo As Decimal, ByVal vBatchNo As String, _
    ByVal vSecQty As Decimal, ByVal vPrimaryLocation As Decimal) As Decimal

        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    If row.Cells("ColPrimaryCodeNo").Value = vPrimaryCodeNo And row.Cells(Me.ColBatch.Index).Value = vBatchNo _
                    And row.Cells(Me.ColSecondaryQty.Index).Value = vSecQty And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                    End If
                Else
                    If row.Cells("ColPrimaryCodeNo").Value = vPrimaryCodeNo And row.Cells(Me.ColBatch.Index).Value = vBatchNo _
                    And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                        'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                        'date yang salah karena sudah dicek terlebih dahulu
                        Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml(ByVal vPrimaryCodeNo As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells("ColPrimaryCodeNo").Value = vPrimaryCodeNo Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Decimal, ByVal vPrimarySO As Decimal) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvItem.Rows
                If row.Cells("ColPrimaryCodeNo").Value = vPrimaryCodeNo And CDec(Me.TxtPrimarySONumber.Text) = vPrimarySO Then
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    Jml = Jml + CDec(row.Cells(Me.ColQty.Index).Value)
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
        Try
            'cek data yang harus diisi
            KoleksiNamaData.Add("sales order no")
            KoleksiTextBox.Add(Me.TxtSONumber)
            KoleksiNamaData.Add("customer code")
            KoleksiTextBox.Add(Me.TxtCustomerCode)
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
            If Me.baru = False Then
                If Me.ObjBookingInv.TanggalTransaksi.Month = Me.DtpCreatedDate.Value.Month And _
                Me.ObjBookingInv.TanggalTransaksi.Year = Me.DtpCreatedDate.Value.Year Then
                    Me.TxtBookingNo.Text = Me.ObjBookingInv.BookingNo
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.DtpCreatedDate.Value = Me.ObjBookingInv.TanggalTransaksi
                        Me.TxtBookingNo.Text = Me.ObjBookingInv.BookingNo
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
            With Me.ObjBookingInv
                .BookingNo = Me.TxtBookingNo.Text
                .TanggalTransaksi = Me.DtpCreatedDate.Value
                .primaryCustomer = Me.TxtPrimaryCustomer.Text
                .primarySO = Me.TxtPrimarySONumber.Text
                .Remarks = Me.TxtRemark.Text
                .ValidUntil = CDec(Me.cboValidUntil.SelectedIndex) + 1
                .softwaregenerate = 0
                .primaryBookingGIT = -1
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjBookingInv
                Me.TxtBookingNo.Text = .BookingNo
                Me.DtpCreatedDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.txtPrimaryBooking.Text = .primary
                Me.TxtPrimaryCustomer.Text = .primaryCustomer
                Me.TxtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                Me.TxtPrimarySONumber.Text = .primarySO
                Me.TxtSONumber.Text = ObjSO.GetNoSO(Me.TxtPrimarySONumber.Text)
                Me.cboValidUntil.SelectedIndex = CDec(.ValidUntil) - 1
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    row.Cells(Me.ColSecondaryQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecondaryQty.Index).Value)
                End If
            Next
            'set nilai acuan untuk pengecekan SO
            Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.TxtPrimarySONumber.Text, Me.PrimaryCurrencySO, _
            Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjLS As New clsMasterListedSoftware
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'cek apakah inventory juga sudah diinstall
            Me.IsInventoryInstalled = ObjLS.IsInventoryInstalled
            If Me.IsInventoryInstalled = True Then
                Me.cboSecondaryQty.DropDownStyle = ComboBoxStyle.DropDownList
                Me.cboBatchNo.DropDownStyle = ComboBoxStyle.DropDownList
            End If
            Me.cboValidUntil.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjLS = Nothing
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
                        If Me.ObjBookingInv.SimpanBaru(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
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
                        If Me.ObjBookingInv.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.KosongkanItem()
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
            If Me.TxtRemark.Text = "" And Me.dgvItem.RowCount = 0 And Me.DtpCreatedDate.Value.Date = Today And _
            Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
                Me.KosongkanItem()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjBookingInv.CariDetail(Me.txtPrimaryBooking.Text, Me.dgvItem, Me)
                        Me.SalinKeField()
                        Me.SetJmlBO_N_FreeQty()
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
            Me.KosongkanItem()
            Me.DtpCreatedDate.Focus()
            Me.ObjBookingInv.ResetObject()
            Me.TxtBookingNo.Text = Me.ObjBookingInv.GetNoBaru(Me.DtpCreatedDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiBookingInv
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()

            If Not ij.PrimaryTransaksi = -1 Then
                If Me.ObjBookingInv.CariDetail(ij.PrimaryTransaksi, Me.dgvItem, Me) = True Then
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
        Try
            If Me.TxtBookingNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtBookingNo.Text <> "" Then
                If Me.ObjBookingInv.ApakahBolehEdit(Me.txtPrimaryBooking.Text, UserCode, True, True) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.IsiCboProduct(Me.TxtPrimarySONumber.Text)
                    Me.dgvItem.Focus()
                    Me.TBL9.Clear()
                    Dim DTASO As New DseTransaksiBookingInvTableAdapters.tbltrans_so_itemTableAdapter
                    DTASO.FillBy_PrimarySO(TBL9, Me.TxtPrimarySONumber.Text)
                    DTASO.Dispose()

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
            If Me.TxtBookingNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtBookingNo.Text <> "" Then
                If Me.ObjBookingInv.ApakahBolehHapus(Me.txtPrimaryBooking.Text, True, UserCode, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjBookingInv.Hapus(Me.txtPrimaryBooking.Text, UserCode, True) = True Then
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
                            If Me.ObjBookingInv.SimpanBaru(Me.dgvItem, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                '                                Me.SalinKeField()
                                '                                Me.NonAktifkanDetail()
                                '                                Me.btnFind.Focus()
                                '                                Me.Close()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjBookingInv.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                '                                Me.SalinKeField()
                                '                                Me.NonAktifkanDetail()
                                '                                Me.btnEdit.Focus()
                                '         Me.Close()
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

    Private Sub IsikanTBLSO()
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiBookingInv.DataTable22Row
            For Each row In Me.dgvItem.Rows
                Dim ketemu As Boolean = False
                For Each row1 In Me.TBL2.Rows
                    If row1.Primary_Customer = Me.TxtPrimaryCustomer.Text And _
                    row1.Primary_Product = row.Cells("ColPrimaryCodeNo").Value And _
                    row1.Primary_SO = CDec(Me.TxtPrimarySONumber.Text) Then
                        'kalo ketemu ubah data order balance
                        row1.Qty = row1.Qty + CDec(row.Cells(Me.ColQty.Index).Value)
                        ketemu = True
                        Exit For
                    End If
                Next
                'kalo ga ketemu tambahkan baris baru
                If ketemu = False Then
                    Dim ExpiryDate As Date = #1/1/1950#
                    If CStr(row.Cells(Me.ColExpiredDate.Index).Value) <> "-" Then
                        ExpiryDate = row.Cells(Me.ColExpiredDate.Index).Value
                    End If
                    Dim SecQty As Decimal = -1
                    If CStr(row.Cells(Me.ColSecondaryQty.Index).Value) <> "-" Then
                        SecQty = row.Cells(Me.ColSecondaryQty.Index).Value
                    End If
                    Me.TBL2.Rows.Add(row.Cells("ColPrimaryCodeNo").Value, CDec(Me.TxtPrimarySONumber.Text), row.Cells(Me.ColCodeNo.Index).Value, _
                    Me.TxtSONumber.Text, Me.TxtBookingNo.Text, Me.DtpCreatedDate.Value, Me.txtPrimaryBooking.Text, row.Cells(Me.ColQty.Index).Value, _
                    Me.txtPrimaryUnits.Text, row.Cells(Me.ColBatch.Index).Value, ExpiryDate, row.Cells(Me.ColRemark.Index).Value, _
                    SecQty, row.Cells(Me.ColPrimarySecondaryUnits.Index).Value, row.Cells(Me.ColValidUntil.Index).Value, _
                    row.Cells(Me.ColDescription.Index).Value, Me.TxtPrimaryCustomer.Text)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboProduct(ByVal vPrimarySO As Decimal)
        Try
            TBL_1.Rows.Clear()
            TBL_2.Rows.Clear()
            TBL_3.Rows.Clear()
            TBL_4.Rows.Clear()
            TBL_5.Rows.Clear()
            TBL_6.Rows.Clear()

            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'isikan dari basis data-persedian terakhir
            Me.cboCodeNo.DataSource = Nothing
            Me.ObjSOItem.IsikanTBLProduct2_Default(Me.TxtPrimarySONumber.Text, Me.TBL3)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiBookingInv.DataTable22Row
            Dim row2 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjBookingInvItem.IsikanTBLProductBySO(Me.txtPrimaryBooking.Text, Me.TBL4, Me.TxtPrimarySONumber.Text)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    For Each row2 In Me.TBL3.Rows
                        If row1.Primary_Customer = row2.Primary_Customer _
                        And row1.Primary_Product = row2.Primary_Product And _
                        row2.Primary_SO = row1.Primary_SO Then
                            row2.Order_Balance = row2.Order_Balance + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL3.Rows.Add(row1.Primary_SO, row1.Primary_Customer, row1.Primary_Product, _
                        row1.Product_Code, row1.Qty, row1.So_No)
                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsiTBLProduct(Me.TxtPrimarySONumber.Text)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                For Each row2 In Me.TBL3.Rows
                    If row1.Primary_Customer = row2.Primary_Customer _
                    And row1.Primary_Product = row2.Primary_Product And _
                    row2.Primary_SO = row1.Primary_SO Then
                        row2.Order_Balance = row2.Order_Balance - row1.Qty
                        If row2.Order_Balance = 0 Then
                            Me.TBL3.Rows.Remove(row2)
                        End If
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

            Me.TBL3.DefaultView.Sort = "[Product Code] ASC"
            Me.cboCodeNo.DataSource = Me.TBL3.DefaultView
            Me.cboCodeNo.DisplayMember = "Product Code"
            Me.cboCodeNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub IsiTBLProduct(ByVal vPrimarySO As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiBookingInv.DataTable22Row
            For Each row In Me.dgvItem.Rows
                If CDec(Me.TxtPrimarySONumber.Text) = vPrimarySO Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Primary_Customer = Me.TxtPrimaryCustomer.Text And _
                        row1.Primary_Product = row.Cells("ColPrimaryCodeNo").Value And _
                        row1.Primary_SO = CDec(Me.TxtPrimarySONumber.Text) Then
                            'kalo ketemu ubah data order balance
                            row1.Qty = row1.Qty + CDec(row.Cells(Me.ColQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ExpiryDate As Date = #1/1/1950#
                        If CStr(row.Cells(Me.ColExpiredDate.Index).Value) <> "-" Then
                            ExpiryDate = row.Cells(Me.ColExpiredDate.Index).Value
                        End If
                        Dim SecQty As Decimal = -1
                        If CStr(row.Cells(Me.ColSecondaryQty.Index).Value) <> "-" Then
                            SecQty = row.Cells(Me.ColSecondaryQty.Index).Value
                        End If
                        Dim BAtchNo As String = ""
                        If String.IsNullOrEmpty(row.Cells(Me.ColBatch.Index).Value) = False Then
                            BAtchNo = row.Cells(Me.ColBatch.Index).Value
                        End If
                        Me.TBL4.Rows.Add(row.Cells("ColPrimaryCodeNo").Value, CDec(Me.TxtPrimarySONumber.Text), row.Cells(Me.ColCodeNo.Index).Value, _
                        Me.TxtSONumber.Text, Me.TxtBookingNo.Text, Me.DtpCreatedDate.Value, Me.txtPrimaryBooking.Text, row.Cells(Me.ColQty.Index).Value, _
                        Me.txtPrimaryUnits.Text, BAtchNo, ExpiryDate, row.Cells(Me.ColRemark.Index).Value, _
                        SecQty, row.Cells(Me.ColPrimarySecondaryUnits.Index).Value, row.Cells(Me.ColValidUntil.Index).Value, _
                        row.Cells(Me.ColDescription.Index).Value, Me.TxtPrimaryCustomer.Text)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Decimal)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiBookingInv.tbltrans_booking_invitem3Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL_2.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColBatch.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDec(row.Cells(Me.ColQty.Index).Value)
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
                        Me.TBL_2.Rows.Add(CStr(row.Cells(Me.ColBatch.Index).Value), _
                        CDec(row.Cells(Me.ColQty.Index).Value), ED)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboSecondary(ByVal vPrimaryProduct As Decimal, ByVal vPrimaryTransaksi As Decimal, _
    ByVal vBatchNo As String, ByVal vPrimaryLokasi As Decimal)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboSecondaryUnit(vPrimaryProduct, Me.cboSecondaryQty, _
            Me.TBL_5, vBatchNo, vPrimaryLokasi)
            Dim row1 As DseTransaksiBookingInv.tbltrans_booking_invitem4Row
            Dim row2 As dseMasterProduct.tblmaster_productqty2Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjBookingInvItem.IsiTBLSecondaryQty(vPrimaryTransaksi, vPrimaryProduct, _
                 vBatchNo, vPrimaryLokasi, Me.TBL_6)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL_6.Rows.Count > 0
                    row1 = Me.TBL_6.Rows(0)
                    For Each row2 In Me.TBL_5.Rows
                        If row1.Secondary_Qty = row2.Qty_Secondary Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL_5.Rows.Add(row1.Secondary_Qty, row1.Qty)
                    End If
                    Me.TBL_6.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLSecondary(vPrimaryProduct, vPrimaryLokasi, vBatchNo)
            While Me.TBL_6.Rows.Count > 0
                row1 = Me.TBL_6.Rows(0)
                For Each row2 In Me.TBL_5.Rows
                    If row1.Secondary_Qty = row2.Qty_Secondary Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL_5.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL_5.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL_6.Rows.Remove(row1)
            End While

            Me.TBL_5.DefaultView.Sort = "[Qty Secondary] ASC"
            Me.cboSecondaryQty.DataSource = Me.TBL_5.DefaultView
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

    Private Sub IsikanTBLSecondary(ByVal vPrimaryProduct As Decimal, ByVal vPrimaryLocation As Decimal, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiBookingInv.tbltrans_booking_invitem4Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                CStr(row.Cells(Me.ColBatch.Index).Value) = vBatchNo And _
                row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL_6.Rows
                        If row1.Secondary_Qty = CDec(row.Cells(Me.ColSecondaryQty.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDec(row.Cells(Me.ColQty.Index).Value)
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
                        Me.TBL_6.Rows.Add(CDec(row.Cells(Me.ColSecondaryQty.Index).Value), CStr(row.Cells(Me.ColQty.Index).Value), _
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
            Me.cboBatchNo.DataSource = Nothing
            Me.cboSecondaryQty.DataSource = Nothing
            Me.cboLocation.DataSource = Nothing
            Me.cboSecondaryQty.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex2(ByVal vPrimaryProduct As Decimal, ByVal BatchNo As String, _
    ByVal SecondaryQty As Decimal, ByVal ExpiryDate As Date, ByVal vPrimarySO As Decimal, _
    ByVal vPrimaryLocation As Decimal) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And CDec(Me.TxtPrimarySONumber.Text) = vPrimarySO And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 _
                        And vPrimaryLocation = row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And CDec(Me.TxtPrimarySONumber.Text) = vPrimarySO And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty _
                        And ExpiryDate = #1/1/1950# And vPrimaryLocation = row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And CDec(Me.TxtPrimarySONumber.Text) = vPrimarySO And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) _
                        And vPrimaryLocation = row.Cells(Me.ColPrimaryLocation.Index).Value Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And CDec(Me.TxtPrimarySONumber.Text) = vPrimarySO And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) _
                        And vPrimaryLocation = row.Cells(Me.ColPrimaryLocation.Index).Value Then
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
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty _
                        And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells("ColPrimaryCodeNo").Value = vPrimaryProduct And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And _
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
        Dim rowDGV As DataGridViewRow

        Try
            If Me.ApakahDataItemBenar = True Then
                Me.TxtSONumber.Enabled = False
                Me.BtnFindSONumber.Enabled = False

                Dim SecQty As Decimal = -1
                If Me.txtSecondaryQty.Text = "" Or Me.txtSecondaryQty.Text = "-1" Then
                    SecQty = -1
                Else
                    SecQty = Me.txtSecondaryQty.Text
                End If
                Dim Index As Integer = Me.GetIndex2(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text, _
                SecQty, Me.DtpExpired.Value, Me.TxtPrimarySONumber.Text, Me.txtPrimaryLocation.Text)

                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQty.Index).Value = CDec(row.Cells(Me.ColQty.Index).Value) + CDec(Me.TxtQtyperBatch.Text)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If
                    Dim BatchNo As String = ""
                    If Me.cboBatchNo.Items.Count > 0 Then
                        Dim DV As DataRowView
                        DV = Me.cboBatchNo.SelectedItem
                        BatchNo = DV(0)
                    End If

                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = "0"
                    rowDGV.Cells(Me.ColCodeNo.Index).Value = Me.cboCodeNo.Text
                    rowDGV.Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatch.Text
                    rowDGV.Cells(Me.ColUnit.Index).Value = Me.lblUnits.Text
                    rowDGV.Cells(Me.ColSecondaryQty.Index).Value = SecQty
                    rowDGV.Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                    rowDGV.Cells(Me.ColBatch.Index).Value = BatchNo
                    rowDGV.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                    rowDGV.Cells(Me.ColValidUntil.Index).Value = Me.dtpValid1.Value
                    rowDGV.Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                    rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.txtPrimaryCodeNo.Text
                    rowDGV.Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnits.Text
                    rowDGV.Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                    rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                    rowDGV.Cells(Me.ColLocation.Index).Value = Me.cboLocation.Text
                    rowDGV.Cells(Me.ColWarehouse.Index).Value = Me.TxtWarehouse.Text

                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If Me.dgvItem.RowCount = 1 Then
                        'set nilai acuan untuk pengecekan SO
                        Me.ObjSO.GetPajakNJenisFakturNPrimaryCurrency(Me.TxtPrimarySONumber.Text, Me.PrimaryCurrencySO, _
                        Me.JenisFakturPajakAcuan, Me.PersenPajakAcuan)
                    End If
                End If

                'kendalikan focus item selanjutnya
                Dim JmlObjLama, JmlSOBBaru, JmlDGV, JmlSOB As Decimal
                JmlDGV = Me.GetJmlDGV(Me.txtPrimaryCodeNo.Text, Me.TxtPrimarySONumber.Text)
                JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.TxtPrimarySONumber.Text, Me.txtPrimaryCodeNo.Text)
                JmlObjLama = Me.ObjBookingInvItem.GetJmlByPrimaryProductNBooking(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text)

                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlSOBBaru = JmlSOB + JmlObjLama - JmlDGV
                Me.TxtQtyperBatch.Text = "0"
                Me.TxtRemarkItem.Text = ""

                If JmlSOBBaru <= 0 Then
                    'remove row, pada code no ybs
                    Dim rowView As DataRowView = Me.cboCodeNo.SelectedItem
                    Dim row As DseTransaksiSO.laporan_sales_orderbalance_view9dRow = rowView.Row
                    Me.TBL3.Rows.Remove(row)
                    Me.cboCodeNo_SelectedValueChanged(sender, e)
                    Me.cboCodeNo.Focus()
                    If Me.cboCodeNo.Items.Count = 0 Then
                        Me.SetGboItem()
                        '                        Me.cboSalesOrderNo.Focus()
                    End If
                Else
                    'focus pada cbo code no
                    Me.cboCodeNo.Focus()
                    'hitung kembali order balance dan sisa free qty
                    Me.SetJmlBO_N_FreeQty()
                    'reload data bacth no/secondary
                    If Me.chkBatchNoControlled.Checked = True Then
                        Me.IsiCboBatchNo(Me.txtPrimaryCodeNo.Text, Me.TxtPrimarySONumber.Text)
                    ElseIf Me.chkLocationControlled.Checked = True Then
                        Me.IsiCboLokasi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, Me.cboBatchNo.Text)
                    ElseIf Me.chkUseSecondaryUnit.Checked = True Then
                        Me.IsiCboSecondary(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                    End If
                End If

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
        Dim JmlObjLama, JmlSOBBaru, JmlDGV, JmlSOB, JmlPendingDO, JmlBookingInv, JmlBookingGIT As Decimal
        Dim mDec As Integer = 0

        Try
            mDec = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            'hitung jml yang bisa dibooking dilihat dari sisi jumlah sales order
            JmlDGV = Me.GetJmlDGV(Me.txtPrimaryCodeNo.Text, Me.TxtPrimarySONumber.Text)
            JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.TxtPrimarySONumber.Text, Me.txtPrimaryCodeNo.Text)
            JmlObjLama = Me.ObjBookingInvItem.GetJmlByPrimaryProductNBooking(Me.txtPrimaryCodeNo.Text, _
            Me.txtPrimaryBooking.Text)
            'JmlPendingDO = Me.objDOItem.GetJmlPendingDO(Me.txtPrimaryCodeNo.Text, _
            'Me.TxtPrimarySONumber.Text)
            JmlBookingInv = Me.ObjBookingInvItem.GetJmlByPrimarySO_n_PrimaryProduct(Me.TxtPrimarySONumber.Text, _
            Me.txtPrimaryCodeNo.Text)
            JmlBookingGIT = Me.ObjBookingGITItem.GetQtyByPrimarySO_n_Product_Multi(Me.TxtPrimarySONumber.Text, _
            Me.txtPrimaryCodeNo.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlSOBBaru = JmlSOB + JmlObjLama - JmlDGV - JmlPendingDO - CDec(Me.TxtQtyperBatch.Text) _
            - JmlBookingInv - JmlBookingGIT
            Me.txtQtyOrderBalance.Text = Me.ObjCD.DigitGrouping(JmlSOBBaru, mDec)

            'hitung nilai available dilihat dari inventory
            Me.txtQtyOnHand.Text = Me.ObjCD.DigitGrouping(Me.ObjProduct.GetJml2(Me.txtPrimaryCodeNo.Text, Me.DTA_Prod_qty, _
            Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text), mDec)
            JmlDGV = Me.GetJmlDGV4(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text, -1, Me.txtPrimaryLocation.Text)
            JmlObjLama = Me.ObjBookingInvItem.GetJml(Me.txtPrimaryBooking.Text, Me.TxtPrimarySONumber.Text, _
            Me.txtSecondaryQty.Text, Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
            JmlBookingInv = Me.ObjBookingInvItem.GetJmlOtherBookingInv2(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, _
            Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text, Me.txtSecondaryQty.Text)
            Me.TxtBooked.Text = Me.ObjCD.DigitGrouping(JmlBookingInv, mDec)
            Me.txtPendingDO.Text = Me.ObjCD.DigitGrouping(Me.objDOItem.GetJmlPendingDO3_Default(Me.txtPrimaryCodeNo.Text, _
            Me.cboBatchNo.Text, Me.txtSecondaryQty.Text), mDec)
            Me.txtAvailable.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtQtyOnHand.Text) - _
            CDec(Me.txtPendingDO.Text) - JmlDGV + JmlObjLama - JmlBookingInv, mDec)
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
                '0. Valid until
                row.Cells(Me.ColValidUntil.Index).Value = Format(CDate(row.Cells(Me.ColValidUntil.Index).Value), _
                "dd MMM yyyy")
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

                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                    If row.Cells(Me.ColSecondaryQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                If row.Cells(Me.ColSecondaryQty.Index).Value = "-1" Then
                    row.Cells(Me.ColSecondaryQty.Index).Value = "-"
                    row.Cells(Me.ColSecondaryUnit.Index).Value = "-"
                End If

                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColBatch.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If

                '1.e tandai bila ada location yang tidak kosong
                If row.Cells(Me.ColPrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.ColLocation.Index).Value = "-"
                    row.Cells(Me.ColWarehouse.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
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
            '3. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColBatch.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColBatch.Index).Visible = True
            End If
            '4. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiredDate.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiredDate.Index).Visible = True
            End If

            'jika semua data tidak ada location, maka kolom lokasi dan warehouse tidak perlu ditampilkan
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = False
                Me.dgvItem.Columns(Me.ColWarehouse.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocation.Index).Visible = True
                Me.dgvItem.Columns(Me.ColWarehouse.Index).Visible = True
            End If

            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
            Next
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

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColSecondaryQty.Index Or e.ColumnIndex = Me.ColBatch.Index _
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
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItem.Item("ColPrimaryCodeNo", e.RowIndex).Value) = False Then
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
            Dim SecondaryQty, JmlDGV As Decimal
            Dim BatchNo As String = ""
            If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) = False Then
                BatchNo = Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value
            End If
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            If e.ColumnIndex = Me.ColQty.Index Then
                If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Try
                End If
                If ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Try
                End If
            End If

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

                        If e.ColumnIndex = Me.ColQty.Index Then
                            Me.EndEditQty(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColSecondaryQty.Index Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColBatch.Index Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = 16 Or e.ColumnIndex = Me.ColBatch.Index Then
                        Dim BufferExpiryDate As String = ""
                        If Me.buffer <> "" Then
                            BufferExpiryDate = Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value
                        End If
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value = CDate(#1/1/1950#)
                        If e.ColumnIndex = Me.ColBatch.Index Then
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
                        Me.FormatDGV()
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
        Dim dummy As Decimal = 0
        Dim SecQty As Decimal = -1

        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'cek apakah sales order balance menjadi minus setelah perubahan data
                Dim JmlPendingDO, JmlFree, JmlAvailable, JmlObjLama, JmlSOBBaru, JmlSOB, JmlSendiri, JmlBookingInv, JmlBookingGIT As Decimal
                'JmlPendingDO = Me.objDOItem.GetJmlPendingDO(.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.TxtPrimarySONumber.Text)
                JmlDGV = Me.GetJmlDGV(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                Me.TxtPrimarySONumber.Text)
                JmlSOB = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.TxtPrimarySONumber.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlObjLama = Me.ObjBookingInvItem.GetJml(Me.txtPrimaryBooking.Text, Me.TxtPrimarySONumber.Text, _
                .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlSendiri = .Cells(Me.ColQty.Index).Value
                JmlBookingInv = Me.ObjBookingInvItem.GetJmlByPrimarySO_n_PrimaryProduct(Me.TxtPrimarySONumber.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                JmlBookingGIT = Me.ObjBookingGITItem.GetQtyByPrimarySO_n_Product_Multi(Me.TxtPrimarySONumber.Text, .Cells(Me.ColPrimaryCodeNo.Index).Value)
                'perhitungan menggunakan nilai jumlah dgv yang baru
                JmlSOBBaru = JmlSOB + JmlObjLama - JmlDGV - JmlBookingInv - JmlBookingGIT - JmlPendingDO

                If JmlSOBBaru < 0 Then
                    dummy = JmlSOB + JmlObjLama - JmlBookingInv - JmlBookingGIT - JmlPendingDO
                    JmlFree = Me.ObjProduct.GetJmlByPrimary(.Cells(Me.ColPrimaryCodeNo.Index).Value) _
                    - Me.ObjBookingInvItem.GetJml(.Cells(Me.ColPrimaryCodeNo.Index).Value) + _
                    Me.ObjBookingInvItem.GetJml(Me.txtPrimaryBooking.Text, Me.TxtPrimarySONumber.Text, _
                    .Cells(Me.ColPrimaryCodeNo.Index).Value) - Me.GetJml(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                    'JmlPendingDO = Me.objDOItem.GetJmlPendingDO(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                    JmlAvailable = JmlFree - JmlPendingDO
                    If JmlAvailable < 0 Then
                        If dummy > JmlFree - (JmlPendingDO - JmlDGV) Then
                            dummy = JmlFree - (JmlPendingDO - JmlDGV)
                        End If
                    End If
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(dummy), _
                    Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)) + _
                    " " + CStr(Me.ObjProduct.GetUnitCode(.Cells(Me.ColPrimaryCodeNo.Index).Value)) + ".", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, _
                e.RowIndex).Value, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))

                'perhitungan berkaitan dengan jumlah available to deliver
                'jumlah available to deliver harus memperhitungkan lokasi dan batch number
                If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                    SecQty = -1
                Else
                    SecQty = .Cells(Me.ColSecondaryQty.Index).Value
                End If
                JmlFree = Me.ObjProduct.GetJml3(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                SecQty, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColBatch.Index).Value) _
                - Me.ObjBookingInvItem.GetJml7(.Cells(Me.ColPrimaryCodeNo.Index).Value, SecQty, _
                .Cells(Me.ColBatch.Index).Value, .Cells(Me.ColPrimaryLocation.Index).Value) + _
                Me.ObjBookingInvItem.GetJml8(Me.txtPrimaryBooking.Text, _
                .Cells(Me.ColPrimaryCodeNo.Index).Value, SecQty, .Cells(Me.ColBatch.Index).Value, _
                .Cells(Me.ColPrimaryLocation.Index).Value) - Me.GetJmlDGV4(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatch.Index).Value, _
                SecQty, .Cells(Me.ColPrimaryLocation.Index).Value)
                JmlPendingDO = Me.objDOItem.GetJmlPendingDO3_Default(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatch.Index).Value, SecQty)
                JmlAvailable = JmlFree - JmlPendingDO
                JmlDGV = Me.GetJmlDGV4(.Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatch.Index).Value, _
                SecQty, .Cells(Me.ColPrimaryLocation.Index).Value)

                If JmlAvailable < 0 Then
                    MsgBox("Maximum available stock is " + Me.ObjCD.DigitGrouping(CStr(JmlFree - (JmlPendingDO - JmlDGV)), Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)) + _
                    " " + CStr(Me.ObjProduct.GetUnitCode(.Cells(Me.ColPrimaryCodeNo.Index).Value)) + ".", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.IsInventoryInstalled = True Then
                    Dim ED As Date
                    If IsDate(.Cells(Me.ColExpiredDate.Index).Value) = True Then
                        ED = .Cells(Me.ColExpiredDate.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        SecQty = -1
                    Else
                        SecQty = .Cells(Me.ColSecondaryQty.Index).Value
                    End If
                    'hanya ketika ada batch no dan sec qty diisikan
                    Dim BatchNoCont As Boolean = Me.ObjProduct.ApakahBatchNoControlled(.Cells("ColPrimaryCodeNo").Value)
                    Dim UseSecondary As Boolean = Me.ObjProduct.ApakahUseSecondary(.Cells("ColPrimaryCodeNo").Value)
                    Dim CekQty As Boolean = True
                    If BatchNoCont = True And .Cells(Me.ColBatch.Index).Value = "" Then
                        CekQty = False
                    End If
                    If CekQty = True Then
                        If UseSecondary = True Then
                            If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                                If UseSecondary = True And .Cells(Me.ColSecondaryQty.Index).Value = -1 Then
                                    CekQty = False
                                End If
                            Else
                                CekQty = False
                            End If
                        End If
                    End If
                    If CekQty = True Then
                        If Me.ApakahEditSah(row, .Cells("ColPrimaryCodeNo").Value, .Cells(Me.ColBatch.Index).Value, ED, _
                        SecQty, JmlDGV, True) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    End If
                End If


                'jika jumlah menurun maka harus load ulang data
                Me.SetGboItem()
                Me.IsiCboProduct(Me.TxtPrimarySONumber.Text)
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
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value <> CStr(-1) Then
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If

                'untuk yang berkaitan dengan order balance, tidak ada pengaruhnya ke sales order balance

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                'msgbox("belum")
                If Me.IsInventoryInstalled = True Then
                    Dim ED As Date
                    If IsDate(.Cells(Me.ColExpiredDate.Index).Value) = True Then
                        ED = .Cells(Me.ColExpiredDate.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    Dim SecQty As Decimal
                    If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        SecQty = -1
                    Else
                        SecQty = .Cells(Me.ColSecondaryQty.Index).Value
                    End If
                    'hanya ketika ada batch no dan sec qty diisikan
                    Dim BatchNoCont As Boolean = Me.ObjProduct.ApakahBatchNoControlled(.Cells("ColPrimaryCodeNo").Value)
                    Dim UseSecondary As Boolean = Me.ObjProduct.ApakahUseSecondary(.Cells("ColPrimaryCodeNo").Value)
                    Dim CekQty As Boolean = True
                    If BatchNoCont = True And .Cells(Me.ColBatch.Index).Value = "" Then
                        CekQty = False
                    End If
                    If CekQty = True Then
                        If UseSecondary = True Then
                            If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                                If UseSecondary = True And .Cells(Me.ColSecondaryQty.Index).Value = -1 Then
                                    CekQty = False
                                End If
                            Else
                                CekQty = False
                            End If
                        End If
                    End If
                    If CekQty = True Then
                        SecondaryQty = Me.buffer
                        If Me.ApakahEditSah(row, .Cells("ColPrimaryCodeNo").Value, BatchNo, ED, _
                        SecQty, JmlDGV) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
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
        Dim BufferExpiryDate As String = "-"

        Try
            With row
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper

                'tidak perlu cek berkaitan dengan jumlah order balance, karena tidak ada pengaruhnya pada 
                'order balance

                'cek berkaitan dengan jumlah digudang
                If Me.IsInventoryInstalled = True Then
                    Dim ED As Date
                    If IsDate(.Cells(Me.ColExpiredDate.Index).Value) = True Then
                        ED = .Cells(Me.ColExpiredDate.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    Dim SecQty As Decimal
                    If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        SecQty = -1
                    Else
                        SecQty = .Cells(Me.ColSecondaryQty.Index).Value
                    End If
                    'hanya ketika ada batch no dan sec qty diisikan
                    Dim BatchNoCont As Boolean = Me.ObjProduct.ApakahBatchNoControlled(.Cells("ColPrimaryCodeNo").Value)
                    Dim UseSecondary As Boolean = Me.ObjProduct.ApakahUseSecondary(.Cells("ColPrimaryCodeNo").Value)
                    Dim CekQty As Boolean = True
                    If BatchNoCont = True And .Cells(Me.ColBatch.Index).Value = "" Then
                        CekQty = False
                    End If
                    If CekQty = True Then
                        If UseSecondary = True Then
                            If IsNumeric(.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                                If UseSecondary = True And .Cells(Me.ColSecondaryQty.Index).Value = -1 Then
                                    CekQty = False
                                End If
                            Else
                                CekQty = False
                            End If
                        End If
                    End If
                    If CekQty = True Then
                        If Me.ApakahEditSah(row, .Cells("ColPrimaryCodeNo").Value, .Cells(Me.ColBatch.Index).Value, ED, _
                        SecQty, JmlDGV) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    End If
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value) = True Then
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItem.Item("ColPrimaryCodeNo", e.RowIndex).Value, _
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

                        'data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItem.Rows
                            If row1.Cells(Me.ColBatch.Index).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells("ColPrimaryCodeNo").Value = Me.dgvItem.Item("ColPrimaryCodeNo", e.RowIndex).Value Then
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
                    row.Cells(Me.ColExpiredDate.Index).Value = row1.Cells(Me.ColExpiredDate.Index).Value And row.Cells(Me.ColBatch.Index).Value = row1.Cells(Me.ColBatch.Index).Value Then
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
                    If CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
                    CStr(row.Cells(Me.ColExpiredDate.Index).Value) = CStr(row1.Cells(Me.ColExpiredDate.Index).Value) Then
                        If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True And IsNumeric(row1.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                            If CDec(row.Cells(Me.ColSecondaryQty.Index).Value) = CDec(row1.Cells(Me.ColSecondaryQty.Index).Value) Then
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

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, _
    ByVal PrimaryProduct As Decimal, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        'msgbox("belum")
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Decimal
        'Dim ObjSC As New ClsTransaksiStockCard

        Try
            JmlGudang = Me.ObjProduct.GetJml4(PrimaryProduct, SecondaryQty, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.ObjBookingInvItem.GetJml3(Me.txtPrimaryBooking.Text, PrimaryProduct, _
            SecondaryQty, BatchNo)
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, BatchNo, SecondaryQty)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal

            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGVLain) + " " + _
                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'cek stock untuk tanggal yang bersangkutan
            'msgbox("belum")

            'sampe sini 
            Return True
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
        Try
            If Me.btnNew.Enabled = False Then
                If Me.dgvItem.Rows.Count = 0 And Me.TxtRemark.Enabled = True Then
                    Me.dgvItem.TabStop = False
                    If Me.TxtRemark.Enabled = True Then
                        Me.btnUP.Enabled = False
                        Me.btnDown.Enabled = False
                    End If
                    Me.TxtSONumber.Enabled = True
                    Me.BtnFindSONumber.Enabled = True
                End If
                Me.SetJmlBO_N_FreeQty()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub DtpCreatedDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpCreatedDate.Enter
        Try
            Me.buffer = Me.DtpCreatedDate.Value.Date.ToShortDateString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Dim row1 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow

            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)

            'cek sampai dimana perlu load ulang data
            If Me.TxtPrimarySONumber.Text = "" Or Me.TxtPrimarySONumber.Text = "-1" Then
                Me.SetGboItem()
            Else
                Dim ketemu As Boolean = False
                For Each row1 In Me.TBL_5.Rows
                    If row1.Primary_Product = Me.bufferProduct Then
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False Then
                    'load ulang data cbo product
                    Me.IsiCboProduct(Me.TxtPrimarySONumber.Text)
                Else
                    'so sama dan product sama
                    Me.SetJmlBO_N_FreeQty()
                    'load ulang sec qty dan batch no
                    Dim BatchNoCont As Boolean = Me.ObjProduct.ApakahBatchNoControlled(Me.bufferProduct)
                    If BatchNoCont = True Then
                        Me.IsiCboBatchNo(Me.bufferProduct, Me.TxtPrimarySONumber.Text)
                    Else
                        Dim SecQtyCont As Boolean = Me.ObjProduct.ApakahUseSecondary(Me.bufferProduct)
                        If SecQtyCont = True Then
                            Me.IsiCboSecondary(Me.bufferProduct, Me.txtPrimaryBooking.Text, Me.bufferBatchNo, Me.bufferPrimaryLocation)
                        End If
                    End If
                End If
            End If

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
            '          If Me.TxtPrimarySONumber.Text = CStr(Me.dgvItem.Item(Me.ColPrimarySO.Index, e.Row.Index).Value) Then
            ' Me.SamaSO = True
            ' Else
            ' Me.SamaSO = False
            ' End If
            If Me.txtPrimaryCodeNo.Text = CStr(Me.dgvItem(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) Then
                Me.SamaCodeNo = True
            Else
                Me.SamaCodeNo = False
            End If
            '           Me.TxtPrimarySONumber.Text = CStr(Me.dgvItem.Item(Me.ColPrimarySO.Index, e.Row.Index).Value)
            Me.bufferProduct = CStr(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value)
            Me.bufferBatchNo = CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.Row.Index).Value)
            Me.bufferPrimaryLocation = CStr(Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.Row.Index).Value)
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
            Me.TxtRemark.BackColor = Me.txtPrimaryCodeNo.BackColor
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
            Me.TxtQtyperBatch.BackColor = Me.txtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboSecondaryQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSecondaryQty.GotFocus
        Try
            Me.cboSecondaryQty.BackColor = Core_Module.GetSMESSBackColorAll
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
            Me.TxtRemarkItem.BackColor = Core_Module.GetSMESSBackColorAll
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
            If Me.chkLocationControlled.Checked = True Then
                If Me.cboBatchNo.Text <> "" Then
                    'isikan data lokasi
                    Me.IsiCboLokasi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, Me.cboBatchNo.Text)
                End If
            Else
                'kosongkan data lokasi
                Me.cboLocation.DataSource = Nothing
                Me.cboLocation.Items.Clear()
                Me.SetJmlBO_N_FreeQty()
            End If
            'isikan data expiry date 
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.DtpExpired.Value = Me.ObjProduct.GetExpiryDate(Me.txtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
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
                Me.IsiCboSecondary(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, _
                Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
            Else
                If Me.chkLocationControlled.Checked = True Then
                    Me.IsiCboLokasi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, Me.cboBatchNo.Text)
                End If
                'Me.SetJmlBO_N_FreeQty()
            End If

            Me.SesuaikanDenganPilihanCBO()
            Me.SesuaikanDenganPilihanCBODGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboLokasi(ByVal vPrimaryProduct As Decimal, ByVal vPrimaryTransaksi As Decimal, ByVal vBatchNo As String)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboLocation(vPrimaryProduct, Me.cboLocation, _
            Me.TBL_3, vBatchNo)
            Dim row1 As DseTransaksiBookingInv.DataTable5Row
            Dim row2 As dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjBookingInvItem.IsikanTBLLocation(Me.TBL_4, vPrimaryProduct, vPrimaryTransaksi, vBatchNo)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL_4.Rows.Count > 0
                    row1 = Me.TBL_4.Rows(0)
                    For Each row2 In Me.TBL_3.Rows
                        If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL_3.Rows.Add(row1.Lot_No, row1.Location, row1.Qty)
                    End If
                    Me.TBL_4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLLocation(vPrimaryProduct, vBatchNo)
            While Me.TBL_4.Rows.Count > 0
                row1 = Me.TBL_4.Rows(0)
                For Each row2 In Me.TBL_3.Rows
                    If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL_3.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL_3.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL_4.Rows.Remove(row1)
            End While

            Me.TBL_3.DefaultView.Sort = "Location ASC"
            Me.cboLocation.DataSource = Me.TBL_3.DefaultView
            Me.cboLocation.ValueMember = "Location"
            Me.cboLocation.DisplayMember = "Location"
            Me.cboLocation.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLLocation(ByVal vPrimaryProduct As Decimal, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiBookingInv.DataTable5Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                CStr(row.Cells(Me.ColBatch.Index).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL_4.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColBatch.Index).Value) And _
                        row1.Location = row.Cells(Me.ColLocation.Index).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDec(row.Cells(Me.ColQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL_4.Rows.Add(CStr(row.Cells(Me.ColBatch.Index).Value), row.Cells(Me.ColLocation.Index).Value, _
                        CDec(row.Cells(Me.ColQty.Index).Value))
                    End If
                End If
            Next
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

    Private Sub cboCodeNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedValueChanged
        Try
            If Me.cboCodeNo.SelectedIndex <> -1 Then
                Dim RowView As DataRowView = Me.cboCodeNo.Items(Me.cboCodeNo.SelectedIndex)
                Dim row2 As DseTransaksiSO.laporan_sales_orderbalance_view9dRow = RowView.Row
                Dim PrimaryCodeNoBaru As Double = row2.Primary_Product

                If Me.txtPrimaryCodeNo.Text <> PrimaryCodeNoBaru Then
                    Me.txtPrimaryCodeNo.Text = PrimaryCodeNoBaru
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)

                    'cek fitur apa saja yang dikontrol
                    Me.ObjProduct.IsiCheckBoxControll(Me.txtPrimaryCodeNo.Text, Me.chkLocationControlled, _
                    Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
                    Me.chkPdoControlled, Me.txtPrimaryUnits, Me.txtPrimarySecondaryUnits)

                    Me.cboBatchNo.DataSource = Nothing
                    Me.cboLocation.DataSource = Nothing
                    Me.cboSecondaryQty.DataSource = Nothing

                    If Me.chkBatchNoControlled.Checked = True Then
                        Me.cboBatchNo.TabStop = True
                        Me.lblBatchNo.Visible = False
                        Me.cboBatchNo.Visible = True
                        'isikan data batch no dari basis data
                        Me.IsiCboBatchNo(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text)
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
                            Me.IsiCboLokasi(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, Me.cboBatchNo.Text)
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
                        Me.cboSecondaryQty.Visible = False
                        Me.cboSecondaryQty.DataSource = Nothing
                        Me.cboSecondaryQty.Items.Clear()
                    Else
                        Me.lblSecondaryUnits.Visible = True
                        Me.lblSecQty.Visible = True
                        Me.lblUnits.Left = Me.lblSecQty.Left - 37
                        Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
                        Me.cboSecondaryQty.Visible = True
                        If Me.chkLocationControlled.Checked = False Then
                            'isikan data secondary jika perlu
                            Me.IsiCboSecondary(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, _
                            Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                        End If
                    End If
                    If Me.txtPrimaryCodeNo.Text <> "-1" Then
                        Me.lblUnits.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)
                        If Me.chkUseSecondaryUnit.Checked = True Then
                            Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.txtPrimaryCodeNo.Text)
                        End If
                    Else
                        Me.lblUnits.Text = "Units"
                        Me.lblSecondaryUnits.Text = "Units"
                    End If
                End If
            Else
                'reset tampilan
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Me.txtPrimaryUnits.Text = "-1"
                Me.txtPrimarySecondaryUnits.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.TxtQtyperBatch.Text = "0"
                Me.cboSecondaryQty.Text = "0"
                Me.txtQtyOnHand.Text = "0"
                Me.txtQtyOrderBalance.Text = "0"
                Me.chkBatchNoControlled.Checked = False
                Me.chkExpiryDateControlled.Checked = False
                Me.chkUseSecondaryUnit.Checked = False
                Me.lblSecondaryUnits.Visible = False
                Me.lblUnits.Left = Me.lblSecondaryUnits.Left
                Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
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
            End If

            'format tampilan
            If Me.txtPrimaryCodeNo.Text <> "-1" Then
                Me.lblUnits.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.txtPrimaryCodeNo.Text)
                End If
                Me.lblUnitFree.Text = Me.lblUnits.Text
                Me.LblUnitOB.Text = Me.lblUnits.Text
                Me.lblUnitsPendingDO.Text = Me.lblUnits.Text
                Me.lblAvailableUnits.Text = Me.lblUnits.Text
                Me.TxtQtyperBatch.Text = "0"
                Me.SetJmlBO_N_FreeQty()
            Else
                Me.lblUnits.Text = "Units"
                Me.lblSecondaryUnits.Text = "Units"
                Me.lblAvailableUnits.Text = "Units"
                Me.lblUnitsPendingDO.Text = "Units"
                Me.lblUnitFree.Text = "Units"
                Me.LblUnitOB.Text = "Units"
            End If
            Me.SesuaikanDenganPilihanCBO()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboBatchNo(ByVal vPrimaryProduct As Decimal, ByVal vPrimaryBooking As Decimal)
        Try
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNo, Me.TBL_1)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiBookingInv.tbltrans_booking_invitem3Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjBookingInvItem.IsiTBLBatchNo(vPrimaryBooking, vPrimaryProduct, Me.TBL_2)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL_2.Rows.Count > 0
                    row1 = Me.TBL_2.Rows(0)
                    For Each row2 In Me.TBL_1.Rows
                        If row1.Lot_No = row2.Lot_No Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL_1.Rows.Add(row1.Lot_No, row1.Expiry_Date, row1.Qty)
                    End If
                    Me.TBL_2.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLBatchNo(vPrimaryProduct)
            While Me.TBL_2.Rows.Count > 0
                row1 = Me.TBL_2.Rows(0)
                For Each row2 In Me.TBL_1.Rows
                    If row1.Lot_No = row2.Lot_No Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL_1.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL_1.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL_2.Rows.Remove(row1)
            End While
            Me.TBL_1.DefaultView.Sort = "[Lot No] ASC"
            Me.cboBatchNo.DataSource = Me.TBL_1.DefaultView
            Me.cboBatchNo.ValueMember = "Lot No"
            Me.cboBatchNo.DisplayMember = "Lot No"
            Me.cboBatchNo.Refresh()

            If Me.chkLocationControlled.Checked = False And _
            Me.chkUseSecondaryUnit.Checked = True And Me.cboBatchNo.Text <> "" Then
                Me.IsiCboSecondary(vPrimaryProduct, vPrimaryBooking, Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindSONumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindSONumber.Click
        Dim ij As New FrmCariTransaksiSO
        Try
            ij.BackColorAll = Me.BackColorAll
            ij.txtCari.Text = Me.TxtSONumber.Text
            ij.ShowDialog()

            If Not ij.PrimaryTransaksi = -1 Then
                Me.TxtPrimarySONumber.Text = ij.PrimaryTransaksi
                Me.TxtSONumber.Text = ObjSO.GetNoSO(Me.TxtPrimarySONumber.Text)
                Me.TxtPrimaryCustomer.Text = ObjSO.GetPrimaryCustomer(Me.TxtPrimarySONumber.Text)
                Me.TxtCustomerCode.Text = ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                Me.IsiCboProduct(Me.TxtPrimarySONumber.Text)
                Me.TBL9.Clear()
                Dim DTASO As New DseTransaksiBookingInvTableAdapters.tbltrans_so_itemTableAdapter
                DTASO.FillBy_PrimarySO(TBL9, Me.TxtPrimarySONumber.Text)
                DTASO.Dispose()
                Me.cboValidUntil.Focus()
            Else
                'kosongkan combo box
                Me.cboCodeNo.DataSource = Nothing
                Me.KosongkanItem()
                Me.TxtCustomerCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtSONumber_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSONumber.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim primary As Decimal = ObjSO.GetPrimary(Me.TxtSONumber.Text)
                If primary <> -1 Then
                    Me.TxtPrimarySONumber.Text = primary
                    Me.TxtSONumber.Text = ObjSO.GetNoSO(primary)
                    Me.TxtPrimaryCustomer.Text = ObjSO.GetPrimaryCustomer(Me.TxtPrimarySONumber.Text)
                    Me.TxtCustomerCode.Text = ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                    Me.IsiCboProduct(Me.TxtPrimarySONumber.Text)
                    Me.TBL9.Clear()
                    Dim DTASO As New DseTransaksiBookingInvTableAdapters.tbltrans_so_itemTableAdapter
                    DTASO.FillBy_PrimarySO(TBL9, Me.TxtPrimarySONumber.Text)
                    DTASO.Dispose()
                Else
                    Me.BtnFindSONumber.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSONumber.Leave
        Try
            If Me.TxtSONumber.Text <> "" Then
                Dim primary As Decimal = ObjSO.GetPrimary(Me.TxtSONumber.Text)
                If primary <> -1 Then
                    Me.TxtPrimarySONumber.Text = primary
                    Me.TxtSONumber.Text = ObjSO.GetNoSO(primary)
                    Me.TxtPrimaryCustomer.Text = ObjSO.GetPrimaryCustomer(Me.TxtPrimarySONumber.Text)
                    Me.TxtCustomerCode.Text = ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomer.Text)
                    Me.IsiCboProduct(Me.TxtPrimarySONumber.Text)
                    Me.TBL9.Clear()
                    Dim DTASO As New DseTransaksiBookingInvTableAdapters.tbltrans_so_itemTableAdapter
                    DTASO.FillBy_PrimarySO(TBL9, Me.TxtPrimarySONumber.Text)
                    DTASO.Dispose()
                Else
                    Me.BtnFindSONumber.PerformClick()
                End If
            Else
                'kosongkan cbo
                Me.cboCodeNo.DataSource = Nothing
                Me.KosongkanItem()
                Me.TxtCustomerCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboValidUntil_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboValidUntil.SelectedIndexChanged
        If Me.TxtRemark.Enabled = True Then
            Me.SesuaikanDenganPilihanCBO()
            Me.SesuaikanDenganPilihanCBODGV()
            Me.FormatDGV()
        End If
    End Sub

    Private Sub SesuaikanDenganPilihanCBO()
        Try
            Dim row As DseTransaksiBookingInv.tbltrans_so_itemRow
            Dim DeliveryCheck As Boolean = False
            Dim deliveryDate As DateTime = #1/1/1950#

            Dim CodeNo As String = ""
            If Me.cboCodeNo.SelectedIndex > -1 Then
                Dim DV As DataRowView
                DV = Me.cboCodeNo.SelectedValue
                CodeNo = DV(3)
            End If
            For Each row In TBL9.Rows
                If row.Product_Code = CodeNo Then
                    DeliveryCheck = row.Delivery_Date_Cont
                    deliveryDate = row.Delivery_Date
                    Exit For
                End If
            Next

            If DeliveryCheck = True Then
                If deliveryDate < Me.DtpCreatedDate.Value.AddDays(CDec(Me.cboValidUntil.SelectedIndex) + 1) Then
                    Me.dtpValid1.Enabled = True
                    Me.dtpValid1.Value = Me.DtpCreatedDate.Value.AddDays(CDec(Me.cboValidUntil.SelectedIndex) + 1)
                    Me.dtpValid1.Enabled = False
                Else
                    Me.dtpValid1.Enabled = True
                    Me.dtpValid1.Value = deliveryDate
                    Me.dtpValid1.Enabled = False
                End If
            Else
                Me.dtpValid1.Enabled = True
                Me.dtpValid1.Value = Me.DtpCreatedDate.Value.AddDays(CDec(Me.cboValidUntil.SelectedIndex) + 1)
                Me.dtpValid1.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SesuaikanDenganPilihanCBODGV()
        Try
            Dim row As DseTransaksiBookingInv.tbltrans_so_itemRow
            Dim rowDGV As DataGridViewRow
            Dim DeliveryCheck As Boolean = False
            Dim deliveryDate As DateTime = #1/1/1950#
            For Each rowDGV In Me.dgvItem.Rows
                For Each row In TBL9.Rows
                    If row.Product_Code = Me.dgvItem.Item(Me.ColCodeNo.Index, rowDGV.Index).Value Then
                        DeliveryCheck = row.Delivery_Date_Cont
                        deliveryDate = row.Delivery_Date
                        Exit For
                    End If
                Next
                If DeliveryCheck = True Then
                    If deliveryDate < Me.DtpCreatedDate.Value.AddDays(CDec(Me.cboValidUntil.SelectedIndex) + 1) Then
                        '    Me.dtpValid1.Enabled = True
                        Me.dgvItem.Item(Me.ColValidUntil.Index, rowDGV.Index).Value = _
                        Format(Me.DtpCreatedDate.Value.AddDays(CDec(Me.cboValidUntil.SelectedIndex) + 1), "dd MMM yyyy")
                        '       Me.dtpValid1.Enabled = False
                    Else
                        Me.dgvItem.Item(Me.ColValidUntil.Index, rowDGV.Index).Value = deliveryDate
                    End If
                Else
                    Me.dgvItem.Item(Me.ColValidUntil.Index, rowDGV.Index).Value = _
                    Format(Me.DtpCreatedDate.Value.AddDays(CDec(Me.cboValidUntil.SelectedIndex) + 1), "dd MMM yyyy")
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpCreatedDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DtpCreatedDate.ValueChanged
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjBookingInv.TanggalTransaksi.Month = Me.DtpCreatedDate.Value.Month And Me.ObjBookingInv.TanggalTransaksi.Year = Me.DtpCreatedDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtBookingNo.Text = Me.ObjBookingInv.BookingNo
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtBookingNo.Text = Me.ObjBookingInv.GetNoBaru(Me.DtpCreatedDate.Value)
                    End If
                Else
                    Me.TxtBookingNo.Text = Me.ObjBookingInv.GetNoBaru(Me.DtpCreatedDate.Value)
                End If

                If Me.TxtCustomerCode.Text <> "" Then
                    Me.SesuaikanDenganPilihanCBO()
                End If
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.SesuaikanDenganPilihanCBODGV()
                End If
                Me.FormatDGV()
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
                Me.SetJmlBO_N_FreeQty()
            Else
                Me.txtPrimaryLocation.Text = objLoc.GetPrimary(Me.cboLocation.Text)
                Me.TxtWarehouse.Text = objLoc.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.IsiCboSecondary(Me.txtPrimaryCodeNo.Text, Me.txtPrimaryBooking.Text, Me.cboBatchNo.Text, _
                    Me.txtPrimaryLocation.Text)
                Else
                    Me.SetJmlBO_N_FreeQty()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objLoc = Nothing
        End Try
    End Sub

    Private Sub TxtSONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSONumber.GotFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtSONumber.ReadOnly = False Then
                    Me.TxtSONumber.BackColor = Me.BackColorAll
                Else
                    Me.TxtSONumber.BackColor = Me.TxtCustomerCode.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSONumber.LostFocus
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtSONumber.ReadOnly = False Then
                    Me.TxtSONumber.BackColor = Me.txtPrimaryBooking.BackColor
                Else
                    Me.TxtSONumber.BackColor = Me.TxtCustomerCode.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryCodeNo.Text <> "" And Me.txtPrimaryCodeNo.Text <> "-1" Then
                    Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, _
                    Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCodeNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCodeNo.SelectedIndexChanged

    End Sub

    'Update SMESS 2026
    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemark.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemark.Enabled = False Then
                If Me.ObjBookingInv.CariDetail(vPrimary, Me.dgvItem, Me) = True Then
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

End Class