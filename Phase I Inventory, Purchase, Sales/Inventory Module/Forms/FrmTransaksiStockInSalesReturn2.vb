Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockInSalesReturn2
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Public TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
    Public objStockIn As New clsTransaksiStockInOut

    Dim baru As Boolean
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
    Dim ObjLicense As New ClsLicense
    Dim buffer As String = ""
    Dim ObjAcc As New Core_Module.ClsTranskasiAccClosing
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjUser As New ClsUser

    'index2 dibawah ini akan digunakan sebagai acuan ketika menambahkan baris baru dari sales module
    Public Index_No As Integer = -1
    Public Index_BatchNo As Integer = -1
    Public Index_CodeNo As Integer = -1
    Public Index_Desc As Integer = -1
    Public Index_Expiry As Integer = -1
    Public Index_Location As Integer = -1
    Public Index_PONoDONo As Integer = -1
    Public Index_PrimaryCodeNo As Integer = -1
    Public Index_PrimaryLocation As Integer = -1
    Public Index_PrimaryPdO As Integer = -1
    Public Index_PrimarySecUnit As Integer = -1
    Public Index_PrimarySO As Integer = -1
    Public Index_PrimaryUnits As Integer = -1
    Public Index_Qty As Integer = -1
    Public Index_Remarks As Integer = -1
    Public Index_SecQty As Integer = -1
    Public Index_SecUnit As Integer = -1
    Public Index_SuppPODONo As Integer = -1
    Public Index_Unit As Integer = -1
    Public Index_Warehouse As Integer = -1

    Public Index_QtyMulti As Integer = -1
    Public Index_UnitsCodeMulti As Integer = -1
    Public Index_PrimaryUnitsMulti As Integer = -1
    Public Index_RatioReference As Integer = -1
    Public Index_RatioDefault As Integer = -1
    Public Index_Ratio_Desc As Integer = -1

    Public TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, False, False)
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
            Me.SetGboItem()
            Me.FormatDGV()

            If Me.dgvItem.RowCount > 0 Then
                Me.txtDONo.Enabled = False
                Me.btnFindDO.Enabled = False
            Else
                Me.txtDONo.Enabled = True
                Me.btnFindDO.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetInputCustom()
        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockInSalesReturn_Custom") = False Then
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
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, False, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, False, False)
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
            Me.TxtPrimaryPihakAcuan.Text = "-1"
            Me.txtPrimaryDO.Text = "-1"
            Me.txtPrimarySO.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
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
            Me.txtPrimarySO.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyperBatchDefault.Text = "0"
            Me.TxtQtyperBatchMulti.Text = "0"
            Me.txtSecondaryQty.Text = "-1"
            Me.chkBatchNoControlled.Checked = False
            Me.chkExpiryDateControlled.Checked = False
            Me.chkLocationControlled.Checked = False
            Me.chkPdoControlled.Checked = False
            Me.chkUseSecondaryUnit.Checked = False
            Me.txtSecondaryQty.Visible = False
            Me.lblSecondaryUnits.Visible = False
            'Me.lblUnitsDefault.Left = Me.lblSecondaryUnits.Left
            'Me.TxtQtyperBatchDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyperBatchDefault.Left - 6
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
            Me.TxtRatioDefault.Text = "1"
            Me.TxtRatioReference.Text = "1"
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
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
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
            If Me.chkBatchNoControlled.Checked = True Then
                Me.txtBatchNo.Text = Me.txtBatchNo.Text.Trim
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

            '9. Bila setelah diinput jumlah product menjadi minus, kalo input baru ga mungkin
            '10. Bila transaksi berkaitan dengan sales dan terinstall software sales maka ada batasan 
            '    berkaitan dengan jumlah item DO. Berkaitan dengan jumlah item DO secara global, 
            '    hanya memperdulikan primary SO dan primary product
            If Me.ApakahKontrolStateSama(Me.TxtPrimaryCodeNo.Text, Me.txtPrimarySO.Text, Me.txtPrimaryDO.Text) = True Then
                'pada bagian ini pengecekan jika control state yang digunakan sama
                Dim JmlDo As Decimal = Me.ObjDO.GetJml3_Default(Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtBatchNo.Text, _
                Me.txtSecondaryQty.Text)
                Dim JmlDGV As Decimal = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtBatchNo.Text, Me.txtPrimarySO.Text)
                Dim JmlObj, JmlReturn As Decimal
                JmlReturn = Me.ObjStockInOutItem.GetJmlReturn_Default(Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryDO.Text, _
                Me.txtBatchNo.Text, Me.txtSecondaryQty.Text)
                If Me.baru = True Then
                    JmlObj = 0
                Else
                    JmlObj = Me.ObjStockInOutItem.GetJmlReturn_Default(Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, _
                    Me.txtPrimaryDO.Text, Me.TxtPrimaryStockIn.Text, Me.txtBatchNo.Text, Me.txtSecondaryQty.Text)
                End If

                If JmlDo - JmlDGV - CDbl(Me.TxtQtyperBatchDefault.Text) - JmlReturn + JmlObj < 0 Then
                    Dim Msg As String = "Maximum value you can enter for:" + Me.TxtCodeNo.Text
                    If Me.txtBatchNo.Text <> "" Then
                        Msg = Msg + ", batch no: " + Me.txtBatchNo.Text
                    End If
                    If Me.txtSecondaryQty.Text <> "-1" And Me.txtSecondaryQty.Text <> "" Then
                        Msg = Msg + ", secondary qty: " + Me.txtSecondaryQty.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                    End If
                    Dim Max_ As String = ObjCD.DigitGrouping(CStr(CDec(JmlDo - JmlDGV - JmlReturn + JmlObj) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text)), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    MsgBox(Msg + " and SO number: " + _
                    Me.txtSalesOrderNo.Text + " is " + Max_ + " " + lblUnitsMulti.Text + ".", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtQtyperBatchMulti.Focus()
                    Return False
                End If
            Else
                'pada bagian ini jika control state yang digunakan tidak sama, maka acuan 
                'yang digunakan hanya product code, SO Number dan DO number
                Dim JmlDo As Decimal = Me.ObjDO.GetJml2_Default(Me.txtPrimaryDO.Text, Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text)
                Dim JmlDGV As Decimal = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.txtPrimarySO.Text)
                Dim JmlObj, JmlReturn As Decimal
                JmlReturn = Me.ObjStockInOutItem.GetJmlReturn(Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryDO.Text)
                If Me.baru = True Then
                    JmlObj = 0
                Else
                    JmlObj = Me.ObjStockInOutItem.GetJmlReturn(Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryDO.Text, Me.TxtPrimaryStockIn.Text)
                End If

                If JmlDo - JmlDGV - CDbl(Me.TxtQtyperBatchDefault.Text) - JmlReturn + JmlObj < 0 Then
                    Dim Msg As String = "Maximum value you can enter for:" + Me.TxtCodeNo.Text
                    If Me.txtBatchNo.Text <> "" Then
                        Msg = Msg + ", batch no: " + Me.txtBatchNo.Text
                    End If
                    If Me.txtSecondaryQty.Text <> "-1" And Me.txtSecondaryQty.Text <> "" Then
                        Msg = Msg + ", secondary qty: " + Me.txtSecondaryQty.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                    End If
                    Dim Max_ As String = ObjCD.DigitGrouping(CStr(CDec(JmlDo - JmlDGV - JmlReturn + JmlObj) * CDec(Me.TxtRatioReference.Text) / CDec(Me.TxtRatioDefault.Text)), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                    MsgBox(Msg + " and SO number: " + _
                    Me.txtSalesOrderNo.Text + " is " + Max_ + " " + lblUnitsMulti.Text + ".", MsgBoxStyle.Information, _
                    "Information")
                    Me.TxtQtyperBatchMulti.Focus()
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' hanya digunakan oleh fungsi diatas ini :)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahKontrolStateSama(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double, ByVal vPrimaryDO As Double) As Boolean
        Dim ApakahAdaED As Boolean
        Dim ApakahAdaBatch As Boolean
        Dim rowDgv As DataGridViewRow

        Try
            ApakahAdaBatch = False
            ApakahAdaED = False

            For Each rowDgv In Me.DgvItemDO2.Rows
                If rowDgv.Cells(Me.ColDO2_PrimarySO.Index).Value = vPrimarySO And rowDgv.Cells(Me.ColDO2_PrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    If rowDgv.Cells(Me.ColDO2_BatchNo.Index).Value <> "" Then
                        ApakahAdaBatch = True
                        Exit For
                    End If
                End If
            Next
            For Each rowDgv In Me.DgvItemDO2.Rows
                If rowDgv.Cells(Me.ColDO2_PrimarySO.Index).Value = vPrimarySO And rowDgv.Cells(Me.ColDO2_PrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    If IsDate(rowDgv.Cells(Me.ColDO2_Expiry.Index).Value) = True Then
                        ApakahAdaED = True
                        Exit For
                    End If
                End If
            Next

            If Me.chkBatchNoControlled.Checked = ApakahAdaBatch And ApakahAdaED = Me.chkExpiryDateControlled.Checked Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception

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

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vBatchNo As String, ByVal vPrimarySO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                row.Cells(Me.ColBatch.Index).Value = vBatchNo Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                        If row.Cells(Me.ColSecondaryQty.Index).Value = Secondary_Qty Then
                            Jml += row.Cells(Me.ColQty.Index).Value
                        End If
                    Else
                        If Secondary_Qty = -1 Then
                            Jml += row.Cells(Me.ColQty.Index).Value
                        End If
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vBatchNo As String, ByVal vPrimarySO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And _
                row.Cells(Me.ColBatch.Index).Value = vBatchNo Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = True Then
                        If row.Cells(Me.ColSecondaryQty.Index).Value = Secondary_Qty Then
                            Jml += row.Cells(Me.ColQtyMulti.Index).Value
                        End If
                    Else
                        If Secondary_Qty = -1 Then
                            Jml += row.Cells(Me.ColQtyMulti.Index).Value
                        End If
                    End If
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
        Dim TglConvSMESS As Date = #1/1/1950#

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh bikin stock in out, pada tanggal sebelum convertion date
            TglConvSMESS = Me.ObjAcc.GetSMESSConvertionDate
            If Me.dtpDate.Value.Date < TglConvSMESS.Date Then
                MsgBox("Minimum transaction date is on SMESS convertion date: " + Format(TglConvSMESS, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Me.dtpDate.Focus()
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

            ''cek bisi ada qty default=0
            Dim ListError As String = ""
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
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

            'Dim objLocationAccess As New FrmUserLocationPrivilegeSetting
            'Dim listOfPrimaryLocation As New List(Of Double)
            'For Each dataGridViewRow As DataGridViewRow In Me.dgvItem.Rows
            '    If dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value <> "-1" Then
            '        If listOfPrimaryLocation.Contains(dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value) = False Then listOfPrimaryLocation.Add(dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value)
            '    End If
            'Next
            'If objLocationAccess.HasAccessPrivileges(Me.UserCode, listOfPrimaryLocation, FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut.STOCK_IN_SALES_RETURN) = False Then
            '    Return False
            'End If

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_IN_SALES_RETURN, Me.TBLDetailLocation) = False Then
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
                    Dim PrimaryPOSO As Double = .Cells(Me.ColPrimarySO.Index).Value
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

            'tidak boleh berdasarkan do yg berasal dari invoice quick
            If Me.ObjDO.ApakahQuick(Me.txtPrimaryDO.Text) = True Then
                MsgBox("You can't enter DO from quick invoice as reference.", MsgBoxStyle.Information, "Information")
                Return False
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

    Public Sub SalinKeObjek()
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
            Me.ObjDOItem.IsiDGV_Multi(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
            Me.FormatDGVDO()
            'digit grouping
            Dim row As DataGridViewRow
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

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemDO2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.WindowState = FormWindowState.Normal Then
                Me.Size = Me.MinimumSize
                Me.StartPosition = FormStartPosition.CenterScreen
            End If
            Index_No = Me.ColNo.Index
            Index_BatchNo = Me.ColBatch.Index
            Index_CodeNo = Me.ColCodeNo.Index
            Index_Desc = Me.ColDescription.Index
            Index_Expiry = Me.ColExpiredDate.Index
            Index_Location = Me.ColLocation.Index
            Index_PONoDONo = Me.ColPONoDONo.Index
            Index_PrimaryCodeNo = Me.ColPrimaryCodeNo.Index
            Index_PrimaryLocation = Me.ColPrimaryLocation.Index
            Index_PrimaryPdO = Me.ColPrimaryPdO.Index
            Index_PrimarySecUnit = Me.ColPrimarySecondaryUnits.Index
            Index_PrimarySO = Me.ColPrimarySO.Index
            Index_PrimaryUnits = Me.ColPrimaryUnits.Index
            Index_Qty = Me.ColQty.Index
            Index_Remarks = Me.ColRemark.Index
            Index_SecQty = Me.ColSecondaryQty.Index
            Index_SecUnit = Me.ColSecondaryUnit.Index
            Index_SuppPODONo = Me.ColSuppDONo.Index
            Index_Unit = Me.ColUnit.Index
            Index_Warehouse = Me.ColWareHouse.Index
            Index_QtyMulti = Me.ColQtyMulti.Index
            Index_UnitsCodeMulti = Me.ColUnitsCodeMulti.Index
            Index_PrimaryUnitsMulti = Me.ColPrimaryUnitsMulti.Index
            Index_RatioReference = Me.ColRatioReference.Index
            Index_RatioDefault = Me.ColRatioDefault.Index
            Index_Ratio_Desc = Me.ColRatioDesc.Index
            Me.KosongkanField()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Me.BtnCariPihakAcuan.Visible = False

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

    Public Sub NewStockInQuickSR()
        Try
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
        Dim hj As New FrmCariTransaksiStockInOut_Sales
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
                If Me.objStockIn.ApakahBolehEdit(Me.UserCode, True) = True Then
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

    Private Function GetJmlDGV_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And vPrimarySO = row.Cells(Me.ColPrimarySO.Index).Value Then
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
            Dim Jml As Double = 0
            Dim OK As Boolean = False

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo And vPrimarySO = row.Cells(Me.ColPrimarySO.Index).Value Then
                    Jml = Jml + CDec(row.Cells(Me.ColQtyMulti.Index).Value)
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

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.chkUseSecondaryUnit, _
            Me.chkPdoControlled, Me.txtPrimaryUnitsDefault, Me.txtPrimarySecondaryUnits)

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
                'Me.lblUnitsDefault.Left = Me.lblSecondaryUnits.Left
                'Me.TxtQtyperBatchDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyperBatchDefault.Left - 6
                Me.txtSecondaryQty.Text = "-1"
                Me.lblSecQty.Visible = False
            Else
                Me.txtSecondaryQty.Visible = True
                Me.lblSecondaryUnits.Visible = True
                Me.lblSecQty.Visible = True
                'Me.lblUnitsDefault.Left = Me.lblSecQty.Left - 37
                'Me.TxtQtyperBatchDefault.Width = Me.lblUnitsDefault.Left - Me.TxtQtyperBatchDefault.Left - 6
                Me.txtSecondaryQty.Text = "0"
            End If
            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.lblUnitsDefault.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.chkUseSecondaryUnit.Checked = True Then
                    Me.lblSecondaryUnits.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                End If
            Else
                Me.lblUnitsDefault.Text = "Units"
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
                        Me.TxtQtyperBatchMulti.Focus()
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
                            Me.TxtQtyperBatchMulti.Focus()
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
                        Me.TxtQtyperBatchMulti.Focus()
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
                            Me.TxtQtyperBatchMulti.Focus()
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
            If Me.txtPrimaryDO.Text <> "" Then
                If Me.txtPrimaryDO.Text = -1 And Me.txtPrimaryDO.Text <> "" Then
                    Me.txtDONo.Focus()
                    MsgBox("To start entering item data, you must specify delivery order data first.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

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
                    Me.TxtQtyperBatchMulti.Focus()
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
    ByVal SecondaryQty As Double, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date, ByVal vPrimarySO As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And row.Cells(Me.ColPrimaryPdO.Index).Value = _
                        PrimaryPdO And ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And row.Cells(Me.ColPrimaryPdO.Index).Value = PrimaryPdO And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimarySO.Index).Value = vPrimarySO And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
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
        Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
        Dim PrimaryLocationLama As Double = -1

        Try
            If Me.ApakahDataItemBenar = True Then
                PrimaryLocationLama = Me.txtPrimaryLocation.Text

                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Me.txtPrimaryPDO.Text, Me.DtpExpired.Value, Me.txtPrimarySO.Text)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) + CDec(Me.TxtQtyperBatchMulti.Text), DefPrecMulti)
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
                    Me.dgvItem.Rows.Add()
                    Dim row As DataGridViewRow
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With row
                        .Cells(Me.ColNo.Index).Value = ""
                        .Cells(Me.ColPONoDONo.Index).Value = Me.txtSalesOrderNo.Text
                        .Cells(Me.ColSuppDONo.Index).Value = ""
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                        .Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                        .Cells(Me.ColBatch.Index).Value = Me.txtBatchNo.Text
                        .Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColLocation.Index).Value = Me.TxtLocation.Text
                        .Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimarySO.Index).Value = Me.txtPrimarySO.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                        .Cells(Me.ColPrimaryPdO.Index).Value = Me.txtPrimaryPDO.Text
                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                        .Cells(Me.ColUnitsCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                        .Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, DefPrecDefault)
                        .Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(Me.TxtRatioReference.Text, DefPrecMulti)
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                        If Me.txtPrimaryUnitsDefault.Text <> Me.txtPrimaryUnitsMulti.Text Then
                            .Cells(Me.ColRatioDesc.Index).Value = Me.TxtRatioReference.Text _
                            & " " & Me.lblUnitsMulti.Text & " = " & Me.TxtRatioDefault.Text & " " & Me.lblUnitsDefault.Text
                        Else
                            .Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                    End With

                    'Me.dgvItem.Rows.Add("", Me.txtSalesOrderNo.Text, "", Me.TxtCodeNo.Text, Me.txtDesc.Text, Me.TxtQtyperBatchDefault.Text, _
                    'Me.lblUnitsDefault.Text, Me.txtSecondaryQty.Text, Me.lblSecondaryUnits.Text, Me.txtBatchNo.Text, _
                    'ExpiredDateStr, Me.TxtLocation.Text, Me.TxtWarehouse.Text, Me.TxtRemarkItem.Text, _
                    'Me.txtPrimarySO.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.txtPrimaryUnitsDefault.Text, _
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

            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnits.Index).Value Then
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
            '4. Expiry Date
            '5. Location Code
            '6. Remarks
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColSecondaryQty.Index Or e.ColumnIndex = Me.ColBatch.Index Or e.ColumnIndex = Me.ColLocation.Index _
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
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItem.Item(Me.ColBatch.Index, e.RowIndex).Value) + " is " + _
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

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefPrecMulti)
                Me.SetQtyDefaultOnDgv(e.RowIndex)
                If CDec(row.Cells(Me.ColQty.Index).Value) <= 0 Then
                    MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(e.RowIndex)
                    Exit Sub
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                If Me.ObjProduct.ApakahGoods(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                    JmlDGV = .Cells(Me.ColQty.Index).Value
                    If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, .Cells(Me.ColPrimaryPdO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Me.SetQtyDefaultOnDgv(e.RowIndex)
                    End If
                End If

                'cek data sales order balance
                Dim JmlDo As Decimal = Me.ObjDO.GetJml3_Multi(Me.txtPrimaryDO.Text, .Cells(Me.ColPrimarySO.Index).Value, _
                .Cells(Me.ColPrimaryCodeNo.Index).Value, .Cells(Me.ColBatch.Index).Value, SecondaryQty)
                JmlDGV = Me.GetJmlDGV_Multi(.Cells(Me.ColPrimaryCodeNo.Index).Value, SecondaryQty, .Cells(Me.ColBatch.Index).Value, .Cells(Me.ColPrimarySO.Index).Value)
                Dim JmlDGVSendiri As Decimal = .Cells(Me.ColQtyMulti.Index).Value
                Dim JmlObj, JmlReturn As Decimal
                JmlReturn = Me.ObjStockInOutItem.GetJmlReturn_Multi(.Cells(Me.ColPrimarySO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                Me.txtPrimaryDO.Text, .Cells(Me.ColBatch.Index).Value, SecondaryQty)
                If Me.baru = True Then
                    JmlObj = 0
                Else
                    JmlObj = Me.ObjStockInOutItem.GetJmlReturn_Multi(.Cells(Me.ColPrimarySO.Index).Value, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    Me.txtPrimaryDO.Text, Me.TxtPrimaryStockIn.Text, .Cells(Me.ColBatch.Index).Value, SecondaryQty)
                End If

                If JmlDo - JmlDGV - JmlReturn + JmlObj < 0 Then
                    Dim Msg As String = "Maximum value you can enter for:" + CStr(.Cells(Me.ColCodeNo.Index).Value)
                    If .Cells(Me.ColBatch.Index).Value <> "" Then
                        Msg = Msg + ", batch no: " + .Cells(Me.ColBatch.Index).Value.ToString
                    End If
                    If .Cells(Me.ColSecondaryQty.Index).Value.ToString <> "-1" And row.Cells(Me.ColSecondaryQty.Index).Value.ToString <> "-" Then
                        Msg = Msg + ", secondary qty: " + SecondaryQty + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(.Cells(Me.ColPrimaryCodeNo.Index).Value)
                    End If

                    MsgBox(Msg + " and SO number: " + _
                    CStr(.Cells(Me.ColPONoDONo.Index).Value) + " is " + Me.ObjCD.DigitGrouping(CStr(JmlDo - JmlDGV + JmlDGVSendiri - JmlReturn + JmlObj), Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)) + " " + _
                    Me.dgvItem.Item(Me.ColUnitsCodeMulti.Index, e.RowIndex).Value + ".", MsgBoxStyle.Information, "Information")

                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                Me.SetQtyDefaultOnDgv(e.RowIndex)
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
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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
                                If CDate(row1.Cells(Me.ColExpiredDate.Index).Value) <> CDate(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) Then
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
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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

                    If row.Cells(Me.ColPONoDONo.Index).Value = row1.Cells(Me.ColPONoDONo.Index).Value And _
                    row.Cells(Me.ColSuppDONo.Index).Value = row1.Cells(Me.ColSuppDONo.Index).Value And _
                    row.Cells(Me.ColCodeNo.Index).Value = row1.Cells(Me.ColCodeNo.Index).Value And _
                    row.Cells(Me.ColDescription.Index).Value = row1.Cells(Me.ColDescription.Index).Value And _
                    CStr(row.Cells(Me.ColBatch.Index).Value) = CStr(row1.Cells(Me.ColBatch.Index).Value) And _
                    row.Cells(Me.ColLocation.Index).Value = row1.Cells(Me.ColLocation.Index).Value Then
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
                    If CStr(row.Cells(Me.ColPONoDONo.Index).Value) = CStr(row1.Cells(Me.ColPONoDONo.Index).Value) And _
                    CStr(row.Cells(Me.ColSuppDONo.Index).Value) = CStr(row1.Cells(Me.ColSuppDONo.Index).Value) And _
                    CStr(row.Cells(Me.ColCodeNo.Index).Value) = CStr(row1.Cells(Me.ColCodeNo.Index).Value) And _
                    CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
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
            'row.Cells(5).Value = CDbl(row.Cells(5).Value) + CDbl(row1.Cells(5).Value)
            row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(row1.Cells(Me.ColQtyMulti.Index).Value), ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
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
                    Dim Max_ As String = ObjCD.DigitGrouping(CStr(CDbl(JmlObjLama - JmlGudang) * CDbl(row.Cells(Me.ColRatioReference.Index).Value) / CDbl(row.Cells(Me.ColRatioDefault.Index).Value)))
                    MsgBox("Minimum value you must can enter is " + Max_ + " " + row.Cells(Me.ColUnitsCodeMulti.Index).Value + ".", _
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

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

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
                    Me.TxtQtyperBatchMulti.Focus()
                Else
                    If Me.TxtLocation.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocation.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocation.Text = i.PrimaryStorageArea
                        Me.TxtWarehouse.Text = Obj.GetWarehouseCode(Me.txtPrimaryLocation.Text)
                        Me.TxtLocation.Text = Obj.GetStorageAreaCode(Me.txtPrimaryLocation.Text)
                        Me.TxtQtyperBatchMulti.Focus()
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
                    Me.TxtQtyperBatchMulti.Focus()
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
                        Me.TxtQtyperBatchMulti.Focus()
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
                Me.TxtQtyperBatchMulti.Focus()
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
            If Me.txtPrimaryDO.Text <> "" And Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryDO.Text = -1 Or Me.txtPrimaryDO.Text = "" Then
                    Me.txtDONo.Focus()
                    MsgBox("To start entering item data, you must specify delivery order data first.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

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

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.GotFocus
        Try
            Me.TxtQtyperBatchDefault.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
        Try
            Me.TxtQtyperBatchDefault.BackColor = Me.TxtPrimaryCodeNo.BackColor
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
                If Me.txtPrimaryDO.Text = -1 Or Me.txtPrimaryDO.Text = "" Then
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

    Private Sub txtDONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDONo.GotFocus, TxtQtyperBatchMulti.GotFocus, TxtQtyperBatchDefault.GotFocus
        Dim ij As TextBox

        Try
            ij = TryCast(sender, TextBox)

            If IsNothing(ij) = False Then
                If ij.Enabled = True Then
                    If ij.ReadOnly = True Then
                        ij.BackColor = Me.TxtKodePihakAcuan.BackColor
                    Else
                        ij.BackColor = Me.BackColorAll
                    End If
                End If
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
                    Me.ObjDOItem.IsiDGV_Multi(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
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
                        Me.ObjDOItem.IsiDGV_Multi(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
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
                    Me.ObjDOItem.IsiDGV_Multi(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
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
                        Me.ObjDOItem.IsiDGV_Multi(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
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
                Me.ObjDOItem.IsiDGV_Multi(Me.dgvItemDO2, Me.txtPrimaryDO.Text)
                Me.FormatDGVDO()
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
        Dim IsMultiUnit As Boolean = False

        Try
            Me.dgvItemDO2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemDO2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItemDO2.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColDO2_Expiry.Index).Value) = False Then
                    row.Cells(Me.ColDO2_Expiry.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColDO2_Expiry.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColDO2_Expiry.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColDO2_Expiry.Index).Value = Format(CDate(row.Cells(Me.ColDO2_Expiry.Index).Value), "dd MMM yyyy")
                    End If
                End If

                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If IsNumeric(row.Cells(Me.ColDO2_QtySecondary.Index).Value) = True Then
                    If row.Cells(Me.ColDO2_QtySecondary.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                If row.Cells(Me.ColDO2_QtySecondary.Index).Value = "-1" Then
                    row.Cells(Me.ColDO2_QtySecondary.Index).Value = "-"
                    row.Cells(Me.ColDO2_SecondaryUnit.Index).Value = "-"
                End If

                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColDO2_BatchNo.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.DgvItemDO2.Columns(Me.ColDO2_QtySecondary.Index).Visible = False
                Me.DgvItemDO2.Columns(Me.ColDO2_SecondaryUnit.Index).Visible = False
            Else
                Me.DgvItemDO2.Columns(Me.ColDO2_QtySecondary.Index).Visible = True
                Me.DgvItemDO2.Columns(Me.ColDO2_SecondaryUnit.Index).Visible = True
            End If
            '3. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.DgvItemDO2.Columns(Me.ColDO2_BatchNo.Index).Visible = False
            Else
                Me.DgvItemDO2.Columns(Me.ColDO2_BatchNo.Index).Visible = True
            End If
            '4. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.DgvItemDO2.Columns(Me.ColDO2_Expiry.Index).Visible = False
            Else
                Me.DgvItemDO2.Columns(Me.ColDO2_Expiry.Index).Visible = True
            End If

            For Each RowDgv As DataGridViewRow In Me.dgvItemDO2.Rows
                With RowDgv
                    If .Cells(Me.ColDO2_PrimaryUnits.Index).Value <> .Cells(Me.ColDO2_PrimaryUnitMulti.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.ColDO2_TotalQty.Visible = False
                Me.ColDO2_Unit.Visible = False
                Me.ColDO2_RatioDesc.Visible = False
            Else
                Me.ColDO2_TotalQty.Visible = True
                Me.ColDO2_Unit.Visible = True
                Me.ColDO2_RatioDesc.Visible = True
            End If

            Me.dgvItemDO2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemDO2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

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

    Private Sub txtBatchNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.Leave
        Try
            Me.txtBatchNo.Text = Me.txtBatchNo.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                    Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'Dim ij As New FrmTransaksiStockInSalesReturn2_Print
        'Try
        '    If Me.TxtPrimaryStockIn.Text = -1 Then
        '        Me.btnFind.PerformClick()
        '    End If

        '    If Me.TxtPrimaryStockIn.Text <> -1 Then
        '        Dim objCetak As New CustomSMESS.ClsTransaksiStockInOut
        '        objCetak.CetakStockInSalesReturn(ij.CrvStockInSalesReturn, Me.TxtPrimaryStockIn.Text)
        '        ij.ShowDialog()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        '    ij.Dispose()
        'End Try
        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                Dim i As New FrmLaporanFormatStockInOut_Bundling

                i.primary = Me.TxtPrimaryStockIn.Text
                i.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInSalesReturn
                i.ShowDialog()

                i = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            Me.SetMultiUnits()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimarySO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimarySO.TextChanged
        Try
            Me.SetMultiUnits()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryDO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryDO.TextChanged
        Try
            Me.SetMultiUnits()
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

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.buffer = Me.TxtQtyperBatchMulti.Text
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyDown
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
                        Me.TxtQtyperBatchMulti.Focus()
                    End If
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

    Public Sub FromCutOff_StockInSalesReturn_Create(ByVal vPrimaryTransaksi As Double, _
    ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core, _
    ByVal vTblPrimarySO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryDO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable)
        Dim TblTemp As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim rowTemp As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateRow
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.objStockIn.ResetObject(Me.dgvItem)
            Me.objStockIn.TipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

            Me.objStockIn.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.TblItemCustom)
            Me.objStockIn.PrimaryDOSales = ObjCutOff_Core.GetPrimaryUpdate(Me.objStockIn.PrimaryDOSales, vTblPrimaryDO)
            Me.UserCode = Me.objStockIn.CreatedBy

            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimarySO.Index).Value) = False And rowD.Cells(Me.ColPrimarySO.Index).Value <> "-1" Then
                    TblTemp.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value)
                    If TblTemp.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), vTblPrimarySO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), PrimaryUpdate, TblTemp)
                        rowD.Cells(Me.ColPrimarySO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimarySO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), TblTemp)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_StockInSalesReturn_Save() As Double
        Try
            Me.baru = True
            Me.objStockIn.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, -1, False, False, True)
            Return Me.objStockIn.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub txtBatchNo_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles txtBatchNo.Layout

    End Sub

    Private Sub txtDONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDONo.LostFocus, TxtQtyperBatchMulti.LostFocus
        Dim ij As TextBox

        Try
            ij = TryCast(sender, TextBox)

            If IsNothing(ij) = False Then
                If ij.Enabled = True Then
                    If ij.ReadOnly = True Then
                        ij.BackColor = Me.TxtKodePihakAcuan.BackColor
                    Else
                        ij.BackColor = Me.txtPrimaryCategory.BackColor
                    End If
                End If
            End If
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
        Dim i As New CustomSMESS.FrmTransaksiStockInSalesReturn2_AddByCustomProduct

        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.txtPrimaryDO.Text = -1 And Me.txtPrimaryDO.Text <> "" Then
                    MsgBox("To start entering item data, you must specify delivery order data first.", MsgBoxStyle.Information, "Information")
                    Me.txtDONo.Focus()
                    Exit Sub
                End If
            End If

            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TblItemCustom = Me.TblItemCustom.Copy
            i.PrimaryTransaksi = Me.TxtPrimaryStockIn.Text
            i.TglTransaksi = Me.dtpDate.Value
            i.PrimaryDO = Me.txtPrimaryDO.Text
            i.baru = Me.baru
            i.vDgvItemDO = Me.DgvItemDO2
            i.BackColorAll = Me.BackColorAll
            'i.PrimaryPackingList = Me.txtPrimaryDO.Text
            If Me.TxtRemark.Enabled = False Then
                i.CanEditForm = False
            End If

            If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblItemCustom.Clear()
                Me.TblItemCustom = i.TblItemCustom.Copy
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
                    PrimaryLocationLama = Me.txtPrimaryLocation.Text
                    IdxRows = Me.GetIndex(.Primary_Product, .Primary_Storage_Area, .Batch_No, -1, .Primary_PdO, .Expiry_Date, .Primary_PO_Purchase)
                    'cari data yang sama persis
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

                        If String.IsNullOrEmpty(.Remarks) = False Then
                            If String.IsNullOrEmpty(rd.Cells(Me.ColRemark.Index).Value) = False Then
                                rd.Cells(Me.ColRemark.Index).Value = rd.Cells(Me.ColRemark.Index).Value & ", " & .Remarks
                            Else
                                rd.Cells(Me.ColRemark.Index).Value = .Remarks
                            End If
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
                        rd.Cells(Me.ColPONoDONo.Index).Value = .po_so_number
                        rd.Cells(Me.ColSuppDONo.Index).Value = "" '.Supp__DO_No
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Multi_Unit, .dec_multi)
                        If .Is_Kain = True Then
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Multi_Unit * .Qty_At, .dec)
                        Else
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, .dec)
                        End If
                        rd.Cells(Me.ColUnitsCodeMulti.Index).Value = .unit_multi
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
                        rd.Cells(Me.ColPrimarySO.Index).Value = .Primary_PO_Purchase
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                        rd.Cells(Me.ColPrimaryPdO.Index).Value = .Primary_PdO
                        rd.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Multi_Units
                        rd.Cells(Me.ColRatioDefault.Index).Value = .Ratio_Default
                        rd.Cells(Me.ColRatioReference.Index).Value = .Rasio
                    End If
                End With
            Next

            'format dgv
            Me.FormatDGV()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.dgvItem.RowCount > 0 Then
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
End Class