Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockTransfer
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public IsTransfer_PickList As Boolean = False

    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim TBLDetailLocation As New DseUserLocationPrivilegeSetting.tblmaster_storage_areaDataTable
    Dim ObjUser As New ClsUser
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim baru As Boolean
    Dim objStockTransfer As New clsTransaksiStockTransfer
    Dim ObjCD As New ClsCekData
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim objStockTransferItem As New clsTransaksiStockTransferItem
    Dim IsTransOK As Boolean = False
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL3 As New dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4 As New dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item1DataTable
    Dim TBL5 As New dseTransaksiStockTransfer.DataTable2DataTable
    Dim TBL6 As New dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item2DataTable

    Dim buffer As String = ""
    Dim QuickAddByRefDoc As Boolean = False
    Dim QuickAddByBarcode As Boolean = False

    'Terkait Transit
    Public ApakahTransaksiReceived As Boolean = False
    Public PrimaryRefDoc As Double = -1
    Public NoRefDocument As String = ""
    'Dim ApakahTransit As Boolean = False

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True)
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

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockPindah_QuickInputDocRef") = True Then
                Me.QuickAddByRefDoc = True
            Else
                Me.QuickAddByRefDoc = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockPindah_QuickInputBarcode") = True Then
                Me.QuickAddByBarcode = True
            Else
                Me.QuickAddByBarcode = False
            End If
            If Me.QuickAddByBarcode = True Or Me.QuickAddByRefDoc = True Then
                Me.llbAdvance.Enabled = True
            Else
                Me.llbAdvance.Enabled = False
            End If
            Me.Atur1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryStockIn.BackColor, False)
            Me.llbAdvance.Enabled = False
            'TR//
            Me.Atur1()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)

            'Terkait Transit
            'Me.ApakahTransit = False
            Me.TxtPrimaryRefDoc.Text = "-1"
            Me.chkUseTransit.Checked = False
            Me.ChkIsReceived.Checked = False
            Me.ChkIsReveivement.Checked = False
            Me.TxtTransitStatus.Text = ""

            Me.TxtPrimaryStockIn.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            'TR//
            Me.TblItemCustom.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnits.Text = "-1"
            Me.txtPrimarySecondaryUnits.Text = "-1"
            Me.txtPrimaryLocation.Text = "-1"
            Me.txtPrimaryDestination.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyperBatch.Text = "0"
            Me.txtSecondaryQty.Text = "-1"
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
            Me.txtQtyOnHand.Text = "0"

            ''transit location
            'Me.ColDestinationLoc.HeaderText = "Destination"
            'Me.ColWarehouseDestination.HeaderText = "Destination Warehouse"
            'Me.ColTransitLocation.Visible = False
            'Me.ColTransitWarehouse.Visible = False
            'Me.ColTransitLocation.HeaderText = "Transit Location"
            'Me.ColTransitWarehouse.HeaderText = "Transit Warehouse"
            'Me.ColQtyReceived.Visible = False
            'Me.ColUnitReceived.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar(ByVal PrimaryTransitLocInput As Double) As Boolean
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
            '1.b. hanya boleh masukkan product yang kontrol lokasi
            If Me.ObjProduct.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text) = False Then
                MsgBox("Please only enter goods with location controll.", MsgBoxStyle.Information, "Information")
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
            '5.b. Bila location controlled, kode lokasi tujuan tidak diisikan
            If Me.chkLocationControlled.Checked = True Then
                If Me.txtPrimaryDestination.Text = "-1" Then
                    MsgBox("Please enter transfer destination.", MsgBoxStyle.Information, _
                    "Information")
                    Me.txtDestinationCode.Focus()
                    Return False
                End If
            End If
            '5.c. Location tujuan dan asal tidak boleh sama
            If Me.txtPrimaryDestination.Text = Me.txtPrimaryLocation.Text Then
                MsgBox("Source dan destination location must be different.", MsgBoxStyle.Information, "Information")
                Return False
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
                    If row.Cells(Me.ColLocation.Index).Value = Me.cboBatchNo.Text And Me.cboBatchNo.Text <> "" And row.Cells(Me.ColPrimaryLocation.Index).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColWareHouse.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(Me.ColWareHouse.Index).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
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
            JmlObjLama = Me.objStockTransferItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
            Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.DtpExpired.Value)
            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
            Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, _
            Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
            JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatch.Text)
            If JmlSetelahTambah < 0 Then
                If Me.chkUseSecondaryUnit.Checked = True Then
                    MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGV) + " " + _
                    Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                    Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, _
                    "Information")
                Else
                    MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGV) + " " + _
                    CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                    "Information")
                End If
                Me.TxtQtyperBatch.Focus()
                Return False
            End If
            '9.b. Cek stock pada tanggal yang bersangkutan
            'msgbox("belum")

            'Terkait Transit
            If Me.CekBatasanInputTransit(PrimaryTransitLocInput) = False Then
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

    Private Function GetJmlDGVDestination(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryDestination As Double, ByVal vBatchNo As String, _
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
                        If vPrimaryDestination <> row.Cells(Me.ColPrimaryDestination.Index).Value Then
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

    Private Sub LoadLocation_To()
        Dim RowDetailLocation As DseUserLocationPrivilegeSetting.tblmaster_storage_areaRow
        Dim rowDGV As DataGridViewRow
        Dim ListPrimary As New List(Of Double)
        Dim PrimaryLocationTo As Double = -1

        Try
            'For Each rowDGV In Me.dgvItem.Rows
            '    If rowDGV.Cells(Me.ColPrimaryDestination.Index).Value <> -1 Then
            '        If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryDestination.Index).Value) = False Then
            '            RowDetailLocation = Me.TBLDetailLocation.Newtblmaster_storage_areaRow
            '            With RowDetailLocation
            '                .Primary = rowDGV.Cells(Me.ColPrimaryDestination.Index).Value
            '                .Storage_Area_Code = Me.ObjLocation.GetStorageAreaCode(.Primary)
            '                .Primary_Warehouse = Me.ObjLocation.GetPrimaryWarehouse(.Primary)
            '            End With
            '            Me.TBLDetailLocation.Rows.Add(RowDetailLocation)
            '            ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryDestination.Index).Value)
            '        End If
            '    End If
            'Next

            For Each rowDGV In Me.dgvItem.Rows
                If Me.chkUseTransit.Checked = False Then
                    PrimaryLocationTo = rowDGV.Cells(Me.ColPrimaryDestination.Index).Value
                Else
                    PrimaryLocationTo = rowDGV.Cells(Me.ColPrimaryTransitLocation.Index).Value
                End If
                If PrimaryLocationTo <> -1 Then
                    If ListPrimary.Contains(PrimaryLocationTo) = False Then
                        RowDetailLocation = Me.TBLDetailLocation.Newtblmaster_storage_areaRow
                        With RowDetailLocation
                            .Primary = PrimaryLocationTo
                            .Storage_Area_Code = Me.ObjLocation.GetStorageAreaCode(.Primary)
                            .Primary_Warehouse = Me.ObjLocation.GetPrimaryWarehouse(.Primary)
                        End With
                        Me.TBLDetailLocation.Rows.Add(RowDetailLocation)
                        ListPrimary.Add(PrimaryLocationTo)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadLocation_From()
        Dim RowDetailLocation As DseUserLocationPrivilegeSetting.tblmaster_storage_areaRow
        Dim rowDGV As DataGridViewRow
        Dim ListPrimary As New List(Of Double)

        Try
            'For Each rowDGV In Me.dgvItem.Rows
            '    If rowDGV.Cells(Me.ColPrimaryLocation.Index).Value <> -1 Then
            '        If ListPrimary.Contains(rowDGV.Cells(Me.ColPrimaryLocation.Index).Value) = False Then
            '            RowDetailLocation = Me.TBLDetailLocation.Newtblmaster_storage_areaRow
            '            With RowDetailLocation
            '                .Primary = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value
            '                .Storage_Area_Code = rowDGV.Cells(Me.ColLocation.Index).Value
            '                .Primary_Warehouse = Me.ObjLocation.GetPrimaryWarehouse(.Primary)
            '            End With
            '            Me.TBLDetailLocation.Rows.Add(RowDetailLocation)
            '            ListPrimary.Add(rowDGV.Cells(Me.ColPrimaryLocation.Index).Value)
            '        End If
            '    End If
            'Next

            If Me.ChkIsReveivement.Checked = False Then
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
            Else
                Dim DTA As New dseTransaksiStockTransferTableAdapters.transaksi_stocktransfer_viewTableAdapter
                Dim TBL As New dseTransaksiStockTransfer.transaksi_stocktransfer_viewDataTable
                Dim row As dseTransaksiStockTransfer.transaksi_stocktransfer_viewRow
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.Fill(TBL, Me.PrimaryRefDoc)
                For Each row In TBL
                    If row.Primary_Storage_Area <> -1 Then
                        If ListPrimary.Contains(row.Primary_Storage_Area) = False Then
                            RowDetailLocation = Me.TBLDetailLocation.Newtblmaster_storage_areaRow
                            With RowDetailLocation
                                .Primary = row.Primary_Storage_Area
                                .Storage_Area_Code = row.Storage_Area_Code
                                .Primary_Warehouse = Me.ObjLocation.GetPrimaryWarehouse(.Primary)
                            End With
                            Me.TBLDetailLocation.Rows.Add(RowDetailLocation)
                            ListPrimary.Add(row.Primary_Storage_Area)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim rowDGV As DataGridViewRow
        Dim ConvertionDate As Date = #1/1/1950#
        Dim ObjAcc As New Core_Module.ClsTranskasiAccClosing

        Try
            ConvertionDate = ObjAcc.GetSMESSConvertionDate
            If ConvertionDate.Date > Me.dtpDate.Value.Date Then
                MsgBox("Minimum transaction date is on SMESS convertion date: " + Format(ConvertionDate, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Me.dtpDate.Focus()
                Return False
            End If

            'For Each rowDGV In Me.dgvItem.Rows
            '    rowDGV.Cells(Me.ColSecondaryQty.Index).Value = -1
            'Next
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("stock transfer number")
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
            'tidak boleh double
            If Me.baru = False Then
                If Me.objStockTransfer.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objStockTransfer.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtStockInNo.Text = Me.objStockTransfer.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objStockTransfer.TanggalTransaksi
                        Me.TxtStockInNo.Text = Me.objStockTransfer.NoTransaksi
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.objStockTransfer.GeTimestamp(Me.TxtPrimaryStockIn.Text)
                If Me.objStockTransfer.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objStockTransfer.CariDetail(Me.objStockTransfer.Primary, Me.dgvItem, Me.ChkIsReveivement.Checked, Me.TblItemCustom)
                        Me.SalinKeField()
                    End If
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
            'If objLocationAccess.HasAccessPrivileges(Me.UserCode, listOfPrimaryLocation, FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut.STOCK_TANSFER_OUT) = False Then
            '    Return False
            'End If
            '''------------------
            'listOfPrimaryLocation.Clear()
            'For Each dataGridViewRow As DataGridViewRow In Me.dgvItem.Rows
            '    If dataGridViewRow.Cells(Me.PrimaryDestination.Index).Value <> "-1" Then
            '        If listOfPrimaryLocation.Contains(dataGridViewRow.Cells(Me.PrimaryDestination.Index).Value) = False Then listOfPrimaryLocation.Add(dataGridViewRow.Cells(Me.PrimaryDestination.Index).Value)
            '    End If
            'Next
            'If objLocationAccess.HasAccessPrivileges(Me.UserCode, listOfPrimaryLocation, FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut.STOCK_TRANSFER_IN) = False Then
            '    Return False
            'End If

            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation_From()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_TANSFER_OUT, Me.TBLDetailLocation) = False Then
                    Return False
                End If
            End If
            Me.TBLDetailLocation.Rows.Clear()
            Me.LoadLocation_To()
            If Me.TBLDetailLocation.Rows.Count > 0 Then
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode, ClsUser.EnumLocationAccessType.STOCK_TRANSFER_IN, Me.TBLDetailLocation) = False Then
                    Return False
                End If
            End If

            If Me.objStockTransfer.CekBatasanMinusQty(Me.TxtPrimaryStockIn.Text, Me.dgvItem, True) = False Then
                Return False
            End If

            'Terkait Transit
            If Me.ChkIsReceived.Checked = True Then
                MsgBox("This transaction has been received, you can't edit this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '5. tidak bisa receive jika sudah dibuat transaksi receive (khusus transaksi receivement)
            If Me.ChkIsReveivement.Checked = True Then
                If Me.objStockTransfer.ApakahSudahDibuatReceive(Me.TxtPrimaryRefDoc.Text, Me.TxtPrimaryStockIn.Text) = True Then
                    MsgBox("Transaction stock transfer " & Me.NoRefDocument.ToString & " has been received", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'Dim rd As DataGridViewRow
            'Dim PesanTidakBisaSave As String = ""
            'If Me.chkUseTransit.Checked = False Then
            '    'jika tidak transit maka col transit (paling kanan) harus kosong
            '    For Each rd In Me.dgvItem.Rows
            '        If String.IsNullOrEmpty(rd.Cells(Me.ColTransitLocation.Index).Value) = False Then
            '            If PesanTidakBisaSave = "" Then
            '                PesanTidakBisaSave = rd.Cells(Me.ColNo.Index).Value
            '            Else
            '                PesanTidakBisaSave &= ", " & rd.Cells(Me.ColNo.Index).Value
            '            End If
            '        End If
            '    Next
            'Else
            '    'jika transit maka col transit (tengah) harus sama semua lokasi nya
            '    For Each rd In Me.dgvItem.Rows
            '        If String.IsNullOrEmpty(rd.Cells(Me.ColTransitLocation.Index).Value) = True Then
            '            If PesanTidakBisaSave = "" Then
            '                PesanTidakBisaSave = rd.Cells(Me.ColNo.Index).Value
            '            Else
            '                PesanTidakBisaSave &= ", " & rd.Cells(Me.ColNo.Index).Value
            '            End If
            '        End If
            '    Next
            'End If

            Dim ObjCust As New CustomSMESS.clsTransaksiStockTransfer
            Dim rd As DataGridViewRow
            Dim TblDummy As New CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable
            TblDummy = Me.TblItemCustom.Copy
            'If TblDummy.Rows.Count > 0 Then
            For i As Integer = 0 To Me.dgvItem.RowCount - 1
                rd = Me.dgvItem.Rows(i)
                With rd
                    Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                    Dim PrimaryLocationDestination As Double = .Cells(Me.ColPrimaryDestination.Index).Value
                    Dim Batch As String = .Cells(Me.ColBatch.Index).Value
                    Dim QtyDef As Decimal = .Cells(Me.ColQty.Index).Value
                    Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                    If ObjCust.CompareQtySMESS_Custom(PrimaryProduct, PrimaryLocation, PrimaryLocationDestination, Batch, QtyDef, TblDummy) = False Then
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
            ObjAcc = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objStockTransfer
                .NoTransaksi = Me.TxtStockInNo.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .Remarks = Me.TxtRemark.Text

                'Terkait Transit
                If Me.chkUseTransit.Checked = False Then
                    .UseTransit = 0
                    .TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit
                Else
                    .UseTransit = 1
                    .TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.On_Transit
                End If

                If Me.ChkIsReceived.Checked = True Then
                    .TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit
                End If

                If Me.ChkIsReveivement.Checked = False Then
                    .IsReceivement = 0
                Else
                    .IsReceivement = 1
                    .TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit
                End If
                .PrimaryRefDoc = Me.TxtPrimaryRefDoc.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objStockTransfer
                Me.TxtStockInNo.Text = .NoTransaksi
                Me.dtpDate.Value = .TanggalTransaksi
                Me.TxtRemark.Text = .Remarks
                Me.TxtPrimaryStockIn.Text = .Primary

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                'Terkait Transit
                If .UseTransit = 0 Then
                    Me.chkUseTransit.Checked = False
                Else
                    Me.chkUseTransit.Checked = True
                End If
                Me.ChkIsReceived.Checked = False
                If .IsReceivement = 0 Then
                    Me.ChkIsReveivement.Checked = False
                    If .TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit Then
                        Me.TxtTransitStatus.Text = "-"
                    ElseIf .TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.On_Transit Then
                        Me.TxtTransitStatus.Text = "On Transit"
                    Else
                        Me.ChkIsReceived.Checked = True
                        Me.TxtTransitStatus.Text = "Received on " & Format(.ReceivedDate, "dd MMM yyyy") & " (User ID : " & .ReceivedBy & " on " & Format(.ReceivedTimestamp, "dd MMM yyyy") + "-" + .ReceivedTimestamp.ToShortTimeString & ")"
                    End If
                Else
                    Me.ChkIsReveivement.Checked = True
                    Me.TxtTransitStatus.Text = "Reference stock transfer no : " & Me.objStockTransfer.GetNoTransaksi(.PrimaryRefDoc)
                End If
                Me.TxtPrimaryRefDoc.Text = .PrimaryRefDoc
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

            Me.SetTransit_InDgv(Me.ChkIsReveivement.Checked)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Terkait Transit
    Private Sub SetTransitStatus(ByVal vApakahTransaksiReceived As Boolean)
        Dim ApakahTransit As Boolean = False

        Try
            If Me.dgvItem.RowCount = 1 Then
                Dim rd As DataGridViewRow = Me.dgvItem.Rows(0)
                If String.IsNullOrEmpty(rd.Cells(Me.ColPrimaryTransitLocation.Index).Value) = False Then
                    If rd.Cells(Me.ColPrimaryTransitLocation.Index).Value <> "-1" Then
                        ApakahTransit = True
                    End If
                End If
            ElseIf Me.dgvItem.RowCount > 1 Then
                Exit Sub
            End If

            If ApakahTransit = False Then
                Me.chkUseTransit.Checked = False
                Me.TxtTransitStatus.Text = "-"
            Else
                Me.chkUseTransit.Checked = True
                Me.TxtTransitStatus.Text = "On Transit"
            End If

            Me.SetTransit_InDgv(vApakahTransaksiReceived)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Terkait Transit
    Private Sub SetTransit_InDgv(ByVal vApakahTransaksiReceived As Boolean)

        Try
            If vApakahTransaksiReceived = False Then
                Me.ColLocation.HeaderText = "Location"
                Me.ColWareHouse.HeaderText = "Warehouse"
                If Me.chkUseTransit.Checked = False Then
                    Me.ColDestinationLoc.HeaderText = "Destination"
                    Me.ColWarehouseDestination.HeaderText = "Destination Warehouse"
                    Me.ColTransitLocation.Visible = False
                    Me.ColTransitWarehouse.Visible = False
                    Me.ColTransitLocation.HeaderText = "Transit Location"
                    Me.ColTransitWarehouse.HeaderText = "Transit Warehouse"
                    Me.ColQtyReceived.Visible = False
                    Me.ColUnitReceived.Visible = False
                Else
                    Me.ColDestinationLoc.HeaderText = "Transit Location"
                    Me.ColWarehouseDestination.HeaderText = "Transit Warehouse"
                    Me.ColTransitLocation.Visible = True
                    Me.ColTransitWarehouse.Visible = True
                    Me.ColTransitLocation.HeaderText = "Destination"
                    Me.ColTransitWarehouse.HeaderText = "Destination Warehouse"
                    Me.ColQtyReceived.Visible = True
                    Me.ColUnitReceived.Visible = True
                End If
            Else
                Me.ColLocation.HeaderText = "Transit Location"
                Me.ColWareHouse.HeaderText = "Transit Warehouse"
                Me.ColDestinationLoc.HeaderText = "Destination"
                Me.ColWarehouseDestination.HeaderText = "Destination Warehouse"
                Me.ColTransitLocation.Visible = False
                Me.ColTransitWarehouse.Visible = False
                Me.ColQtyReceived.Visible = False
                Me.ColUnitReceived.Visible = False
            End If
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
                        If Me.objStockTransfer.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                            If IsTransfer_PickList = True Then
                                Me.IsTransOK = True
                                Me.Close()
                            End If
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockTransfer.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If

                If Me.ApakahTransaksiReceived = True Then
                    Me.Close()
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
            If IsTransfer_PickList = True Then
                Me.Close()
            Else
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
                        If Me.ApakahTransaksiReceived = True Then
                            Me.Close()
                        End If

                        Me.KosongkanItem()
                        If baru = True Then
                            Me.KosongkanField()
                        Else
                            Me.objStockTransfer.CariDetail(Me.TxtPrimaryStockIn.Text, Me.dgvItem, Me.ChkIsReveivement.Checked, Me.TblItemCustom)
                            Me.SalinKeField()
                        End If
                    Else
                        Exit Sub
                    End If
                    Me.NonAktifkanDetail()
                    Me.btnNew.Focus()
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
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.dtpDate.Focus()
            Me.objStockTransfer.ResetObject(Me.dgvItem)
            Me.TxtStockInNo.Text = Me.objStockTransfer.GetNoBaru(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiStockTransfer
        Try
            hj.ShowDialog()
            If Not hj.PrimaryTransaksiStockTransfer = -1 Then
                If Me.objStockTransfer.CariDetail(hj.PrimaryTransaksiStockTransfer, Me.dgvItem, Me.ChkIsReveivement.Checked, Me.TblItemCustom) = True Then
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
                If Me.objStockTransfer.ApakahBolehEdit(Me.UserCode, True) = True Then
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

        'by custom admiral
        Dim TblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ListProduct As New List(Of Double)

        Try
            If Me.TxtStockInNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtStockInNo.Text <> "" Then
                If Me.objStockTransfer.ApakahBolehHapus(Me.UserCode, True, TblProdQty, ListProduct) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objStockTransfer.Hapus(Me.TxtPrimaryStockIn.Text, Me.UserCode, True) = True Then
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
            If Me.ApakahTransaksiReceived = False Then
                If IsTransfer_PickList = False Then
                    If Me.TxtRemark.Enabled = True Then
                        Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ApakahDataBenar() = True Then
                                Me.SalinKeObjek()
                                If baru = True Then
                                    'simpan baru
                                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                    If Me.objStockTransfer.SimpanBaru(Me.dgvItem, Me.UserCode, Me.TblItemCustom) = True Then
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
                                    If Me.objStockTransfer.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) = True Then
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
            Dim row1 As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item1Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockTransferItem.IsikanTBLBatchNo(Me.TBL4, vPrimaryProduct, vPrimaryTransaksi)
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
            Me.TBL.DefaultView.Sort = "[Lot No] ASC"
            Me.cboBatchNo.DataSource = Me.TBL.DefaultView
            Me.cboBatchNo.ValueMember = "Lot No"
            Me.cboBatchNo.DisplayMember = "Lot No"
            Me.cboBatchNo.Refresh()

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
            Dim row1 As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item1Row
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
            Dim row1 As dseTransaksiStockTransfer.DataTable2Row
            Dim row2 As dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockTransferItem.IsikanTBLLocation(Me.TBL5, vPrimaryProduct, vPrimaryTransaksi, vBatchNo)
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLLocation(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As dseTransaksiStockTransfer.DataTable2Row
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
            Dim row1 As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item2Row
            Dim row2 As dseMasterProduct.tblmaster_productqty2Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockTransferItem.IsikanTBLSecondaryUnit(Me.TBL6, vPrimaryProduct, _
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
            Dim row1 As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item2Row
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
                'Me.lblUnits.Left = Me.lblSecQty.Left - 37
                'Me.TxtQtyperBatch.Width = Me.lblUnits.Left - Me.TxtQtyperBatch.Left - 6
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
        i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Location_Control_Only
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct
        i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Location_Control_Only

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
        Dim i As New frmCariMasterProduct
        i.TipePencarian = frmCariMasterProduct.enumTipePencarian.Location_Control_Only

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
    ''' <param name="ExpiryDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Double, ByVal vPrimaryDestination As Double, ByVal ExpiryDate As Date, ByVal vPrimaryTransit As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsDate(row.Cells(Me.ColExpiredDate.Index).Value) = False Then
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And vPrimaryDestination = row.Cells(Me.ColPrimaryDestination.Index).Value And _
                        row.Cells(Me.ColPrimaryTransitLocation.Index).Value = vPrimaryTransit Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And vPrimaryDestination = row.Cells(Me.ColPrimaryDestination.Index).Value And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And ExpiryDate = #1/1/1950# And row.Cells(Me.ColPrimaryTransitLocation.Index).Value = vPrimaryTransit Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And vPrimaryDestination = row.Cells(Me.ColPrimaryDestination.Index).Value And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryTransitLocation.Index).Value = vPrimaryTransit Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And vPrimaryDestination = row.Cells(Me.ColPrimaryDestination.Index).Value And _
                        row.Cells(Me.ColBatch.Index).Value = BatchNo And row.Cells(Me.ColSecondaryQty.Index).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(Me.ColExpiredDate.Index).Value) And row.Cells(Me.ColPrimaryTransitLocation.Index).Value = vPrimaryTransit Then
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

    'Terkait Transit
    Private Function GetTransitData(ByRef vPrimaryTransitLoc As Double, ByRef vTransitLoc As String, ByRef vTransitWarehouse As String, ByVal vPrimaryLocationAsal As Double) As Boolean
        Try
            If String.IsNullOrEmpty(vPrimaryLocationAsal) = False Then
                If IsNumeric(vPrimaryLocationAsal) = True Then
                    If Me.ObjLocation.GetTransitData(vPrimaryLocationAsal, vPrimaryTransitLoc, vTransitLoc, vTransitWarehouse) = True Then
                        Return True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim Input_PrimaryDestination, Input_PrimaryTransit As Double
        Dim Input_Destination, Input_Transit, Input_DestinationWarehouse, Input_TransitWarehouse As String
        Dim QtyReceived As Decimal = 0
        Dim UnitReceived As String = Me.lblUnits.Text
        Dim DestinationWarehouse As String = Me.ObjLocation.GetWarehouseCode(Me.txtPrimaryDestination.Text)


        Dim PrimaryTransitLocation As Double = -1
        Dim TransitLocation As String = ""
        Dim TransitWarehouse As String = ""

        Try
            'Terkait Transit
            Me.GetTransitData(PrimaryTransitLocation, TransitLocation, TransitWarehouse, Me.txtPrimaryLocation.Text)

            'Jika warehouse lokasi asal sama dengan warehouse lokasi tujuan maka tidak perlu transit
            If PrimaryTransitLocation <> -1 Then
                If Me.TxtWarehouse.Text = DestinationWarehouse Then
                    PrimaryTransitLocation = -1
                End If
            End If

            If Me.ApakahDataItemBenar(PrimaryTransitLocation) = True Then
                'Terkait Transit
               If PrimaryTransitLocation = -1 Then
                    Input_PrimaryDestination = Me.txtPrimaryDestination.Text
                    Input_Destination = Me.txtDestinationCode.Text
                    Input_DestinationWarehouse = DestinationWarehouse
                    Input_PrimaryTransit = -1
                    Input_Transit = ""
                    Input_TransitWarehouse = ""
                Else
                    Input_PrimaryDestination = PrimaryTransitLocation
                    Input_Destination = TransitLocation
                    Input_DestinationWarehouse = TransitWarehouse
                    Input_PrimaryTransit = Me.txtPrimaryDestination.Text
                    Input_Transit = Me.txtDestinationCode.Text
                    Input_TransitWarehouse = DestinationWarehouse
                End If
                '/////////////Terkait Transit////////////////////

                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, _
                Me.txtSecondaryQty.Text, Input_PrimaryDestination, Me.DtpExpired.Value, Input_PrimaryTransit)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + CDbl(Me.TxtQtyperBatch.Text), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpired.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpired.Value.Date, "dd MMM yyyy")
                    End If

                    Dim rd As DataGridViewRow
                    Me.dgvItem.Rows.Add()
                    rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                    rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                    rd.Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                    rd.Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                    rd.Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatch.Text
                    rd.Cells(Me.ColUnit.Index).Value = Me.lblUnits.Text
                    rd.Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                    rd.Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                    rd.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text
                    rd.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                    rd.Cells(Me.ColLocation.Index).Value = Me.cboLocation.Text
                    rd.Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                    rd.Cells(Me.ColDestinationLoc.Index).Value = Input_Destination
                    rd.Cells(Me.ColWarehouseDestination.Index).Value = Input_DestinationWarehouse
                    rd.Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                    rd.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                    rd.Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                    rd.Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnits.Text
                    rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                    rd.Cells(Me.ColPrimaryDestination.Index).Value = Input_PrimaryDestination
                    rd.Cells(Me.ColTransitLocation.Index).Value = Input_Transit
                    rd.Cells(Me.ColTransitWarehouse.Index).Value = Input_TransitWarehouse
                    rd.Cells(Me.ColQtyReceived.Index).Value = QtyReceived
                    rd.Cells(Me.ColUnitReceived.Index).Value = UnitReceived
                    rd.Cells(Me.ColPrimaryTransitLocation.Index).Value = Input_PrimaryTransit

                    ''set ulang nomor
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

                'Terkait Transit
                Me.SetTransitStatus(Me.ChkIsReveivement.Checked)
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
                row.Cells(Me.ColSecondaryQty.Index).Value = -1
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
                    If row.Cells(Me.ColSecondaryQty.Index).Value > "0" Then
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
                Me.ColSecondaryQty.Visible = False
                Me.ColSecondaryUnit.Visible = False
            Else
                Me.ColSecondaryQty.Visible = True
                Me.ColSecondaryUnit.Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.ColLocation.Visible = False
                Me.ColWareHouse.Visible = False
            Else
                Me.ColLocation.Visible = True
                Me.ColWareHouse.Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.ColBatch.Visible = False
            Else
                Me.ColBatch.Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.ColExpiredDate.Visible = False
            Else
                Me.ColExpiredDate.Visible = True
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
            If Me.ChkIsReveivement.Checked = True Then
                If e.ColumnIndex = Me.ColRemark.Index Then
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    MsgBox("You can't edit this row.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
                
                Exit Sub
            End If
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Location Code
            '5. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.TxtRemark.Enabled = True Then
                If e.ColumnIndex = Me.ColQty.Index Or e.ColumnIndex = Me.ColSecondaryQty.Index Or e.ColumnIndex = Me.ColBatch.Index Or e.ColumnIndex = Me.ColExpiredDate.Index _
                Or e.ColumnIndex = Me.ColDestinationLoc.Index Or e.ColumnIndex = Me.ColRemark.Index Or e.ColumnIndex = Me.ColTransitLocation.Index Then
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
                    Else
                        If e.ColumnIndex = Me.ColDestinationLoc.Index Then
                            If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColTransitLocation.Index, e.RowIndex).Value) = False Then
                                e.Cancel = True
                                Exit Sub
                            End If
                        ElseIf e.ColumnIndex = Me.ColTransitLocation.Index Then
                            If String.IsNullOrEmpty(Me.dgvItem.Item(Me.ColTransitLocation.Index, e.RowIndex).Value) = True Then
                                e.Cancel = True
                                Exit Sub
                            End If
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

                        If e.ColumnIndex = Me.ColQty.Index Then
                            Me.EndEditQty(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColSecondaryQty.Index Then
                            Me.EndEditSecondaryUnit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColBatch.Index Then
                            Me.EndEditBatchNo(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColLocation.Index Then
                            Me.EndEditLocation(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColDestinationLoc.Index Then
                            Me.EndEditLocationDestination(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = Me.ColTransitLocation.Index Then
                            Me.EndEditLocationTransit(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        End If
                        Me.SetQtyOnHand()
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
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
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
                JmlDGV = .Cells(Me.ColQty.Index).Value
                'gudang asal
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek gudang tujuan
                    If Me.ApakahEditSahDestination(row, .Cells(Me.ColPrimaryDestination.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
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
                'cek gudang asal
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatch.Index).Value, ExpiryDate, .Cells(Me.ColSecondaryQty.Index).Value, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek gudang tujuan
                    If Me.ApakahEditSahDestination(row, .Cells(Me.ColPrimaryDestination.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    .Cells(Me.ColBatch.Index).Value, ExpiryDate, .Cells(Me.ColSecondaryQty.Index).Value, JmlDGV) = False Then
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
                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'batch no, cek gudang asal jangan sampe minus
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek gudang tujuan
                    If Me.ApakahEditSahDestination(row, .Cells(Me.ColPrimaryDestination.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
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
                                        BufferExpiryDate = CDate(Me.dgvItem.Item(Me.ColExpiredDate.Index, e.RowIndex).Value)
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
                .Cells(Me.ColPrimaryLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                'update data gudang
                .Cells(Me.ColWareHouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)

                'cek jangan sampe location tujuan sama dengan location awal
                If .Cells(Me.ColPrimaryLocation.Index).Value = .Cells(Me.ColPrimaryDestination.Index).Value Then
                    MsgBox("Source dan destination location must be different.", MsgBoxStyle.Information, "Information")
                    'kembalikan ke data semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    .Cells(Me.ColPrimaryLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColLocation.Index).Value)
                    'update data gudang
                    .Cells(Me.ColWareHouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryLocation.Index).Value)
                    Exit Sub
                End If

                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSah(row, .Cells(Me.ColPrimaryLocation.Index).Value, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
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
                            'kembalikan ke data semula
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

    Private Sub EndEditLocationDestination(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim Obj As New clsMasterStorageArea
        Dim PrimaryDestinationLama As Double
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
                PrimaryDestinationLama = .Cells(Me.ColPrimaryDestination.Index).Value
                .Cells(Me.ColPrimaryDestination.Index).Value = Obj.GetPrimary(.Cells(Me.ColDestinationLoc.Index).Value)
                'update data gudang
                .Cells(Me.ColWarehouseDestination.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryDestination.Index).Value)

                'cek jangan sampe location tujuan sama dengan location awal
                If .Cells(Me.ColPrimaryLocation.Index).Value = .Cells(Me.ColPrimaryDestination.Index).Value Then
                    MsgBox("Source dan destination location must be different.", MsgBoxStyle.Information, "Information")
                    'kembalikan ke data semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    .Cells(Me.ColPrimaryDestination.Index).Value = Obj.GetPrimary(.Cells(Me.ColDestinationLoc.Index).Value)
                    'update data gudang
                    .Cells(Me.ColWarehouseDestination.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryDestination.Index).Value)
                    Exit Sub
                End If

                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSahDestination(row, PrimaryDestinationLama, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    'kembalikan ke data semula
                    .Cells(Me.ColPrimaryDestination.Index).Value = Obj.GetPrimary(.Cells(Me.ColDestinationLoc.Index).Value)
                    'update data gudang
                    .Cells(Me.ColWarehouseDestination.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryDestination.Index).Value)
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
        Finally
            Obj = Nothing
        End Try
    End Sub

    'Terkait Transit
    Private Sub EndEditLocationTransit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Double, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim Obj As New clsMasterStorageArea
        Dim PrimaryTransitLama As Double
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
                PrimaryTransitLama = .Cells(Me.ColPrimaryTransitLocation.Index).Value
                .Cells(Me.ColPrimaryTransitLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColTransitLocation.Index).Value)
                'update data gudang
                .Cells(Me.ColTransitWarehouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryTransitLocation.Index).Value)

                'cek jangan sampe location tujuan sama dengan location awal
                If .Cells(Me.ColPrimaryLocation.Index).Value = .Cells(Me.ColPrimaryTransitLocation.Index).Value Then
                    MsgBox("Source dan destination location must be different.", MsgBoxStyle.Information, "Information")
                    'kembalikan ke data semula
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    .Cells(Me.ColPrimaryTransitLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColTransitLocation.Index).Value)
                    'update data gudang
                    .Cells(Me.ColTransitWarehouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryTransitLocation.Index).Value)
                    Exit Sub
                End If

                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSahDestination(row, PrimaryTransitLama, -1, .Cells(Me.ColPrimaryCodeNo.Index).Value, _
                .Cells(Me.ColBatch.Index).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    'kembalikan ke data semula
                    .Cells(Me.ColPrimaryTransitLocation.Index).Value = Obj.GetPrimary(.Cells(Me.ColTransitLocation.Index).Value)
                    'update data gudang
                    .Cells(Me.ColTransitWarehouse.Index).Value = Obj.GetWarehouseCode(.Cells(Me.ColPrimaryTransitLocation.Index).Value)
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

                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = row1.Cells(Me.ColPrimaryCodeNo.Index).Value And row.Cells(Me.ColDescription.Index).Value = row1.Cells(Me.ColDescription.Index).Value And _
                    row.Cells(Me.ColUnit.Index).Value = row1.Cells(Me.ColUnit.Index).Value And _
                    CStr(row.Cells(Me.ColBatch.Index).Value) = CStr(row1.Cells(Me.ColBatch.Index).Value) And CStr(row.Cells(Me.ColPrimaryDestination.Index).Value) = _
                    CStr(row1.Cells(Me.ColPrimaryDestination.Index).Value) And CStr(row.Cells(Me.ColPrimaryLocation.Index).Value) = CStr(row1.Cells(Me.ColPrimaryLocation.Index).Value) And _
                    CStr(row.Cells(Me.ColPrimaryTransitLocation.Index).Value) = CStr(row1.Cells(Me.ColPrimaryTransitLocation.Index).Value) Then
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
                    If CStr(row.Cells(Me.ColPrimaryCodeNo.Index).Value) = CStr(row1.Cells(Me.ColPrimaryCodeNo.Index).Value) And CStr(row.Cells(Me.ColDescription.Index).Value) = CStr(row1.Cells(Me.ColDescription.Index).Value) And _
                    CStr(row.Cells(Me.ColBatch.Index).Value) = CStr(row1.Cells(Me.ColBatch.Index).Value) And CStr(row.Cells(Me.ColUnit.Index).Value) = CStr(row1.Cells(Me.ColUnit.Index).Value) And _
                    CStr(row.Cells(Me.ColLocation.Index).Value) = CStr(row1.Cells(Me.ColLocation.Index).Value) And _
                    CStr(row.Cells(Me.ColDestinationLoc.Index).Value) = CStr(row1.Cells(Me.ColDestinationLoc.Index).Value) And _
                    CStr(row.Cells(Me.ColTransitLocation.Index).Value) = CStr(row1.Cells(Me.ColTransitLocation.Index).Value) Then
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
            row.Cells(Me.ColQty.Index).Value = CDbl(row.Cells(Me.ColQty.Index).Value) + CDbl(row1.Cells(Me.ColQty.Index).Value)
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
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Double, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Double
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockTransferItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
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
                    MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGVLain) + " " + _
                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            'cek stock untuk tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(PrimaryProduct, Me.dtpDate.Value)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockTransferItem.GetJml(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
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
                    MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + CStr(JmlGudang + JmlObjLama - JmlDGVLain) + " " + _
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

    Private Function ApakahEditSahDestination(ByVal row As DataGridViewRow, ByVal PrimaryDestination As Double, _
     ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
     ByVal ExpiryDate As Date, ByVal SecondaryQty As Double, ByVal JmlDGV As Double, _
     Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Double
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryDestination, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockTransferItem.GetJmlDestination(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryDestination, BatchNo, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGVDestination(PrimaryProduct, SecondaryQty, PrimaryDestination, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang - JmlObjLama + JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Minimum value you must enter is " + CStr(JmlObjLama - JmlGudang + JmlDGV - JmlDGVTotal) + " " + _
                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            'cek stock untuk tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(PrimaryProduct, Me.dtpDate.Value)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockTransferItem.GetJmlDestination(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
            SecondaryQty, PrimaryDestination, BatchNo, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGVDestination(PrimaryProduct, SecondaryQty, PrimaryDestination, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary)
            JmlSetelahEdit = JmlGudang - JmlObjLama + JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + "  minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Minimum value you must enter is " + CStr(JmlObjLama - JmlGudang + JmlDGV - JmlDGVTotal) + " " + _
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
            Me.SetQtyOnHand()

            'Terkait Transit
            Me.SetTransitStatus(Me.ChkIsReveivement.Checked)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            If Me.ChkIsReveivement.Checked = True Then
                MsgBox("You can't delete this row.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If
            'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
            Dim PrimaryLocation, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
            JmlDGV, JmlSetelahHapus As Double
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = e.Row
            Dim BatchNoCont, LocationCont, ExpiryDateCont, UseSecondary As Boolean

            With row
                PrimaryLocation = .Cells(Me.ColPrimaryDestination.Index).Value
                PrimaryProduct = .Cells(Me.ColPrimaryCodeNo.Index).Value
                Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, _
                ExpiryDateCont, UseSecondary)

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
                'karena akan dihapus, perhitungan menggunakan nilai jumlah dgv = total - yang akan dihapuskan
                JmlDGV = Me.GetJmlDGVDestination(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
                BatchNoCont, LocationCont, UseSecondary) - CDbl(Me.dgvItem.Item(Me.ColQty.Index, e.Row.Index).Value)
                JmlObjLama = Me.objStockTransferItem.GetJmlDestination(Me.TxtPrimaryStockIn.Text, PrimaryProduct, _
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
                Me.SetQtyOnHand()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objLoc = Nothing
        End Try
    End Sub

    Sub SetQtyOnHand()
        Dim qtyOnHand As Decimal = 0
        Dim qtyStock As Decimal = 0
        Dim qtyDBFrom As Decimal = 0
        Dim qtyDgvFrom As Decimal = 0
        Dim qtyDBTo As Decimal = 0
        Dim qtyDgvTo As Decimal = 0

        Try
            If Me.TxtPrimaryCodeNo.Text = "-1" Then
                Me.txtQtyOnHand.Text = Me.ObjCD.DigitGrouping(qtyOnHand, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
            End If

            qtyStock = Me.objStockTransfer.GetQty(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
            qtyDBFrom = Me.objStockTransfer.GetQtyDB(True, Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
            For Each r As DataGridViewRow In Me.dgvItem.Rows
                If r.Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text And r.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And r.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text Then
                    qtyDgvFrom += r.Cells(Me.ColQty.Index).Value
                End If
            Next
            qtyOnHand = qtyStock + qtyDBFrom - qtyDgvFrom
            qtyDBTo = Me.objStockTransfer.GetQtyDB(False, Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
            Dim PrimaryLokasiTujuan As Double = -1
            For Each r As DataGridViewRow In Me.dgvItem.Rows
                PrimaryLokasiTujuan = r.Cells(Me.ColPrimaryDestination.Index).Value
                If PrimaryLokasiTujuan = Me.txtPrimaryLocation.Text And r.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text And r.Cells(Me.ColBatch.Index).Value = Me.cboBatchNo.Text Then
                    qtyDgvTo += r.Cells(Me.ColQty.Index).Value
                End If
            Next
            qtyOnHand = qtyOnHand - qtyDBTo + qtyDgvTo
            Me.txtQtyOnHand.Text = Me.ObjCD.DigitGrouping(qtyOnHand, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub FrmTransaksiStockPindah_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'form ini hanya bisa digunakan jika ada minimal ada 1 product yang location 
        'controlled dan minimal 2 storage area

        Try
            If Me.IsTransfer_PickList = True Then
                Me.baru = True
                Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
                Exit Sub
            End If
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.KosongkanField()

            'betulin posisi txt remarks header
            Me.TxtRemark.Width = Me.cboLocation.Width
            Me.TxtRemark.Left = Me.TxtRemarkItem.Left + Me.gboItem.Left
            Me.ShowIcon = False

            'Terkait Transit
            If Me.ApakahTransaksiReceived = True Then
                Me.ProgressBar1.Visible = True
                Me.LblProgress.Visible = True
                Me.BtnStop.Visible = True
                Me.ProgressBar1.Anchor = AnchorStyles.Bottom + AnchorStyles.Right
                Me.LblProgress.Anchor = AnchorStyles.Bottom + AnchorStyles.Right

                Me.CreateTransaksiReceived(Me.PrimaryRefDoc, Me.NoRefDocument)
            Else
                Me.ProgressBar1.Visible = False
                Me.LblProgress.Visible = False
                Me.BtnStop.Visible = False
            End If
            Me.ColPrimaryTransitLocation.Visible = False
            Me.BtnFindCodeNo.BringToFront()
            Me.ChkIsReveivement.Visible = False
            Me.ChkIsReceived.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDestinationCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDestinationCode.GotFocus
        Try
            Me.txtDestinationCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDestinationCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDestinationCode.LostFocus
        Try
            Me.txtDestinationCode.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDestinationCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDestinationCode.KeyDown
        Dim i As New FrmCariMasterStorageArea
        Dim Obj As New clsMasterStorageArea

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.txtDestinationCode.Text) = True Then
                    Me.txtPrimaryDestination.Text = Obj.GetPrimary(Me.txtDestinationCode.Text)
                    Me.TxtRemarkItem.Focus()
                Else
                    If Me.txtDestinationCode.Text <> "" Then
                        i.txtCari.Text = Me.txtDestinationCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryDestination.Text = i.PrimaryStorageArea
                        Me.txtDestinationCode.Text = Obj.GetStorageAreaCode(Me.txtPrimaryDestination.Text)
                        Me.TxtRemarkItem.Focus()
                    Else
                        Me.txtDestinationCode.Clear()
                        Me.txtPrimaryDestination.Text = "-1"
                        Me.txtDestinationCode.Focus()
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

    Private Sub txtDestinationCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDestinationCode.Leave
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            If Me.txtDestinationCode.Text <> "" Then
                If obj.ApakahSudahAda(Me.txtDestinationCode.Text) = True Then
                    Me.TxtRemarkItem.Focus()
                    Me.txtPrimaryDestination.Text = obj.GetPrimary(Me.txtDestinationCode.Text)
                Else
                    If Me.txtDestinationCode.Text <> "" Then
                        i.txtCari.Text = Me.txtDestinationCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryDestination.Text = i.PrimaryStorageArea
                        Me.txtDestinationCode.Text = obj.GetStorageAreaCode(Me.txtPrimaryDestination.Text)
                        Me.TxtRemarkItem.Focus()
                    Else
                        Me.txtDestinationCode.Clear()
                        Me.txtPrimaryDestination.Text = "-1"
                        Me.txtDestinationCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryDestination.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindDestination_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindDestination.Click
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtDestinationCode.Text <> "" Then
                i.txtCari.Text = Me.txtDestinationCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.txtPrimaryDestination.Text = i.PrimaryStorageArea
                Me.txtDestinationCode.Text = obj.GetStorageAreaCode(Me.txtPrimaryDestination.Text)
                Me.TxtRemarkItem.Focus()
            Else
                Me.txtDestinationCode.Clear()
                Me.txtDestinationCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub dtpDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.Leave
        Dim DateDocRef As Date = Now

        Try
            If Me.ChkIsReveivement.Checked = True Then
                'transaksi receivement stock transfer dtp nya tidak bisa lebih kecil dari dtp reference documentnya
                DateDocRef = Me.objStockTransfer.GetTanggalTransaksi(Me.TxtPrimaryRefDoc.Text)
                If Me.dtpDate.Value.Date < DateDocRef Then
                    MsgBox("Receivement date can't be smaller from stock transfer date.", MsgBoxStyle.Information, "Information")
                    Me.dtpDate.Focus()
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
                    If Me.objStockTransfer.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objStockTransfer.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtStockInNo.Text = Me.objStockTransfer.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtStockInNo.Text = Me.objStockTransfer.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtStockInNo.Text = Me.objStockTransfer.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatch.Leave
        Try
            If Me.TxtRemark.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockTransfer_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            If ObjProduct.GetJml_LocationControlled = 0 And ObjLocation.GetJml() <= 1 Then
                MsgBox("Stock location transfer can only be done if you have at least one location controlled product and 2 storage area.", _
                MsgBoxStyle.Information, "Information")
                Me.Close()
                Exit Sub
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
                If Me.ChkIsReveivement.Checked = True Then
                    'jika stock transfer receivement maka tidak bisa print
                    MsgBox("You can't print this data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Dim i As New FrmLaporanFormatStockInOut_Bundling

                i.primary = Me.TxtPrimaryStockIn.Text
                i.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockTransfer
                i.ShowDialog()

                i = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.objStockTransfer.CariDetail(vPrimary, Me.dgvItem, Me.ChkIsReveivement.Checked, Me.TblItemCustom) = True Then
                Me.SalinKeField()
                Me.btnEdit.Focus()
            Else
                Me.btnFind.Focus()
            End If
            Me.btnFind.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub EditTanggal(ByVal vTglTujuan As Date)
        'Dim hasil As MsgBoxResult

        Try
            If Me.TxtStockInNo.Text <> "" Then
                If Me.objStockTransfer.ApakahBolehEdit(Me.UserCode, True) = True Then
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
                        '    If Me.objStockTransfer.SimpanBaru(Me.dgvItem, Me.UserCode) Then
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
                        If Me.objStockTransfer.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me.TblItemCustom) Then
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

    Public Sub FromCutOff_LocationTrans_Create(ByVal vPrimaryTransaksi As Double)
        Try
            Me.baru = False
            Me.objStockTransfer.ResetObject(Me.dgvItem)

            Me.objStockTransfer.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.ChkIsReveivement.Checked, Me.TblItemCustom)
            Me.UserCode = Me.objStockTransfer.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_LocationTrans_Save()
        Try
            Me.baru = True
            Me.objStockTransfer.SimpanBaru1(Me.dgvItem, Me.UserCode, Me.TblItemCustom, False, False, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function InputData_TransferPickList(ByVal sender As Object, ByVal e As System.EventArgs, ByRef vIsTransOK As Boolean, ByRef vTblTransfer As DseLaporanProd_PickList.tbldetailDataTable, ByVal vListTransfer As List(Of String), ByVal vListOfTransfer As DseLaporanProd_PickList.DataTable4DataTable, ByVal vIsFormat As Integer) As Boolean
        Dim Row As DseLaporanProd_PickList.tbldetailRow
        Dim TblDummy As New DseLaporanProd_PickList.tbldetailDataTable
        Dim RowDummy As DseLaporanProd_PickList.tbldetailRow
        Dim ListDone As New List(Of String)
        Dim tbl2 As New DseLaporanProd_PickList.DataTable2DataTable
        Dim Row2 As DseLaporanProd_PickList.DataTable2Row
        Dim tblDuplicate As New DseLaporanProd_PickList.tbldetailDataTable
        'Dim row2_ As dseMasterProduct.stock_location_viewRow
        'Dim row3_ As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item1Row
        'Dim rowListOfTransfer As DseLaporanProd_PickList.DataTable4Row

        Try
            tblDuplicate = vTblTransfer.Copy
            tbl2.Rows.Clear()
            ListDone.Clear()

            Me.CekApakahBenarExpireDate_AllItem(vTblTransfer, vListTransfer)

            TblDummy = vTblTransfer.Copy
            If vIsFormat = 1 Then
                For ij As Integer = vTblTransfer.Rows.Count - 1 To 0 Step -1
                    Row = vTblTransfer.Rows(ij)
                    If vListTransfer.Contains(Row.Identfier) = False Then
                        vTblTransfer.Rows.Remove(Row)
                    End If
                Next
            Else
                For ij As Integer = TblDummy.Rows.Count - 1 To 0 Step -1
                    RowDummy = TblDummy.Rows(ij)
                    vListOfTransfer.DefaultView.RowFilter = "[Primary RM] = " & RowDummy.Primary_Product.ToString & " and [Location To Code] ='" & RowDummy.To_Location & "'"
                    If vListOfTransfer.DefaultView.Count = 0 Then
                        TblDummy.Rows.Remove(RowDummy)
                    End If
                Next
            End If

            For ij As Integer = 0 To vTblTransfer.Rows.Count - 1
                Row = vTblTransfer.Rows(ij)
                tbl2.DefaultView.RowFilter = Nothing
                tbl2.DefaultView.RowFilter = "[Primary From]=" & Row.Primary_From_Location & " AND [Primary To]=" & Row.Primary_To_Location
                If tbl2.DefaultView.Count = 0 Then
                    TblDummy.DefaultView.RowFilter = Nothing
                    TblDummy.DefaultView.RowFilter = "[Primary From Location]=" & Row.Primary_From_Location & " AND [Primary To Location]=" _
                    & Row.Primary_To_Location & "AND [Qty] > 0 "
                    If TblDummy.DefaultView.Count > 0 Then
                        Me.NonAktifkanDetail()
                        Me.NewTransfer()
                        Me.TxtPrimaryStockIn.Text = "-1"

                        For i As Integer = 0 To TblDummy.DefaultView.Count - 1
                            RowDummy = TblDummy.DefaultView(i).Row
                            With RowDummy
                                If vListTransfer.Contains(RowDummy.Identfier) = True Then
                                    Me.TxtCodeNo.Text = RowDummy.Code_No
                                    Me.TxtCodeNo_Leave(sender, e)

                                    Me.cboBatchNo.SelectedIndex = -1
                                    Me.cboBatchNo.SelectedValue = RowDummy.Batch
                                    Me.cboLocation.SelectedIndex = -1
                                    Me.cboLocation.SelectedValue = RowDummy.From_Location

                                    Me.txtPrimaryLocation.Text = RowDummy.Primary_From_Location
                                    Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(RowDummy.Qty, RowDummy._Decimal)
                                    Me.txtDestinationCode.Text = RowDummy.To_Location
                                    Me.txtPrimaryDestination.Text = RowDummy.Primary_To_Location
                                    Me.Input()
                                End If
                            End With
                        Next

                        Row2 = tbl2.NewDataTable2Row
                        Row2.Primary_From = Row.Primary_From_Location
                        Row2.Primary_To = Row.Primary_To_Location
                        Row2.Identifier = Row.Identfier
                        tbl2.AddDataTable2Row(Row2)

                        Me.btnOK.Focus()
                        Me.ShowDialog()
                    End If

                    vIsTransOK = Me.IsTransOK
                End If
            Next

            vTblTransfer.Rows.Clear()
            vTblTransfer = tblDuplicate.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Dispose()
        End Try
    End Function

    Public Function Input() As Boolean
        Dim Lot As String = ""

        Dim PrimaryTransitLocation As Double = -1
        Dim TransitLocation As String = ""
        Dim TransitWarehouse As String = ""

        Try
            If Me.cboBatchNo.SelectedValue = Nothing Then
                Lot = ""
            Else
                Lot = Me.cboBatchNo.SelectedValue
            End If

            'Get dulu data transit locationnya
            Me.GetTransitData(PrimaryTransitLocation, TransitLocation, TransitWarehouse, Me.txtPrimaryLocation.Text)

            If Me.ApakahDataItemBenar1(PrimaryTransitLocation) = True Then
                'cek dulu apakah transit location?
                Dim Input_PrimaryDestination, Input_PrimaryTransit As Double
                Dim Input_Destination, Input_Transit, Input_DestinationWarehouse, Input_TransitWarehouse As String
                Dim QtyReceived As Decimal = 0
                Dim UnitReceived As String = Me.lblUnits.Text
                Dim DestinationWarehouse As String = Me.ObjLocation.GetWarehouseCode(Me.txtPrimaryDestination.Text)
                If PrimaryTransitLocation = -1 Then
                    Input_PrimaryDestination = Me.txtPrimaryDestination.Text
                    Input_Destination = Me.txtDestinationCode.Text
                    Input_DestinationWarehouse = DestinationWarehouse
                    Input_PrimaryTransit = -1
                    Input_Transit = ""
                    Input_TransitWarehouse = ""
                Else
                    Input_PrimaryDestination = PrimaryTransitLocation
                    Input_Destination = TransitLocation
                    Input_DestinationWarehouse = TransitWarehouse
                    Input_PrimaryTransit = Me.txtPrimaryDestination.Text
                    Input_Transit = Me.txtDestinationCode.Text
                    Input_TransitWarehouse = DestinationWarehouse
                End If

                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, _
                 Me.txtSecondaryQty.Text, Input_PrimaryDestination, Me.DtpExpired.Value, Input_PrimaryTransit)
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

                    Dim rd As DataGridViewRow
                    Me.dgvItem.Rows.Add()
                    rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                    rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                    rd.Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                    rd.Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                    rd.Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatch.Text
                    rd.Cells(Me.ColUnit.Index).Value = Me.lblUnits.Text
                    rd.Cells(Me.ColSecondaryQty.Index).Value = Me.txtSecondaryQty.Text
                    rd.Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                    rd.Cells(Me.ColBatch.Index).Value = Lot
                    rd.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                    rd.Cells(Me.ColLocation.Index).Value = Me.cboLocation.SelectedValue
                    rd.Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                    rd.Cells(Me.ColDestinationLoc.Index).Value = Input_Destination
                    rd.Cells(Me.ColWarehouseDestination.Index).Value = Input_DestinationWarehouse
                    rd.Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                    rd.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                    rd.Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                    rd.Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnits.Text
                    rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = Me.txtPrimarySecondaryUnits.Text
                    rd.Cells(Me.ColPrimaryDestination.Index).Value = Input_PrimaryDestination
                    rd.Cells(Me.ColTransitLocation.Index).Value = Input_Transit
                    rd.Cells(Me.ColTransitWarehouse.Index).Value = Input_TransitWarehouse
                    rd.Cells(Me.ColQtyReceived.Index).Value = QtyReceived
                    rd.Cells(Me.ColUnitReceived.Index).Value = UnitReceived
                    rd.Cells(Me.ColPrimaryTransitLocation.Index).Value = Input_PrimaryTransit

                    ''set ulang nomor
                    'Me.ObjCD.SetNumber(Me.dgvItem)
                End If
                'kosongkan gbo item
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
                'format dgv
                Me.FormatDGV()
                'fokuskan ke baris terakhir
                'While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                'Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                '    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                'End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub NewTransfer()
        Try
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.dtpDate.Focus()
            Me.objStockTransfer.ResetObject(Me.dgvItem)
            Me.TxtStockInNo.Text = Me.objStockTransfer.GetNoBaru(Me.dtpDate.Value)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim vListExpireDate_ As List(Of Double)
    Private Function CekApakahBenarExpireDate_AllItem(ByVal vTBL As DseLaporanProd_PickList.tbldetailDataTable, _
    ByVal vListTransfer As List(Of String)) As Boolean
        Dim Row As DseLaporanProd_PickList.tbldetailRow
        Dim tblList As New DseLaporanProd_PickList.DataTable3DataTable
        Dim RowList As DseLaporanProd_PickList.DataTable3Row
        Dim OK As Boolean = True
        Try
            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer

            For i As Integer = vTBL.Rows.Count - 1 To 0 Step -1
                Row = vTBL.Rows(i)
                If vListTransfer.Contains(Row.Identfier) = True Then
                    If Row.Expiry_Date_Controlled = True Then
                        If Row.Expiry_Date < Now Then
                            OK = False
                            tblList.DefaultView.RowFilter = "[Code No]='" & Row.Code_No & "' AND [Batch]='" & Row.Batch & "' AND [Expiry Date]='" & Row.Expiry_Date & "'"
                            If tblList.DefaultView.Count = 0 Then
                                RowList = tblList.NewDataTable3Row
                                RowList.Code_No = Row.Code_No
                                RowList.Batch = Row.Batch
                                RowList.Expiry_Date = Row.Expiry_Date
                                tblList.AddDataTable3Row(RowList)
                            End If
                        End If
                    End If
                End If
            Next

            Dim Dummy As String = ""
            Dim No As Integer = 1
            For Each RowList In tblList
                If No <> 21 Then
                    If Dummy = "" Then
                        Dummy = No & ". Code No :" & RowList.Code_No & ",  Batch No :" & RowList.Batch & ", Expiry Date :" & Format(RowList.Expiry_Date, "dd MMM yyy")
                    Else
                        Dummy &= vbNewLine & No & ". Code No :" & RowList.Code_No & ", Batch No :" & RowList.Batch & ", Expiry Date : " & Format(RowList.Expiry_Date, "dd MMM yyy")
                    End If
                    No += 1
                Else
                    Dummy &= " any more...."
                    Exit For
                End If
            Next

            If OK = False Then
                If MsgBox("This product expiry date is smaller than today. " & vbNewLine & _
                           Dummy & vbNewLine & "Do you want to continue?", MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                    For Each RowList In tblList.Rows
                        vTBL.DefaultView.RowFilter = ""
                        vTBL.DefaultView.RowFilter = "[Code No]='" & RowList.Code_No & "'"
                        If vTBL.DefaultView.Count > 0 Then
                            For i As Integer = vTBL.DefaultView.Count - 1 To 0 Step -1
                                Row = vTBL.DefaultView(i).Row
                                vTBL.Rows.Remove(Row)
                            Next
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataItemBenar1(ByVal PrimaryTransitLocInput As Double) As Boolean
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
            '1.b. hanya boleh masukkan product yang kontrol lokasi
            If Me.ObjProduct.ApakahLocationControlled(Me.TxtPrimaryCodeNo.Text) = False Then
                MsgBox("Please only enter goods with location controll.", MsgBoxStyle.Information, "Information")
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
            '5.b. Bila location controlled, kode lokasi tujuan tidak diisikan
            If Me.chkLocationControlled.Checked = True Then
                If Me.txtPrimaryDestination.Text = "-1" Then
                    MsgBox("Please enter transfer destination.", MsgBoxStyle.Information, _
                    "Information")
                    Me.txtDestinationCode.Focus()
                    Return False
                End If
            End If
            '5.c. Location tujuan dan asal tidak boleh sama
            If Me.txtPrimaryDestination.Text = Me.txtPrimaryLocation.Text Then
                MsgBox("Source dan destination location must be different.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            ''6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            ''   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            ''   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            'If Me.chkExpiryDateControlled.Checked = True Then
            '    If Me.DtpExpired.Value.Date < Me.dtpDate.Value.Date Then
            '        If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
            '        MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
            '            Me.DtpExpired.Focus()
            '            Return False
            '        End If
            '    End If
            'End If

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
                    If row.Cells(Me.ColLocation.Index).Value = Me.cboBatchNo.Text And Me.cboBatchNo.Text <> "" And row.Cells(Me.ColPrimaryLocation.Index).Value = _
                    Me.TxtPrimaryCodeNo.Text Then
                        If CDate(row.Cells(Me.ColWareHouse.Index).Value) <> Me.DtpExpired.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNo.Text + " on " + _
                            Format(CDate(row.Cells(Me.ColWareHouse.Index).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
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
            Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Double
            JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockTransferItem.GetJml(Me.TxtPrimaryStockIn.Text, Me.TxtPrimaryCodeNo.Text, _
            Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text, Me.DtpExpired.Value)
            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNo.Text, Me.txtSecondaryQty.Text, Me.txtPrimaryLocation.Text, _
            Me.cboBatchNo.Text, Me.chkBatchNoControlled.Checked, _
            Me.chkLocationControlled.Checked, Me.chkUseSecondaryUnit.Checked)
            JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatch.Text)
            If JmlSetelahTambah < 0 Then
                If Me.chkUseSecondaryUnit.Checked = True Then
                    MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGV) + " " + _
                    Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) + " @ " + Me.txtSecondaryQty.Text + " " + _
                    Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text) + ".", MsgBoxStyle.Information, _
                    "Information")
                Else
                    MsgBox("Maximum value you can enter is " + CStr(JmlGudang + JmlObjLama - JmlDGV) + " " + _
                    CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)) + ".", MsgBoxStyle.Information, _
                    "Information")
                End If
                Me.TxtQtyperBatch.Focus()
                Return False
            End If
            '9.b. Cek stock pada tanggal yang bersangkutan
            'msgbox("belum")

            '10. Cek batasan transit location
            If Me.CekBatasanInputTransit(PrimaryTransitLocInput) = False Then
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

    Private Function CekBatasanInputTransit(ByVal PrimaryTransitLocInput As Double) As Boolean
        Try
            If Me.dgvItem.RowCount > 0 Then
                If Me.chkUseTransit.Checked = True Then
                    'jika transit location maka primary destination dengan primary transit location input harus sama
                    If PrimaryTransitLocInput = -1 Then
                        MsgBox("You must input a location that has a transit", MsgBoxStyle.Information, "Information")
                        Return False
                    Else
                        If Me.dgvItem.Item(Me.ColPrimaryDestination.Index, 0).Value <> PrimaryTransitLocInput Then
                            MsgBox("You must input the same transit location (" & Me.dgvItem.Item(Me.ColDestinationLoc.Index, 0).Value & " <> " & Me.ObjLocation.GetStorageAreaCode(PrimaryTransitLocInput) & ")", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                Else
                    If PrimaryTransitLocInput <> -1 Then
                        MsgBox("Can't input locations that have transit", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
        'Dim ij As New FrmTransaksiStockAdvanceAdd
        Dim Obj As New ClsTransaksiBarcode
        Dim ij As New FrmTransaksiStockAdvanceAdd
        Dim TBLData As New Core_Module.DseBarcodeInput.DataTable1DataTable
        Dim InputType As CustomSMESS.ClsCustomBarcode.EnumInputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add
        Dim Index As Integer = -1
        Dim rowData As Core_Module.DseBarcodeInput.DataTable1Row

        Dim PrimaryTransitLocation As Double = -1
        Dim TransitLocation As String = ""
        Dim TransitWarehouse As String = ""

        Dim Input_PrimaryDestination, Input_PrimaryTransit As Double
        Dim Input_Destination, Input_Transit, Input_DestinationWarehouse, Input_TransitWarehouse As String
        Dim QtyReceived As Decimal = 0

        Try
            'ij.BackColorAll = Me.BackColorAll
            'ij.UserCode = Me.UserCode
            'ij.ParentForm1 = FrmTransaksiStockAdvanceAdd.EnumRefDoc.Location_Transfer_Source
            'ij.ShowDialog()
            Obj.GetDataInput(TBLData, InputType, True, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Location_Transfer_Source, Me.dgvItem, Me.TxtPrimaryStockIn.Text, Me.UserCode)
            'Obj.GetDataInput(TBLData, InputType, True, FrmTransaksiStockAdvanceAdd.EnumRefDoc.Location_Transfer_Source)

            If InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add Then

            ElseIf InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Replace Then
                If TBLData.Rows.Count > 0 Then
                    Me.dgvItem.Rows.Clear()
                End If
            End If

            If TBLData.Rows.count > 0 Then
                rowData = TBLData.rows(0)

                Me.GetTransitData(PrimaryTransitLocation, TransitLocation, TransitWarehouse, rowData.Primary_Location)

                If Me.CekBatasanInputTransit(PrimaryTransitLocation) = False Then
                    Exit Sub
                End If

                If PrimaryTransitLocation = -1 Then
                    Input_PrimaryDestination = rowData.Primary_Location_Target
                    Input_Destination = rowData.Locatian_Target 'Me.txtDestinationCode.Text
                    Input_DestinationWarehouse = rowData.Warehouse  'DestinationWarehouse
                    Input_PrimaryTransit = -1
                    Input_Transit = ""
                    Input_TransitWarehouse = ""
                Else
                    Input_PrimaryDestination = PrimaryTransitLocation
                    Input_Destination = TransitLocation
                    Input_DestinationWarehouse = TransitWarehouse
                    Input_PrimaryTransit = rowData.Primary_Location_Target 'Me.txtPrimaryDestination.Text
                    Input_Transit = rowData.Locatian_Target ' Me.txtDestinationCode.Text
                    Input_TransitWarehouse = rowData.Warehouse_Target 'DestinationWarehouse
                End If
            End If

            For Each rowData In TBLData
                With rowData
                    Index = Me.GetIndex2(.Primary_Product, .Primary_Location, .Batch, -1, .Primary_Location_Target, -1)
                    'cari data yang sama persis
                    If Index <> -1 Then
                        'tambahkan nilai qty
                        Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                        row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + CDec(.Qty), ._Decimal)
                    Else
                        'tambahkan baris baru
                        Dim ExpiredDateStr As String
                        If .Expired_Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                        End If

                        Dim rd As DataGridViewRow
                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColCodeNo.Index).Value = .Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rd.Cells(Me.ColUnit.Index).Value = .Unit
                        rd.Cells(Me.ColSecondaryQty.Index).Value = -1
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = "-"
                        rd.Cells(Me.ColBatch.Index).Value = .Batch
                        rd.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        rd.Cells(Me.ColLocation.Index).Value = .Location
                        rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse
                        rd.Cells(Me.ColDestinationLoc.Index).Value = Input_Destination
                        rd.Cells(Me.ColWarehouseDestination.Index).Value = Input_DestinationWarehouse
                        rd.Cells(Me.ColRemark.Index).Value = .Remarks
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Location
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Unit
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                        rd.Cells(Me.ColPrimaryDestination.Index).Value = Input_PrimaryDestination
                        rd.Cells(Me.ColTransitLocation.Index).Value = Input_Transit
                        rd.Cells(Me.ColTransitWarehouse.Index).Value = Input_TransitWarehouse
                        rd.Cells(Me.ColQtyReceived.Index).Value = QtyReceived
                        rd.Cells(Me.ColUnitReceived.Index).Value = .Unit
                        rd.Cells(Me.ColPrimaryTransitLocation.Index).Value = Input_PrimaryTransit
                    End If
                End With
            Next
            Me.ObjCD.SetNumber(Me.dgvItem)
            'format dgv
            Me.FormatDGV()

            Me.SetTransitStatus(Me.ChkIsReveivement.Checked)

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
        Dim rowData As Core_Module.DseBarcodeInput.DataTable1Row
        Dim Index As Integer = -1
        Dim AdaService As Boolean = False

        Dim PrimaryTransitLocation As Double = -1
        Dim TransitLocation As String = ""
        Dim TransitWarehouse As String = ""

        Dim Input_PrimaryDestination, Input_PrimaryTransit As Double
        Dim Input_Destination, Input_Transit, Input_DestinationWarehouse, Input_TransitWarehouse As String
        Dim QtyReceived As Decimal = 0

        Try
            Obj.GetDataInput(TBLData, InputType, True, CustomSMESS.ClsCustomBarcode.EnumTransaksiAcuan.LocationTransfer)

            If InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Add Then

            ElseIf InputType = CustomSMESS.ClsCustomBarcode.EnumInputType.Replace Then
                If TBLData.Rows.Count > 0 Then
                    Me.dgvItem.Rows.Clear()
                End If
            End If

            If TBLData.Rows.Count > 0 Then
                rowData = TBLData.Rows(0)

                Me.GetTransitData(PrimaryTransitLocation, TransitLocation, TransitWarehouse, rowData.Primary_Location)

                If Me.CekBatasanInputTransit(PrimaryTransitLocation) = False Then
                    Exit Sub
                End If

                If PrimaryTransitLocation = -1 Then
                    Input_PrimaryDestination = rowData.Primary_Location_Target
                    Input_Destination = rowData.Locatian_Target 'Me.txtDestinationCode.Text
                    Input_DestinationWarehouse = rowData.Warehouse 'DestinationWarehouse
                    Input_PrimaryTransit = -1
                    Input_Transit = ""
                    Input_TransitWarehouse = ""
                Else
                    Input_PrimaryDestination = PrimaryTransitLocation
                    Input_Destination = TransitLocation
                    Input_DestinationWarehouse = TransitWarehouse
                    Input_PrimaryTransit = rowData.Primary_Location_Target 'Me.txtPrimaryDestination.Text
                    Input_Transit = rowData.Locatian_Target ' Me.txtDestinationCode.Text
                    Input_TransitWarehouse = rowData.Warehouse_Target  'DestinationWarehouse
                End If
            End If

            For Each rowData In TBLData
                With rowData
                    If rowData.Is_Service = 1 Then
                        AdaService = True
                    Else
                        Index = Me.GetIndex2(.Primary_Product, .Primary_Location, .Batch, -1, .Primary_Location_Target, -1)
                        'cari data yang sama persis
                        If Index <> -1 Then
                            'tambahkan nilai qty
                            Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                            row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + CDec(.Qty), ._Decimal)
                        Else
                            'tambahkan baris baru
                            Dim ExpiredDateStr As String
                            If .Expired_Date = #1/1/1950# Then
                                ExpiredDateStr = "-"
                            Else
                                ExpiredDateStr = Format(.Expired_Date, "dd MMM yyyy")
                            End If

                            Dim rd As DataGridViewRow
                            Me.dgvItem.Rows.Add()
                            rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                            rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                            rd.Cells(Me.ColCodeNo.Index).Value = .Code
                            rd.Cells(Me.ColDescription.Index).Value = .Description
                            rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                            rd.Cells(Me.ColUnit.Index).Value = .Unit
                            rd.Cells(Me.ColSecondaryQty.Index).Value = -1
                            rd.Cells(Me.ColSecondaryUnit.Index).Value = "-"
                            rd.Cells(Me.ColBatch.Index).Value = .Batch
                            rd.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                            rd.Cells(Me.ColLocation.Index).Value = .Location
                            rd.Cells(Me.ColWareHouse.Index).Value = .Warehouse
                            rd.Cells(Me.ColDestinationLoc.Index).Value = Input_Destination
                            rd.Cells(Me.ColWarehouseDestination.Index).Value = Input_DestinationWarehouse
                            rd.Cells(Me.ColRemark.Index).Value = .Remarks
                            rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                            rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Location
                            rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Unit
                            rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                            rd.Cells(Me.ColPrimaryDestination.Index).Value = Input_PrimaryDestination
                            rd.Cells(Me.ColTransitLocation.Index).Value = Input_Transit
                            rd.Cells(Me.ColTransitWarehouse.Index).Value = Input_TransitWarehouse
                            rd.Cells(Me.ColQtyReceived.Index).Value = 0 'QtyReceived
                            rd.Cells(Me.ColUnitReceived.Index).Value = .Unit
                            rd.Cells(Me.ColPrimaryTransitLocation.Index).Value = Input_PrimaryTransit
                        End If
                    End If
                End With
            Next
            Me.ObjCD.SetNumber(Me.dgvItem)
            'format dgv
            Me.FormatDGV()

            Me.SetTransitStatus(Me.ChkIsReveivement.Checked)

            'fokuskan ke baris terakhir
            If Me.dgvItem.Rows.Count > 0 Then
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                         Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
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

    Private Function GetIndex2(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
   ByVal SecondaryQty As Double, ByVal vPrimaryDestination As Double, ByVal vPrimaryTransit As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If IsNumeric(row.Cells(Me.ColSecondaryQty.Index).Value) = False Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                    row.Cells(Me.ColBatch.Index).Value = BatchNo And vPrimaryDestination = row.Cells(Me.ColPrimaryDestination.Index).Value And _
                    row.Cells(Me.ColPrimaryTransitLocation.Index).Value = vPrimaryTransit Then
                        Return row.Index
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation And _
                    vPrimaryDestination = row.Cells(Me.ColPrimaryDestination.Index).Value And row.Cells(Me.ColBatch.Index).Value = BatchNo And _
                    row.Cells(Me.ColPrimaryTransitLocation.Index).Value = vPrimaryTransit Then
                        Return row.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CreateTransaksiReceived(ByVal vPrimaryRefDocument As Double, ByVal vNoRefDocument As String) As Boolean
        Try
            Me.btnNew.PerformClick()

            Me.objStockTransfer.PrimaryRefDoc = vPrimaryRefDocument
            Me.objStockTransfer.UseTransit = 0
            Me.objStockTransfer.TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit

            If Not vPrimaryRefDocument = -1 Then
                If Me.objStockTransfer.CariDetail(vPrimaryRefDocument, Me.dgvItem, Me.ApakahTransaksiReceived, Me.TblItemCustom, True) = True Then
                    With Me.objStockTransfer
                        Me.TxtRemark.Text = Me.objStockTransfer.Remarks
                        Me.objStockTransfer.Primary = -1
                        Me.objStockTransfer.ResetObject()
                        Me.dtpDate.Value = Now
                        Me.TxtPrimaryStockIn.Text = -1
                        Me.TxtPrimaryRefDoc.Text = vPrimaryRefDocument
                        Me.TxtStockInNo.Text = Me.objStockTransfer.GetNoBaru(Me.dtpDate.Value)

                        Me.chkUseTransit.Checked = False
                        Me.ChkIsReceived.Checked = False
                        Me.ChkIsReveivement.Checked = True
                        Me.TxtTransitStatus.Text = "Reference stock transfer no : " & vNoRefDocument
                    End With

                    Me.FormatDGV()
                    Me.SetTransit_InDgv(Me.ApakahTransaksiReceived)

                    Me.NonAktifkanDetail_Item()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function DeleteTransaksiReceived(ByVal vPrimaryReceiveDocument As Double, ByVal vPrimaryRefDoc As Double, ByVal vUserCode As String) As Boolean
        Dim DateDefault As Date = #1/1/1950#

        Try
            If Not vPrimaryReceiveDocument = -1 Then
                With Me.objStockTransfer
                    .Primary = vPrimaryReceiveDocument
                    .IsReceivement = 0
                    .TransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.No_Transit
                    .CreatedBy = vUserCode
                End With

                If Me.objStockTransfer.Hapus(vPrimaryReceiveDocument, vUserCode, False) = True Then
                    Me.objStockTransfer.UpdateStatusReceived("", DateDefault, DateDefault, vPrimaryRefDoc, clsTransaksiStockTransfer.EnumTransitStatus.On_Transit)
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub NonAktifkanDetail_Item()
        Try
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.TxtQtyperBatch.Enabled = False
            Me.llbAdvance.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnStop.Click
        Try
            If MsgBox("Stop creating transaction receive?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Abort
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        Dim i As New CustomSMESS.FrmTransaksiStockTransfer__AddByCustomProduct
        Try
            i.baru = Me.baru
            i.BackColorAll = Me.BackColorAll
            i.TanggalTransaksi = Me.dtpDate.Value.Date
            i.PrimaryStockIn = Me.TxtPrimaryStockIn.Text

            If Me.TxtRemark.Enabled = True Then
                i.edit = True
            Else
                i.edit = False
            End If
            If Me.baru = False Then
                Me.TblDBItemCustom = Me.TblItemCustom
                i.TBLDBCustomItem = Me.TblDBItemCustom.Copy
            Else
                i.TBLDBCustomItem = Me.TblItemCustom.Copy
            End If
            i.TBLCustomItem = Me.TblItemCustom.Copy

            If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.TblItemCustom = i.TBLCustomItem.Copy

                Me.IsiDgvItemFromTblCustom()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    'TR//
    Public TblItemCustom As New CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable
    Public TblDBItemCustom As New CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemDataTable
    Sub Atur1()
        Try
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockPindah") = True Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockTransfer__AddByCustomProduct") = False Then

                    'Me.gboItem.Enabled = True
                    Me.SplitContainer3.Enabled = True
                    Me.AddItemByCustom.Visible = False
                    Me.dgvItem.ReadOnly = False
                    Me.dgvItem.AllowUserToDeleteRows = True

                Else
                    Me.dgvItem.ReadOnly = True
                    'Me.gboItem.Enabled = False
                    Me.SplitContainer3.Enabled = False
                    Me.AddItemByCustom.Visible = True
                    Me.dgvItem.AllowUserToDeleteRows = False
                    Me.llbAdvance.Enabled = True
                End If
            Else
                'Me.gboItem.Enabled = True
                Me.SplitContainer3.Enabled = True
                Me.AddItemByCustom.Visible = False
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Function IsiDgvItemFromTblCustom() As Boolean
        Dim Input_PrimaryDestination, Input_PrimaryTransit As Double
        Dim Input_Destination, Input_Transit, Input_DestinationWarehouse, Input_TransitWarehouse As String
        Dim QtyReceived As Decimal = 0
        Dim UnitReceived As String = Me.lblUnits.Text
        Dim DestinationWarehouse As String = Me.ObjLocation.GetWarehouseCode(Me.txtPrimaryDestination.Text)


        Dim PrimaryTransitLocation As Double = -1
        Dim TransitLocation As String = ""
        Dim TransitWarehouse As String = ""

        Dim RowCustom As CustomSMESS.dseTransaksiStockTransfer.tbltrans_custom_stocktransfer_itemRow
        Dim RowDgv As DataGridViewRow
        Dim rd As DataGridViewRow
        Try

            'Terkait Transit
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Input_PrimaryDestination = Me.txtPrimaryDestination.Text
            Input_Destination = Me.txtDestinationCode.Text
            Input_DestinationWarehouse = DestinationWarehouse
            Input_PrimaryTransit = -1
            Input_Transit = ""
            Input_TransitWarehouse = ""
            'kosongkan dgv
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AllowUserToAddRows = False

            '/////////////Terkait Transit////////////////////
            For Each RowCustom In Me.TblItemCustom.Rows
                With RowCustom
                    Dim Index As Integer = Me.GetIndex(.Primary_Product, .Primary_Storage_Area, .Batch_No, _
                                -1, .primary_destination_area, .Expiry_Date, -1)
                    'cari data yang sama persis
                    If Index <> -1 Then
                        'tambahkan nilai qty
                        Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                        row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQty.Index).Value) + RowCustom.Qty, Me.ObjProduct.GetAngkaDecimal(RowCustom.Primary_Product))
                        If String.IsNullOrEmpty(row.Cells(Me.ColRemark.Index).Value) = False Then
                            row.Cells(Me.ColRemark.Index).Value = row.Cells(Me.ColRemark.Index).Value & ", " & .Remarks
                        Else
                            row.Cells(Me.ColRemark.Index).Value = .Remarks
                        End If
                    Else
                        Dim ExpiredDateStr As String

                        'tambahkan baris baru

                        If Me.DtpExpired.Value.Date = #1/1/1950# Then
                            ExpiredDateStr = "-"
                        Else
                            ExpiredDateStr = Format(.Expiry_Date, "dd MMM yyyy")
                        End If


                        Me.dgvItem.Rows.Add()
                        rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                        rd.Cells(Me.ColNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, Me.ObjProduct.GetAngkaDecimal(.Primary_Product))
                        rd.Cells(Me.ColUnit.Index).Value = .unit
                        rd.Cells(Me.ColSecondaryQty.Index).Value = -1
                        rd.Cells(Me.ColSecondaryUnit.Index).Value = ""
                        rd.Cells(Me.ColBatch.Index).Value = .Batch_No
                        rd.Cells(Me.ColExpiredDate.Index).Value = ExpiredDateStr
                        rd.Cells(Me.ColLocation.Index).Value = .location_from
                        rd.Cells(Me.ColWareHouse.Index).Value = .warehouse_from
                        rd.Cells(Me.ColDestinationLoc.Index).Value = .location_to
                        rd.Cells(Me.ColWarehouseDestination.Index).Value = .warehouse_to
                        rd.Cells(Me.ColRemark.Index).Value = .Remarks
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Storage_Area
                        rd.Cells(Me.ColPrimaryUnits.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecondaryUnits.Index).Value = -1
                        rd.Cells(Me.ColPrimaryDestination.Index).Value = .primary_destination_area
                        rd.Cells(Me.ColTransitLocation.Index).Value = ""
                        rd.Cells(Me.ColTransitWarehouse.Index).Value = ""
                        rd.Cells(Me.ColQtyReceived.Index).Value = 0
                        rd.Cells(Me.ColUnitReceived.Index).Value = ""
                        rd.Cells(Me.ColPrimaryTransitLocation.Index).Value = Input_PrimaryTransit


                        ''set ulang nomor
                        'Me.ObjCD.SetNumber(Me.dgvItem)
                    End If
                End With
            Next



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

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class