Option Explicit On
Option Strict Off

Imports Inventory_Module
Imports Core_Module
Imports CustomSMESS

Public Class FrmTransaksiDO3_PickQty
    Dim Buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    Dim vTBLQty As New dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable
    Dim vTBLQty_TdkBoleh As New dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable
    Dim LangsungCLose As Boolean = False

    Public PrimaryDO As Double = -1
    Public PrimarySO As Double = -1
    Public PrimaryProduct As Double = -1
    Public PrimaryUnitDefault As Double = -1
    Public PrimaryUnitMulti As Double = -1
    Public DefPrecDefault As Integer = 0
    Public DefPrecMulti As Integer = 0
    Public FrmRef As FrmTransaksiDO3
    Public vDGVDOItem As DataGridView
    Public UserCode As String = "" 
    Public ExpiryControl As Boolean = False
    Public BatchNoCOntrol As Boolean = False

    Public UnitDefault As String = ""
    Public UnitMulti As String = ""
    Public SONumber As String = ""
    Public CodeNo As String = ""
    Public Desc As String = ""
    Public RatioDesc As String = ""

    Public RatioMulti As Decimal = 1
    Public RatioDefault As Decimal = 1
    Public QtyOBMulti As Decimal = 0
    Public QtyOBDefault As Decimal = 0
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public QtyAvailableDefault_Ref As Decimal = 0
    Public QtyAvailableMulti_Ref As Decimal = 0

    'aul
    Dim ApakahMemilikiAksesAddByBarcode As Boolean = False
    Dim ObjLoc As New Core_Module.clsMasterStorageArea
    Dim ObjWH As New Core_Module.clsMasterWarehouse
    Dim ObjBarcode As New CustomSMESS.ClsTransaksiBarcode
    Dim ObjBarcodeScanItem As New CustomSMESS.ClsTransaksiBarcode_ScanItem
    Dim InputBarcode As EnumInput = EnumInput.ManualInput
    Dim MaxLabel As Integer = 99
    Public BarcodeType As ClsTransaksiBarcode_Scan.EnumBarcodeType = ClsTransaksiBarcode_Scan.EnumBarcodeType.BarcodeSMESS
    Public BatchOption As FrmBarcode_Scan_Batch.EnumBatch = FrmBarcode_Scan_Batch.EnumBatch.Batch_Select

    Enum EnumMethod As Integer
        LocFind
        LocKeyDown
        LocLeave
    End Enum
    Enum EnumCounter As Integer
        Tambah
        Kurang
    End Enum
    Enum EnumInput As Integer
        ScanInput
        ManualInput
    End Enum

    Private Sub FrmTransaksiDO3_PickQty_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungCLose = False Then

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub FrmTransaksiDO3_PickQty_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItemDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'fill tampilan header
            Me.TxtCodeNo.Text = Me.CodeNo
            Me.TxtDesc.Text = Me.Desc
            Me.TxtSONumber.Text = Me.SONumber
            If Me.PrimaryUnitDefault = Me.PrimaryUnitMulti Then
                Me.SplitContainer2.Panel2Collapsed = True
                Me.SplitContainer3.Panel2Collapsed = True
                Me.SplitContainer4.Panel2Collapsed = True
                Me.LblRatio.Visible = False
            End If
            Me.TxtQtyDeliveryDefault.Text = Me.ObjCD.DigitGrouping(0, Me.DefPrecDefault)
            Me.TxtQtyDeliveryMulti.Text = Me.ObjCD.DigitGrouping(0, Me.DefPrecMulti)
            Me.TxtQtyOrderBalanceDefault.Text = Me.ObjCD.DigitGrouping(Me.QtyOBDefault, Me.DefPrecDefault)
            Me.TxtQtyOrderBalanceMulti.Text = Me.ObjCD.DigitGrouping(Me.QtyOBMulti, Me.DefPrecMulti)
            Me.LblUnitAvailableDefault.Text = Me.UnitDefault
            Me.LblUnitAvailableMulti.Text = Me.UnitMulti
            Me.LblUnitDeliveryDefault.Text = Me.UnitDefault
            Me.LblUnitDeliveryMulti.Text = Me.UnitMulti
            Me.LblUnitOrderBalanceMulti.Text = Me.UnitMulti
            Me.LblUnitOrderBalanceDefault.Text = Me.UnitDefault
            Me.LblRatio.Text = Me.RatioDesc 

            'fill data item
            Me.FillItem()
            If Me.dgvItemDetail.RowCount = 0 Then
                MsgBox("No available stock available.", MsgBoxStyle.Information, "Information")
                Me.LangsungCLose = True
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If

            'aul
            Me.TxtBarcodeLabel.MaxLength = 2
            Me.CancelButton = Me.btnCancel
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.lblCounter.Text = -1
            Me.HitungCounter(EnumCounter.Tambah)

            Me.SettingColumn_DgvBarcode()

            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiDO_AddByBarcode") = True Then
                Me.ApakahMemilikiAksesAddByBarcode = True
            End If

            If Me.ApakahMemilikiAksesAddByBarcode = False Then
                Me.GroupBox1.Visible = False
                Me.llbSelectAll.Top = Me.llbSelectAll.Top - (Me.llbLocationAccess.Top - Me.GroupBox1.Top)
                Me.llbUnselectAll.Top = Me.llbSelectAll.Top
                Me.TabControl1.Height = Me.TabControl1.Height + (Me.llbLocationAccess.Top - Me.GroupBox1.Top)
                Me.TabControl1.Top = Me.TabControl1.Top - (Me.llbLocationAccess.Top - Me.GroupBox1.Top)
                Me.llbLocationAccess.Top = Me.GroupBox1.Top
            End If

            Me.SplitContainer2.TabStop = True
            Me.SplitContainer3.TabStop = True

            '2019-09-18: iwan
            'dihapus info ini karena selain informasi, user tidak bisa melakukan apa pun
            'lebih jauh lagi, klo untuk barang yg batch no control, ini bermasalah
            'If CDec(Me.TxtQtyAvailableDefault.Text) <> CDec(QtyAvailableDefault_Ref) Then
            '    Me.TxtQtyAvailableDefault.Text = Me.ObjCD.DigitGrouping(Me.QtyAvailableDefault_Ref, Me.DefPrecDefault)
            '    Me.TxtQtyAvailableMulti.Text = Me.ObjCD.DigitGrouping(Me.QtyAvailableMulti_Ref, Me.DefPrecMulti)
            '    MsgBox("Warning: inconsistent available qty due to inconsistent do data input. Please check location data for each pending do.", MsgBoxStyle.Information, "Information")
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FillItem()
        Dim vRowQty As dseTransaksiStockInOut_Sales.tblmaster_productqtyRow
        Dim rowDGV As DataGridViewRow
        Dim JmlDefault As Decimal = 0
        Dim JmlMulti As Decimal = 0
        Dim DTADOPending As New DseTransaksiDO_AddByLocationTableAdapters.tbltrans_do_itemTableAdapter
        Dim TBLDOPending As New DseTransaksiDO_AddByLocation.tbltrans_do_itemDataTable
        Dim rowDOPending As DseTransaksiDO_AddByLocation.tbltrans_do_itemRow
        Dim TBLDODB As New DseTransaksiDO_AddByLocation.tbltrans_do_itemDataTable
        Dim rowDODB As DseTransaksiDO_AddByLocation.tbltrans_do_itemRow

        Try
            Me.vTBLQty.Rows.Clear()
            Me.vTBLQty_TdkBoleh.Rows.Clear()

            DTADOPending.Connection.ConnectionString = Core_Module.GetConstring

            'load dari item product qty
            If Me.LoadQty(Me.vTBLQty, Me.PrimaryProduct) = True Then
                Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                'tambahkan dengan item db di db atas primary product dan primary do tsb
                DTADOPending.FillBy_PrimaryDO_PrimaryProduct(TBLDODB, Me.PrimaryDO, Me.PrimaryProduct)
                For Each rowDODB In TBLDODB.Rows
                    If rowDODB.primary_location <> -1 Then
                        'tambahkan ke baris product qty
                        vTBLQty.DefaultView.RowFilter = "[Primary Product] = " + rowDODB.Primary_Product.ToString + _
                        " AND [Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowDODB.Batch_No) + "'" + _
                        " AND [Primary Storage Area] = " + rowDODB.primary_location.ToString

                        If vTBLQty.DefaultView.Count > 0 Then
                            vRowQty = vTBLQty.DefaultView(0).Row
                            vRowQty.Qty += rowDODB.Total_Qty
                        Else
                            vRowQty = vTBLQty.Newtblmaster_productqtyRow
                            vRowQty.Expiry_Date = rowDODB.Expiry_Date
                            vRowQty.Lot_No = rowDODB.Batch_No
                            vRowQty.Primary = -1
                            vRowQty.Primary_Product = rowDODB.Primary_Product
                            vRowQty.Primary_Storage_Area = rowDODB.primary_location
                            vRowQty.Primary_Warehouse = rowDODB.Primary_Warehouse
                            vRowQty.Qty = rowDODB.Total_Qty
                            vRowQty.Qty_Secondary = -1
                            vRowQty.Storage_Area_Code = rowDODB.Storage_Area_Code
                            vRowQty.Urutan = 0
                            vRowQty.Warehouse_Code = rowDODB.Warehouse_Code

                            vTBLQty.Rows.Add(vRowQty)
                        End If
                    End If
                Next

                'load item do pending
                DTADOPending.Fill(TBLDOPending)
                'kurangkan item on hand dengan qty pending do dari do lain yang ada primary lokasinya
                For ij As Double = TBLDOPending.Rows.Count - 1 To 0 Step -1
                    rowDOPending = TBLDOPending.Rows(ij)
                    If rowDOPending.primary_location <> -1 Then
                        vTBLQty.DefaultView.RowFilter = "[Primary Product] = " + rowDOPending.Primary_Product.ToString + _
                        " AND [Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowDOPending.Batch_No) + "'" + _
                        " AND [Primary Storage Area] = " + rowDOPending.primary_location.ToString

                        If vTBLQty.DefaultView.Count > 0 Then
                            vRowQty = vTBLQty.DefaultView(0).Row
                            vRowQty.Qty -= rowDOPending.Total_Qty
                            If vRowQty.Qty <= 0 Then
                                vTBLQty.Rows.Remove(vRowQty)
                            End If
                        End If
                    End If
                Next

                'kurangkan tabel 1, dengan qty dari dgv atas primary product yg sama
                For Each rowDGV In Me.vDGVDOItem.Rows
                    If rowDGV.Cells(Me.FrmRef.ColPrimaryCodeNo.Index).Value = Me.PrimaryProduct Then
                        vTBLQty.DefaultView.RowFilter = "[Primary Product] = " + Me.PrimaryProduct.ToString + _
                        " AND [Lot No] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(rowDGV.Cells(Me.FrmRef.ColBatchNo.Index).Value) + "'" + _
                        " AND [Primary Storage Area] = " + rowDGV.Cells(Me.FrmRef.ColPrimaryLocation.Index).Value.ToString

                        If vTBLQty.DefaultView.Count > 0 Then
                            vRowQty = vTBLQty.DefaultView(0).Row
                            vRowQty.Qty -= CDec(rowDGV.Cells(Me.FrmRef.ColTotalQty.Index).Value)
                            If vRowQty.Qty <= 0 Then
                                vTBLQty.Rows.Remove(vRowQty)
                            End If
                        End If
                    End If
                Next

                'split menjadi 2 table, tabel kedua berisikan lokasi yang tidak boleh
                Me.Split2Tabel(Me.vTBLQty, Me.vTBLQty_TdkBoleh)

                'kurangkan table 2, dengan qty pending dari do lain

                'kurangkan tabel 1, dengan qty pending dari do lain


                'tampilkan ke dgv
                'sort berdasarkan warehouse, location, expiry, batch
                Me.vTBLQty.DefaultView.RowFilter = ""
                Me.vTBLQty.DefaultView.Sort = "Urutan, [Warehouse Code], [Storage Area Code], [Expiry Date], [Lot No]"
                For Each DV As DataRowView In Me.vTBLQty.DefaultView
                    vRowQty = DV.Row

                    Me.dgvItemDetail.Rows.Add()
                    rowDGV = Me.dgvItemDetail.Rows(Me.dgvItemDetail.Rows.Count - 1)
                    rowDGV.Cells(Me.ColBatchNo.Index).Value = vRowQty.Lot_No
                    rowDGV.Cells(Me.ColExpiry.Index).Value = vRowQty.Expiry_Date
                    rowDGV.Cells(Me.ColLocation.Index).Value = vRowQty.Storage_Area_Code
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItemDetail.Rows.Count
                    rowDGV.Cells(Me.ColQtyAvailable.Index).Value = vRowQty.Qty
                    rowDGV.Cells(Me.ColQtyDelivery.Index).Value = 0
                    rowDGV.Cells(Me.ColUnit.Index).Value = Me.UnitDefault
                    rowDGV.Cells(Me.ColWarehouse.Index).Value = vRowQty.Warehouse_Code
                    rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = vRowQty.Primary_Storage_Area

                    JmlDefault += vRowQty.Qty
                Next

                Me.ColExpiry.DefaultCellStyle.Format = "dd MMM yyyy"
                Me.ColQtyDelivery.DefaultCellStyle.Format = "N" + Me.DefPrecDefault.ToString
                Me.ColQtyAvailable.DefaultCellStyle.Format = "N" + Me.DefPrecDefault.ToString
                If Me.ExpiryControl = True Then
                    Me.ColExpiry.Visible = True
                Else
                    Me.ColExpiry.Visible = False
                End If
                If Me.BatchNoCOntrol = True Then
                    Me.ColBatchNo.Visible = True
                Else
                    Me.ColBatchNo.Visible = False
                End If
                Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

                Me.TxtQtyAvailableDefault.Text = Me.ObjCD.DigitGrouping(JmlDefault, Me.DefPrecDefault)
                Me.TxtQtyAvailableMulti.Text = ObjCD.DigitGrouping((JmlDefault * Me.RatioMulti) / Me.RatioDefault, Me.DefPrecMulti, Core_Module.ClsCekData.EnumTipeRounding.Up)
            Else
                Me.LangsungCLose = True
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Split2Tabel(ByVal vTBLBoleh As dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable, ByVal vTBLTidakBoleh As dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable)
        Dim Split As Boolean = True
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim DefaultBoleh1 As Boolean = False
        Dim Dummy As String = ""
        Dim DTAUserLocationAccess As New Core_Module.DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilegeTableAdapter
        Dim TBLUserLocationAccess As New Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeDataTable
        Dim rowUserLocationAccess As Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilegeRow

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Location_Privileges, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_User_Not_IncludeInList)
            'cek apakah user tidak masuk ke dalam daftar location access, dan secara default boleh akses semua lokasi jika tidak ada dalam daftar lokasi
            If Dummy = 0 Then
                'secara default, klo ga ada di dalam daftar maka boleh akses semua lokasi
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode) = False Then
                    'user ybs tidak ada di dalam daftar, maka bisa akses semua lokasi
                    Split = False
                Else
                    'user ada di dalam daftar
                    DTAUserLocationAccess.FillBy_UserCode(TBLUserLocationAccess, Me.UserCode)
                    rowUserLocationAccess = TBLUserLocationAccess.Rows(0)
                    If rowUserLocationAccess.default_access = 0 Then
                        DefaultBoleh1 = True '-> klo ga ada di dalam daftar secara default boleh, maka fokus pada yang ga boleh aja
                    Else
                        DefaultBoleh1 = False '-> klo ga ada di dalam daftar secara ga default boleh, maka fokus pada yang boleh aja
                    End If
                    Split = True
                End If
            Else
                'secara default, jika tidak ada didalam daftar, maka tidak boleh akses semua lokasi
                If Me.ObjUser.ApakahMemilikiHakAksesLokasi(Me.UserCode) = False Then
                    'tidak ada di dalam daftar, maka tidak tidak bisa akses semua lokasi, kosongkan aja semua tabel qty
                    vTBLBoleh.Rows.Clear()
                    Exit Sub
                Else
                    'user ada di dalam daftar
                    DTAUserLocationAccess.FillBy_UserCode(TBLUserLocationAccess, Me.UserCode)
                    rowUserLocationAccess = TBLUserLocationAccess.Rows(0)
                    If rowUserLocationAccess.default_access = 0 Then
                        DefaultBoleh1 = True '-> klo ga ada di dalam daftar secara default boleh, maka fokus pada yang ga boleh aja
                    Else
                        DefaultBoleh1 = False '-> klo ga ada di dalam daftar secara ga default boleh, maka fokus pada yang boleh aja
                    End If
                    Split = True
                End If
            End If

            If Split = True Then
                If DefaultBoleh1 = True Then
                    'secara default boleh, maka fokus pada yang ga boleh aja -> isi daftar yang tidak boleh saja
                    Me.Split2Tabel_1(vTBLBoleh, vTBLQty_TdkBoleh, True)
                Else
                    'secara ga default boleh, maka fokus pada yang boleh aja -> isi daftar yang boleh saja
                    Me.Split2Tabel_1(vTBLBoleh, vTBLQty_TdkBoleh, False)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Split2Tabel_1(ByVal vTBLBoleh As dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable, ByVal vTBLTidakBoleh As dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable, ByVal DefaultValue As Boolean)
        Dim DTAUserDetail As New Core_Module.DseUserLocationPrivilegeSettingTableAdapters.tbl_setting_user_location_privilege_itemTableAdapter
        Dim TBLUserDetail As New Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemDataTable
        Dim rowUserDetail As Core_Module.DseUserLocationPrivilegeSetting.tbl_setting_user_location_privilege_itemRow
        Dim ListPrimaryLocationTdkBoleh As New List(Of Double)
        Dim ListPrimaryLocationBoleh As New List(Of Double)
        Dim vRowQty, vRowQty1 As dseTransaksiStockInOut_Sales.tblmaster_productqtyRow
        Dim Ketemu As Boolean = False

        Try
            DTAUserDetail.FillBy_UserCode(TBLUserDetail, Me.UserCode)
            For Each vRowQty In vTBLBoleh.Rows
                If ListPrimaryLocationTdkBoleh.Contains(vRowQty.Primary_Storage_Area) = False And _
                ListPrimaryLocationBoleh.Contains(vRowQty.Primary_Storage_Area) = False Then
                    Ketemu = False

                    'cek apakah ada kode lokasi cocok
                    For Each rowUserDetail In TBLUserDetail.Rows
                        If rowUserDetail.location_level = 1 Then
                            If rowUserDetail.primary_location = vRowQty.Primary_Storage_Area Then
                                If rowUserDetail.stock_out_sales = 1 Then
                                    'boleh
                                    ListPrimaryLocationBoleh.Add(vRowQty.Primary_Storage_Area)
                                Else
                                    'ga boleh
                                    ListPrimaryLocationTdkBoleh.Add(vRowQty.Primary_Storage_Area)
                                End If
                                Ketemu = True
                            End If
                        End If
                    Next
                    'cek apakah ada kode warehouse yang cocok
                    If Ketemu = False Then
                        For Each rowUserDetail In TBLUserDetail.Rows
                            If rowUserDetail.location_level = 0 Then
                                If rowUserDetail.primary_location = vRowQty.Primary_Warehouse Then
                                    If rowUserDetail.stock_out_sales = 1 Then
                                        'boleh
                                        ListPrimaryLocationBoleh.Add(vRowQty.Primary_Storage_Area)
                                    Else
                                        'ga boleh
                                        ListPrimaryLocationTdkBoleh.Add(vRowQty.Primary_Storage_Area)
                                    End If
                                    Ketemu = True
                                End If
                            End If
                        Next
                    End If

                    'klo ga ketemu di lokasi dan warehouse, maka secara default boleh
                    If Ketemu = False Then
                        If DefaultValue = True Then
                            'secara default boleh
                            ListPrimaryLocationBoleh.Add(vRowQty.Primary_Storage_Area)
                        Else
                            'secara default tidak boleh
                            ListPrimaryLocationTdkBoleh.Add(vRowQty.Primary_Storage_Area)
                        End If
                    End If
                End If
            Next

            'pindahkan semua yang tidak boleh
            For Each PrimaryLocation As Double In ListPrimaryLocationTdkBoleh
                vTBLBoleh.DefaultView.RowFilter = "[Primary Storage Area] = " + PrimaryLocation.ToString
                While vTBLBoleh.DefaultView.Count > 0
                    vRowQty = vTBLBoleh.DefaultView(0).Row
                    vRowQty1 = vTBLQty_TdkBoleh.Newtblmaster_productqtyRow
                    For gh As Integer = 0 To vTBLQty_TdkBoleh.Columns.Count - 1
                        vRowQty1.Item(gh) = vRowQty.Item(gh)
                    Next
                    vTBLQty_TdkBoleh.Rows.Add(vRowQty1)
                    vTBLBoleh.Rows.Remove(vRowQty)
                End While
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function LoadQty(ByRef vTBLQty As dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable, ByVal vPrimaryProduct As Double) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.PrimaryDO, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiDO_LoadDataTBLQty, Me.UserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.LoadQty1(vTBLQty, vPrimaryProduct)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Function LoadQty1(ByRef vTBLQty As dseTransaksiStockInOut_Sales.tblmaster_productqtyDataTable, ByVal vPrimaryProduct As Double) As Boolean
        Dim DtaQty As New dseTransaksiStockInOut_SalesTableAdapters.tblmaster_productqtyTableAdapter
        Dim rowQty As dseTransaksiStockInOut_Sales.tblmaster_productqtyRow

        Try
            vTBLQty.Rows.Clear()
            DtaQty.FillBy1(vTBLQty, vPrimaryProduct)
            For Each rowQty In vTBLQty.Rows
                rowQty.Urutan = Integer.MaxValue
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Try
            If CDec(Me.TxtQtyDeliveryDefault.Text) > CDec(Me.TxtQtyOrderBalanceDefault.Text) Then
                MsgBox("Maximum delivery qty is: " + Me.TxtQtyOrderBalanceDefault.Text + Me.UnitDefault, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.ObjCD.CekNonNegatif(Me.TxtQtyDeliveryDefault.Text) = False Or Me.ObjCD.CekNonNegatif(Me.TxtQtyDeliveryMulti.Text) = False Then
                MsgBox("Please specify positive number for delivery qty.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If MsgBox("Input data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.LangsungCLose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemDetail_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemDetail.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColQtyDelivery.Index Then
                Me.Buffer = Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemDetail_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemDetail.CellEndEdit
        Try
            If IsNothing(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = CDec(Me.Buffer)
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColQtyDelivery.Index Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = CDec(Me.Buffer)
                        Exit Sub
                    Else
                        'digit grouping
                        Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = CDec(Me.ObjCD.DigitGrouping(Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value, Me.DefPrecDefault))
                        'max value = qty available
                        If Me.dgvItemDetail.Item(Me.ColQtyAvailable.Index, e.RowIndex).Value < Me.dgvItemDetail.Item(Me.ColQtyDelivery.Index, e.RowIndex).Value Then
                            MsgBox("Maximum delivery qty is " + Me.ObjCD.DigitGrouping(Me.dgvItemDetail.Item(Me.ColQtyAvailable.Index, e.RowIndex).Value, Me.DefPrecDefault) + " " + Me.dgvItemDetail.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString, MsgBoxStyle.Information, "Information")
                            Me.dgvItemDetail.Item(e.ColumnIndex, e.RowIndex).Value = CDec(Me.Buffer)
                            Exit Sub
                        End If
                        'hitung qty delivery
                        Me.HitungQtyDelivery()
                    End If
                End If
            End If

            Me.SettingControlInput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungQtyDelivery()
        Dim rowDGV As DataGridViewRow
        Dim JmlDelivery As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemDetail.Rows
                JmlDelivery += CDec(rowDGV.Cells(Me.ColQtyDelivery.Index).Value)
            Next
            If JmlDelivery = CDec(Me.TxtQtyOrderBalanceDefault.Text) Then
                Me.TxtQtyDeliveryDefault.Text = Me.ObjCD.DigitGrouping(JmlDelivery, Me.DefPrecDefault)
                Me.TxtQtyDeliveryMulti.Text = Me.TxtQtyOrderBalanceMulti.Text
            Else
                Me.TxtQtyDeliveryDefault.Text = Me.ObjCD.DigitGrouping(JmlDelivery, Me.DefPrecDefault)
                Me.TxtQtyDeliveryMulti.Text = Me.ObjCD.DigitGrouping(Math.Round((JmlDelivery * Me.RatioMulti) / Me.RatioDefault, _
                Me.DefPrecMulti, MidpointRounding.AwayFromZero), Me.DefPrecMulti)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbLocationAccess_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbLocationAccess.LinkClicked
        Try
            Me.ObjUser.ShowHakAksesLokasi(Me.UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSelectAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSelectAll.LinkClicked
        Dim rowDGV As DataGridViewRow
        Dim JmlOB As Decimal = 0

        Try
            JmlOB = CDec(Me.TxtQtyOrderBalanceDefault.Text)
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowDGV In Me.dgvItemDetail.Rows
                If JmlOB > 0 Then
                    If JmlOB >= CDec(rowDGV.Cells(Me.ColQtyAvailable.Index).Value) Then
                        rowDGV.Cells(Me.ColQtyDelivery.Index).Value = rowDGV.Cells(Me.ColQtyAvailable.Index).Value
                        JmlOB -= rowDGV.Cells(Me.ColQtyDelivery.Index).Value
                    Else
                        rowDGV.Cells(Me.ColQtyDelivery.Index).Value = JmlOB
                        JmlOB = 0
                    End If
                Else
                    Exit For
                End If
            Next
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.HitungQtyDelivery()

            Me.SettingControlInput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUnselectAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUnselectAll.LinkClicked
        Dim rowDGV As DataGridViewRow
        Try
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowDGV In Me.dgvItemDetail.Rows
                rowDGV.Cells(Me.ColQtyDelivery.Index).Value = 0
            Next
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.TxtQtyDeliveryDefault.Text = Me.ObjCD.DigitGrouping(0, Me.DefPrecDefault)
            Me.TxtQtyDeliveryMulti.Text = Me.ObjCD.DigitGrouping(0, Me.DefPrecMulti)

            Me.SettingControlInput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemDetail_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemDetail.Sorted
        Dim rowDGV As DataGridViewRow
        Try
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each rowDGV In Me.dgvItemDetail.Rows
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
            Next
            Me.dgvItemDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemDetail.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul
    Private Sub BtnInputBarcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInputBarcode.Click
        Dim BarcodeID As String = ""
        Dim LabelID As String = ""

        Try
            If Me.ApakahDataItemBenar_Barcode(Me.InputBarcode, BarcodeID, LabelID, Me.BarcodeType) = True Then
                'tambahkan dgv
                Me.InputToDgv(BarcodeID, LabelID, Me.BarcodeType)
                'tambahkan counter
                Me.HitungCounter(EnumCounter.Tambah)
            Else
                Me.TxtScanHere.Clear()
            End If

            Me.SettingControlInput()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.InputBarcode = EnumInput.ManualInput
            Me.BarcodeType = ClsTransaksiBarcode_Scan.EnumBarcodeType.BarcodeSMESS
        End Try
    End Sub

    Private Sub BtnFindLocationSource_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindLocationSource.Click
        Try
            If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationSource.Text, Me.TxtLocationSource.Text, Me.TxtLocationSourceDesc.Text, _
            Me.TxtPrimaryWarehouseSource.Text, Me.TxtWarehouseSource.Text, Me.TxtWarehouseSourceDesc.Text, EnumMethod.LocFind) = True Then
                Me.TxtScanHere.Focus()

            Else
                Me.TxtLocationSource.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationSource_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationSource.GotFocus, TxtScanHere.GotFocus, TxtBarcodeID.GotFocus, TxtBarcodeLabel.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtLocationSourceDesc.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationSource_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationSource.LostFocus, TxtScanHere.LostFocus, TxtBarcodeID.LostFocus, TxtBarcodeLabel.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtLocationSourceDesc.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryLocationSource.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationSource_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocationSource.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationSource.Text, Me.TxtLocationSource.Text, Me.TxtLocationSourceDesc.Text, _
                Me.TxtPrimaryWarehouseSource.Text, Me.TxtWarehouseSource.Text, Me.TxtWarehouseSourceDesc.Text, EnumMethod.LocKeyDown) = True Then
                    Me.TxtScanHere.Focus()
                Else
                    Me.TxtLocationSource.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationSource_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationSource.Leave
        Try
            If Me.TxtLocationSource.Text <> "" Then
                If Me.GetLocationAndWarehouse(Me.TxtPrimaryLocationSource.Text, Me.TxtLocationSource.Text, Me.TxtLocationSourceDesc.Text, _
                Me.TxtPrimaryWarehouseSource.Text, Me.TxtWarehouseSource.Text, Me.TxtWarehouseSourceDesc.Text, EnumMethod.LocLeave) = True Then
                    Me.TxtScanHere.Focus()
                Else
                    Me.TxtLocationSource.Focus()
                End If
            Else
                Me.TxtPrimaryLocationSource.Text = -1
                Me.TxtLocationSource.Text = ""
                Me.TxtLocationSourceDesc.Text = ""
                Me.TxtPrimaryWarehouseSource.Text = -1
                Me.TxtWarehouseSource.Text = ""
                Me.TxtWarehouseSourceDesc.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtScanHere_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtScanHere.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.InputBarcode = EnumInput.ScanInput
                Me.BtnInputBarcode.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindBarcode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindBarcode.Click
        Dim i As New FrmBarcodeItem_Cari

        Try
            If Me.TxtBarcodeID.Text <> "" Then
                i.txtCari.Text = Me.TxtBarcodeID.Text
            End If

            i.ShowDialog()
            If i.BarcodeID > "" Then
                Me.TxtBarcodeID.Text = i.BarcodeID
                Me.TxtBarcodeLabel.Focus()
            Else
                Me.TxtBarcodeID.Clear()
                Me.TxtBarcodeID.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBarcode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBarcodeID.KeyDown
        Dim i As New FrmBarcodeItem_Cari

        Try
            If e.KeyValue = 13 Then
                If Me.ObjBarcode.ApakahSudahAda_BarcodeIDItem(Me.TxtBarcodeID.Text) = True Then
                    Me.TxtBarcodeLabel.Focus()
                Else
                    If Me.TxtBarcodeID.Text <> "" Then
                        i.txtCari.Text = Me.TxtBarcodeID.Text
                    End If

                    i.ShowDialog()
                    If i.BarcodeID > 0 Then
                        Me.TxtBarcodeID.Text = i.BarcodeID
                        Me.TxtBarcodeLabel.Focus()
                    Else
                        Me.TxtBarcodeID.Clear()
                        Me.TxtBarcodeID.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBarcode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBarcodeID.Leave
        Dim i As New FrmBarcodeItem_Cari

        Try
            If Me.TxtBarcodeID.Text <> "" Then
                If Me.ObjBarcode.ApakahSudahAda_BarcodeIDItem(Me.TxtBarcodeID.Text) = True Then
                    Me.TxtBarcodeLabel.Focus()
                Else
                    If Me.TxtBarcodeID.Text <> "" Then
                        i.txtCari.Text = Me.TxtBarcodeID.Text
                    End If

                    i.ShowDialog()
                    If i.BarcodeID > "" Then
                        Me.TxtBarcodeID.Text = i.BarcodeID
                        Me.btnOK.Focus()
                    Else
                        Me.TxtBarcodeID.Clear()
                        Me.TxtBarcodeID.Focus()
                    End If
                End If
            Else
                Me.TxtBarcodeID.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtBarcodeLabel_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBarcodeLabel.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtBarcodeLabel, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBarcodeLabel_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBarcodeLabel.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtBarcodeLabel, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBarcodeLabel_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBarcodeLabel.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtBarcodeLabel, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBarcodeLabel_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBarcodeLabel.Leave
        Try
            If String.IsNullOrEmpty(Me.TxtBarcodeLabel.Text) = True Then
                Exit Sub
            End If
            Me.ObjCD.Number_Leave(Me.TxtBarcodeLabel, sender, e, 0, Me.ObjCD)
            If CDec(Me.TxtBarcodeLabel.Text) < 10 And CDec(Me.TxtBarcodeLabel.Text) >= 0 And Me.TxtBarcodeLabel.Text.Length + 1 <= 2 Then
                Me.TxtBarcodeLabel.Text = "0" & CDec(Me.TxtBarcodeLabel.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.Col1Remark.Index Then
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

        Try
            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If e.ColumnIndex <> Me.Col1Remark.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit  
                        If e.ColumnIndex = Me.Col1Qty.Index Then
                            Me.EndEditQty(sender, e, row)
                        End If
                    Else
                        'tandai karakter aneh2
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                    End If
                Else
                    If e.ColumnIndex = Me.Col1Remark.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim DefPrec As Integer = 0
        Dim RatioMulti As Integer = 0
        Dim RatioDefault As Integer = 0
        Dim HitungQtyDefault As Decimal = 0

        Try
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                    End If
                End If

                'pastikan diisikan
                If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'pastikan angka
                If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please use a positive number for qty data.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If

                'digit grouping
                DefPrec = Me.dgvItem.Item(Me.Col1DecimalDefault.Index, e.RowIndex).Value
                RatioMulti = Me.dgvItem.Item(Me.Col1RatioMulti.Index, e.RowIndex).Value
                RatioDefault = Me.dgvItem.Item(Me.Col1RatioDefault.Index, e.RowIndex).Value
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                HitungQtyDefault = (Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value / RatioMulti) * RatioDefault
                Me.dgvItem.Item(Me.Col1QtyDefault.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(HitungQtyDefault, DefPrec)

                'pastikan positive
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please use a positive number for qty data.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.HitungCounter(EnumCounter.Kurang)

            Me.SettingControlInput()
            Me.SettingColumn_DgvBarcode()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungCounter(ByVal vCounter As EnumCounter)
        If IsNumeric(Me.lblCounter.Text) = False Then
            Me.lblCounter.Text = 0
        End If
        Dim JmlCounter As Integer = CInt(Me.lblCounter.Text)

        Try
            'If vCounter = EnumCounter.Tambah Then
            '    Me.lblCounter.Text = Me.ObjCD.DigitGrouping(JmlCounter + 1)
            'Else
            '    Me.lblCounter.Text = Me.ObjCD.DigitGrouping(JmlCounter - 1)
            'End If

            Me.lblCounter.Text = Me.dgvItem.Rows.Count
            Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.lblCounter.Text)
            If CInt(Me.lblCounter.Text) > 0 Then
                Me.lblCounter.Visible = True
            Else
                Me.lblCounter.Visible = False
            End If
            Me.HitungRowsAndQty()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetLocationAndWarehouse(ByRef vPrimaryLocation As Double, ByRef vLocationCode As String, ByRef vLocationDesc As String, _
    ByRef vPrimaryWarehouse As Double, ByRef vWarehouseCode As String, ByRef vWarehouseDesc As String, ByVal vMethod As EnumMethod) As Boolean
        Dim i As New Core_Module.FrmCariMasterStorageArea
        Dim GetWarehouse As Boolean = False

        Try
            If vMethod = EnumMethod.LocFind Then
                If vLocationCode <> "" Then
                    i.txtCari.Text = vLocationCode
                End If

                i.ShowDialog()
                If i.PrimaryStorageArea > 0 Then
                    vPrimaryLocation = i.PrimaryStorageArea
                    vLocationCode = Me.ObjLoc.GetStorageAreaCode(vPrimaryLocation)
                    vLocationDesc = Me.ObjLoc.GetDescription(vPrimaryLocation)
                    GetWarehouse = True
                End If
            Else
                If vLocationCode <> "" Then
                    If Me.ObjLoc.ApakahSudahAda(vLocationCode, Core_Module.clsMasterStorageArea.enumTransit.NotTransit) = True Then
                        vPrimaryLocation = Me.ObjLoc.GetPrimary(vLocationCode)
                        vLocationCode = Me.ObjLoc.GetStorageAreaCode(vPrimaryLocation)
                        vLocationDesc = Me.ObjLoc.GetDescription(vPrimaryLocation)
                        GetWarehouse = True
                    Else
                        If vLocationCode <> "" Then
                            i.txtCari.Text = vLocationCode
                        End If

                        i.ShowDialog()
                        If i.PrimaryStorageArea > 0 Then
                            vPrimaryLocation = i.PrimaryStorageArea
                            vLocationCode = Me.ObjLoc.GetStorageAreaCode(vPrimaryLocation)
                            vLocationDesc = Me.ObjLoc.GetDescription(vPrimaryLocation)
                            GetWarehouse = True
                        End If
                    End If
                End If
            End If

            If GetWarehouse = True Then
                vPrimaryWarehouse = Me.ObjLoc.GetPrimaryWarehouse(vPrimaryLocation)
                vWarehouseCode = Me.ObjWH.GetWarehouseCode(vPrimaryWarehouse)
                vWarehouseDesc = Me.ObjWH.GetDescription(vPrimaryWarehouse)

                Return True
            End If

            vPrimaryLocation = -1
            vLocationCode = ""
            vLocationDesc = ""
            vPrimaryWarehouse = -1
            vWarehouseCode = ""
            vWarehouseDesc = ""

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            vPrimaryLocation = -1
            vLocationCode = ""
            vLocationDesc = ""
            vPrimaryWarehouse = -1
            vWarehouseCode = ""
            vWarehouseDesc = ""
            Return False
        Finally
            i.Dispose()
        End Try
    End Function

    Private Sub HitungRowsAndQty()
        Dim TotalRows As Integer = 0
        Dim TotalQty As Decimal = 0
        Dim TotalQtyDefault As Decimal = 0

        Try
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                TotalRows += 1
                If IsNumeric(rd.Cells(Me.Col1Qty.Index).Value) = True Then
                    TotalQty += CDec(rd.Cells(Me.Col1Qty.Index).Value)
                End If
                If IsNumeric(rd.Cells(Me.Col1QtyDefault.Index).Value) = True Then
                    TotalQtyDefault += CDec(rd.Cells(Me.Col1QtyDefault.Index).Value)
                End If
            Next

            Me.TxtTotalRows.Text = TotalRows
            Me.TxtQtyMulti.Text = Me.ObjCD.BuangNolDibelakangDecimal(TotalQty)
            Me.TxtQtyDefault.Text = Me.ObjCD.BuangNolDibelakangDecimal(TotalQtyDefault)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar_Barcode(ByVal vInput As EnumInput, ByRef vBarcodeID As String, ByRef vLabelID As String, _
    ByRef vBarcodeType As ClsTransaksiBarcode_Scan.EnumBarcodeType) As Boolean
        Dim rd As DataGridViewRow
        Dim cScane As New ClsTransaksiBarcode_Scan

        Try
            vBarcodeType = ClsTransaksiBarcode_Scan.EnumBarcodeType.BarcodeSMESS

            'lokasi sumber tidak boleh kosong
            If Me.TxtPrimaryLocationSource.Text = -1 Or Me.TxtLocationSource.Text = "" Then
                MsgBox("Please insert source location.", MsgBoxStyle.Information, "Information")
                Me.TxtLocationSource.Focus()
                Return False
            End If

            If vInput = EnumInput.ScanInput Then
                If Me.TxtScanHere.Text = "" Then
                    'MsgBox("Please scan barcode first", MsgBoxStyle.Information, "Information")
                    'Me.TxtScanHere.Focus()
                    Me.TxtBarcodeID.Focus()
                    Return False
                End If

                'pastikan format barcode yang di scan benar
                If Me.TxtScanHere.Text.Length = 13 Then
                    vBarcodeID = Me.TxtScanHere.Text.Substring(0, 10)
                    vLabelID = Me.TxtScanHere.Text.Substring(11, 2)
                    If IsNumeric(vBarcodeID) = False Then
                        'MsgBox("Barcode format is wrong.", MsgBoxStyle.Information, "Information")
                        'Return False
                        GoTo SetBarcodeLMB
                    End If
                    If IsNumeric(vLabelID) = False Then
                        'MsgBox("Barcode format is wrong.", MsgBoxStyle.Information, "Information")
                        'Return False
                        GoTo SetBarcodeLMB
                    End If

                    For Each rd In Me.dgvItem.Rows 'noted coding
                        If rd.Cells(Me.Col1Barcode.Index).Value = vBarcodeID And rd.Cells(Me.Col1LabelID.Index).Value = vLabelID Then
                            MsgBox("Duplication in barcode id, you can't add the following barcode (" & vBarcodeID & "-" & vLabelID & ")", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next

                    Dim QtyPrint As Decimal = cScane.GetQtyPrint(vBarcodeID) 'noted coding
                    If CDec(vLabelID) > QtyPrint Then ' max 5 qty print
                        MsgBox("Maximum print qty is:" & QtyPrint, MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Else
                    'MsgBox("Barcode format is wrong.", MsgBoxStyle.Information, "Information")
                    'Return False
SetBarcodeLMB:
                    vBarcodeID = Me.TxtScanHere.Text
                    vLabelID = "00"
                    vBarcodeType = ClsTransaksiBarcode_Scan.EnumBarcodeType.BarcodeLMB

                    For Each rd In Me.dgvItem.Rows 'noted coding
                        If rd.Cells(Me.Col1Barcode.Index).Value = vBarcodeID Then
                            MsgBox("Duplication in barcode id, you can't add the following barcode (" & vBarcodeID & ")", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Next
                End If
            Else
                If Me.TxtBarcodeID.Text = "" Then 'noted coding
                    MsgBox("Please insert barcode id.", MsgBoxStyle.Information, "Information")
                    Me.TxtBarcodeID.Focus()
                    Return False
                End If

                If Me.TxtBarcodeLabel.Text = "" Then
                    MsgBox("Please insert label id.", MsgBoxStyle.Information, "Information")
                    Me.TxtBarcodeLabel.Focus()
                    Return False
                End If
                If Me.ObjCD.CekBilanganPositif(Me.TxtBarcodeLabel.Text) = False Then
                    MsgBox("Please specify positive number for label id.", MsgBoxStyle.Information, "Information")
                    Me.TxtBarcodeLabel.Focus()
                    Return False
                End If

                vBarcodeID = Me.TxtBarcodeID.Text
                vLabelID = Me.TxtBarcodeLabel.Text

                For Each rd In Me.dgvItem.Rows
                    If rd.Cells(Me.Col1Barcode.Index).Value = vBarcodeID And rd.Cells(Me.Col1LabelID.Index).Value = vLabelID Then
                        MsgBox("Duplication in barcode id, you can't add the following barcode (" & vBarcodeID & "-" & vLabelID & ")", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Next

                Dim QtyPrint As Decimal = cScane.GetQtyPrint(Me.TxtBarcodeID.Text)
                If CDec(Me.TxtBarcodeLabel.Text) > QtyPrint Then
                    MsgBox("Maximum print qty is:" & QtyPrint, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlLabelByBarcode(ByVal vBarCode As String) As Integer
        Dim Jml As Integer = 0
        Try
            For Each r As DataGridViewRow In Me.dgvItem.Rows
                If r.Cells(Me.Col1Barcode.Index).Value = vBarCode Then
                    Jml += 1
                End If
            Next
            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub InputToDgv(ByVal vBarcodeID As String, ByVal vLabelID As String, _
    ByRef vBarcodeType As ClsTransaksiBarcode_Scan.EnumBarcodeType)
        Dim rd As DataGridViewRow
        Dim RatioStr As String = ""
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjBS As New ClsTransaksiBarcode_Scan
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim QtyDefault As Decimal = 0

        Try
            If vBarcodeType = ClsTransaksiBarcode_Scan.EnumBarcodeType.BarcodeSMESS Then
                Dim dta As New DseBarcode_PrintTableAdapters.DataTable1TableAdapter
                Dim tbl As New DseBarcode_Print.DataTable1DataTable
                Dim row As DseBarcode_Print.DataTable1Row
                'dta.Connection.ConnectionString = Core_Module.GetConstring
                If dta.FillBy_Barcode(tbl, vBarcodeID) > 0 Then
                    row = tbl.Rows(0)
                    With row
                        If .ratio_default > 0 And .ratio_unit_multi > 0 Then
                            QtyDefault = Me.ObjCD.DigitGrouping(Math.Round((1 / .ratio_unit_multi) * .ratio_default, .def_prec_unit, MidpointRounding.AwayFromZero), .def_prec_unit)
                            'QtyDefault = Me.ObjCD.DigitGrouping((CDec(rd.Cells(Me.Col1Qty.Index).Value) / .ratio_unit_multi) * .ratio_default, .def_prec_unit)
                        Else
                            QtyDefault = Me.ObjCD.DigitGrouping(1, .def_prec_unit)
                        End If

                        If Me.UpdateQtyDeliverDOSalesByBarcode(.Product_Code, .Batch_No, Me.TxtLocationSource.Text, .Unit_Code, QtyDefault) = True Then
                            Me.dgvItem.Rows.Add()
                            rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                            rd.Cells(Me.Col1No.Index).Value = Me.dgvItem.RowCount
                            rd.Cells(Me.Col1BarcodeType.Index).Value = Me.ObjBarcodeScanItem.GetBarcodeType_String(ClsTransaksiBarcode_Scan.EnumBarcodeType.BarcodeSMESS) 'add barcode lmb
                            rd.Cells(Me.Col1Barcode.Index).Value = vBarcodeID
                            rd.Cells(Me.Col1LabelID.Index).Value = vLabelID
                            rd.Cells(Me.Col1CodeNo.Index).Value = .Product_Code
                            rd.Cells(Me.Col1Description.Index).Value = .Product_Desc
                            rd.Cells(Me.Col1Qty.Index).Value = Me.ObjCD.DigitGrouping(1)
                            rd.Cells(Me.Col1Unit.Index).Value = .Unit_Code_Multi
                            rd.Cells(Me.Col1QtyDefault.Index).Value = Me.ObjCD.DigitGrouping(QtyDefault, .def_prec_unit)
                            rd.Cells(Me.Col1UnitDefault.Index).Value = .Unit_Code
                            rd.Cells(Me.Col1RatioMulti.Index).Value = Me.ObjCD.DigitGrouping(.ratio_unit_multi, .def_prec_unit_multi)
                            rd.Cells(Me.Col1RatioDefault.Index).Value = Me.ObjCD.DigitGrouping(.ratio_default, .def_prec_unit)
                            RatioStr = Me.ObjCD.DigitGrouping(.ratio_unit_multi, .def_prec_unit_multi) & " " & .Unit_Code_Multi & " = " & Me.ObjCD.DigitGrouping(rd.Cells(Me.Col1RatioDefault.Index).Value, .def_prec_unit) & " " & .Unit_Code
                            If .primary_unit_multi <> .primary_unit_default Then
                                rd.Cells(Me.Col1Ratio.Index).Value = RatioStr
                            Else
                                rd.Cells(Me.Col1Ratio.Index).Value = "-"
                            End If
                            rd.Cells(Me.Col1Batch.Index).Value = .Batch_No
                            Dim Expiry As Date = Me.GetExpiryControl(.primary_code_no, .Batch_No)
                            If Expiry = #1/1/1950# Then
                                rd.Cells(Me.Col1ExpiredDate.Index).Value = "-"
                            Else
                                rd.Cells(Me.Col1ExpiredDate.Index).Value = Format(Expiry, "dd MMM yyyy")
                            End If
                            rd.Cells(Me.Col1Remark.Index).Value = ""
                            rd.Cells(Me.Col1PrimaryCodeNo.Index).Value = .primary_code_no
                            rd.Cells(Me.Col1PrimaryUnits.Index).Value = .primary_unit_multi
                            rd.Cells(Me.Col1PrimaryUnitDefault.Index).Value = .primary_unit_default
                            rd.Cells(Me.Col1Decimal.Index).Value = .def_prec_unit_multi
                            rd.Cells(Me.Col1DecimalDefault.Index).Value = .def_prec_unit
                            rd.Cells(Me.Col1LocationControl.Index).Value = .Location_Controlled
                            rd.Cells(Me.Col1IsService.Index).Value = .Type

                            rd.Cells(Me.Col1LocationSourcePrimary.Index).Value = Me.TxtPrimaryLocationSource.Text
                            rd.Cells(Me.Col1LocationSource.Index).Value = Me.TxtLocationSource.Text
                            rd.Cells(Me.Col1WareHouseSource.Index).Value = Me.TxtWarehouseSource.Text
                        End If
                    End With
                Else
                    MsgBox("Barcode not found", MsgBoxStyle.Information, "Information")
                End If
            Else
                Dim tLMB As New DseTransaksiBarcode_Scan.tblBarcode_lmbDataTable
                Dim rLMB As DseTransaksiBarcode_Scan.tblBarcode_lmbRow
                Dim UnitCode As String = ""
                Dim BatchStr As String = ""
                ObjBS.IsiTblBarcodeLMB(vBarcodeID, tLMB)
                If tLMB.Rows.Count > 0 Then
                    rLMB = tLMB.Rows(0)
                    With rLMB
                        Dim ij As New FrmBarcode_Scan_Batch
                        ij.Barcode = vBarcodeID
                        ij.ProductCode = .Product_Code
                        ij.PrimaryProduct = .primary_code_no
                        ij.BatchOption = Me.BatchOption
                        ij.PrimaryLocation = Me.TxtPrimaryLocationSource.Text
                        'konfirmasi ulang terkait pemilihan cara get batch
                        If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                            BatchStr = ij.BatchSelect
                        Else
                            Exit Sub
                        End If

                        UnitCode = ObjUnit.GetUnitCode(.primary_unit)
                        If Me.UpdateQtyDeliverDOSalesByBarcode(.Product_Code, BatchStr, Me.TxtLocationSource.Text, UnitCode, 1) = True Then
                            Me.dgvItem.Rows.Add()
                            rd = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                            rd.Cells(Me.Col1No.Index).Value = Me.dgvItem.RowCount
                            rd.Cells(Me.Col1BarcodeType.Index).Value = Me.ObjBarcodeScanItem.GetBarcodeType_String(ClsTransaksiBarcode_Scan.EnumBarcodeType.BarcodeLMB) 'add barcode lmb
                            rd.Cells(Me.Col1Barcode.Index).Value = vBarcodeID
                            rd.Cells(Me.Col1LabelID.Index).Value = vLabelID
                            rd.Cells(Me.Col1CodeNo.Index).Value = .Product_Code
                            rd.Cells(Me.Col1Description.Index).Value = .Description
                            rd.Cells(Me.Col1Qty.Index).Value = Me.ObjCD.DigitGrouping(1)
                            rd.Cells(Me.Col1Unit.Index).Value = UnitCode
                            rd.Cells(Me.Col1QtyDefault.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.Col1Qty.Index).Value), .Dec_Default)
                            rd.Cells(Me.Col1UnitDefault.Index).Value = UnitCode
                            rd.Cells(Me.Col1RatioMulti.Index).Value = Me.ObjCD.DigitGrouping(1, .Dec_Default)
                            rd.Cells(Me.Col1RatioDefault.Index).Value = Me.ObjCD.DigitGrouping(1, .Dec_Default)
                            rd.Cells(Me.Col1Ratio.Index).Value = "-"
                            rd.Cells(Me.Col1Batch.Index).Value = BatchStr
                            Dim Expiry As Date = Me.GetExpiryControl(.primary_code_no, BatchStr)
                            If Expiry = #1/1/1950# Then
                                rd.Cells(Me.Col1ExpiredDate.Index).Value = "-"
                            Else
                                rd.Cells(Me.Col1ExpiredDate.Index).Value = Format(Expiry, "dd MMM yyyy")
                            End If
                            rd.Cells(Me.Col1Remark.Index).Value = ""
                            rd.Cells(Me.Col1PrimaryCodeNo.Index).Value = .primary_code_no
                            rd.Cells(Me.Col1PrimaryUnits.Index).Value = .primary_unit
                            rd.Cells(Me.Col1PrimaryUnitDefault.Index).Value = .primary_unit
                            rd.Cells(Me.Col1Decimal.Index).Value = .Dec_Default
                            rd.Cells(Me.Col1DecimalDefault.Index).Value = .Dec_Default
                            rd.Cells(Me.Col1LocationControl.Index).Value = .Location_Controlled
                            rd.Cells(Me.Col1IsService.Index).Value = .Type

                            rd.Cells(Me.Col1LocationSourcePrimary.Index).Value = Me.TxtPrimaryLocationSource.Text
                            rd.Cells(Me.Col1LocationSource.Index).Value = Me.TxtLocationSource.Text
                            rd.Cells(Me.Col1WareHouseSource.Index).Value = Me.TxtWarehouseSource.Text
                        End If
                    End With
                Else
                    MsgBox("Barcode not found", MsgBoxStyle.Information, "Information")
                End If
            End If

            Me.TxtScanHere.Clear()
            Me.TxtBarcodeID.Clear()
            Me.TxtBarcodeLabel.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.SettingColumn_DgvBarcode()
        End Try
    End Sub

    Private Function ApakahExpiryControl(ByVal vPrimaryProduct As Double) As Boolean
        Dim d As New DseBarcodeInputTableAdapters.tblmaster_productTableAdapter
        Dim Ret As New Nullable(Of Boolean)
        Try
            Ret = d.ScalarQuery_GetExpiryDateControlled(vPrimaryProduct)
            If Ret.HasValue = True Then
                If Ret.Value = True Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            d.Dispose()
        End Try
    End Function

    Private Function GetExpiryControl(ByVal vPrimaryProduct As Double, ByVal vBatch As String) As Date
        Dim d As New DseBarcodeInputTableAdapters.tblmaster_productTableAdapter
        Dim d2 As New DseBarcodeInputTableAdapters.tblmaster_productTableAdapter
        Dim Ret As Nullable(Of Date) = Date.MinValue
        Try
            If ApakahExpiryControl(vPrimaryProduct) = True Then
                Ret = d.ScalarQuery_GetExpiryDateInOut(vPrimaryProduct, vBatch)
                If Ret.HasValue = True Then
                    If Ret.Value <> Date.MinValue Then
                        Return Ret.Value
                    End If
                Else
                    Ret = d2.ScalarQuery_GetExpiryDateBundlingAfter(vPrimaryProduct, vBatch)
                    If Ret.HasValue = True Then
                        If Ret.Value <> Date.MinValue Then
                            Return Ret.Value
                        End If
                    End If
                End If
            End If

            Return CDate("1/1/1950")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            d.Dispose()
            d2.Dispose()
        End Try
    End Function

    'aul new
    'dianggil oleh select/unselect all, dgv do sales edit, btn input barcode, delete rows dgv barcode
    Private Sub SettingControlInput()
        Dim QtyDO_InputManual As Decimal = 0
        Dim QtyDO_InputBarcode As Decimal = 0

        Try
            For Each rd As DataGridViewRow In Me.dgvItemDetail.Rows
                QtyDO_InputManual += rd.Cells(Me.ColQtyDelivery.Index).Value
            Next
            QtyDO_InputBarcode = Me.dgvItem.Rows.Count

            If QtyDO_InputManual > 0 And QtyDO_InputBarcode = 0 Then 'manual
                '1 jika do ada qty deliver dan barcode nothing maka tidak bisa input barcode
                Me.GroupBox1.Enabled = False
                Me.llbSelectAll.Enabled = True
                Me.llbUnselectAll.Enabled = True
                Me.dgvItemDetail.ReadOnly = False
            ElseIf QtyDO_InputManual = 0 And QtyDO_InputBarcode = 0 Then 'barcode/manual
                '2 jika do nothing dan barcode nothing maka bisa barcode/manual
                Me.GroupBox1.Enabled = True
                Me.llbSelectAll.Enabled = True
                Me.llbUnselectAll.Enabled = True
                Me.dgvItemDetail.ReadOnly = False
            ElseIf QtyDO_InputBarcode > 0 Then 'barcode
                '3 jika barcode ada baris maka tidak bisa manual
                Me.GroupBox1.Enabled = True
                Me.llbSelectAll.Enabled = False
                Me.llbUnselectAll.Enabled = False
                Me.dgvItemDetail.ReadOnly = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SettingColumn_DgvBarcode()
        Dim ApakahUnitMulti As Boolean = False
        Dim ApakahExpDate As Boolean = False

        Try
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                If rd.Cells(Me.Col1Unit.Index).Value <> rd.Cells(Me.Col1UnitDefault.Index).Value Then
                    ApakahUnitMulti = True
                End If
                If IsDate(rd.Cells(Me.Col1ExpiredDate.Index).Value) = True Then
                    ApakahExpDate = True
                End If

                If ApakahUnitMulti = True And ApakahExpDate = True Then
                    Exit For
                End If
            Next

            If ApakahUnitMulti = True Then
                Me.Col1QtyDefault.Visible = True
                Me.Col1UnitDefault.Visible = True
                Me.Col1Ratio.Visible = True
            Else
                Me.Col1QtyDefault.Visible = False
                Me.Col1UnitDefault.Visible = False
                Me.Col1Ratio.Visible = False
            End If
            If ApakahExpDate = True Then
                Me.Col1ExpiredDate.Visible = True
            Else
                Me.Col1ExpiredDate.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function UpdateQtyDeliverDOSalesByBarcode(ByVal vProductCode As String, ByVal vBatch As String, _
    ByVal vLocation As String, ByVal vUnitDef As String, ByVal vQtyDef As Decimal) As Boolean
        Dim Ketemu As Boolean = False
        Dim QtyAvailable As Decimal = 0
        Dim QtyDeliver As Decimal = 0

        Try
            Me.dgvItemDetail.ClearSelection()
            If vQtyDef = 0 Then
                MsgBox("Qty barcode default is 0, can't input this barcode", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1 kombinasi product + batch + location + unit dan tidak bisa lebih besar dari qty available
            For Each rd As DataGridViewRow In Me.dgvItemDetail.Rows
                QtyAvailable = rd.Cells(Me.ColQtyAvailable.Index).Value
                QtyDeliver = rd.Cells(Me.ColQtyDelivery.Index).Value

                If Me.TxtCodeNo.Text = vProductCode And rd.Cells(Me.ColBatchNo.Index).Value = vBatch _
                And rd.Cells(Me.ColLocation.Index).Value = vLocation _
                And rd.Cells(Me.ColUnit.Index).Value = vUnitDef Then
                    If (QtyDeliver + vQtyDef) > QtyAvailable Then
                        MsgBox("Maximum delivery qty is " + Me.ObjCD.DigitGrouping(QtyAvailable, Me.DefPrecDefault) + " " + rd.Cells(Me.ColUnit.Index).Value.ToString, MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If (CDec(Me.TxtQtyDeliveryDefault.Text) + vQtyDef) > CDec(Me.TxtQtyOrderBalanceDefault.Text) Then
                        MsgBox("Maximum delivery qty is: " + Me.TxtQtyOrderBalanceDefault.Text + Me.UnitDefault, MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    rd.Cells(Me.ColQtyDelivery.Index).Value = CDec(Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColQtyDelivery.Index).Value) + vQtyDef, Me.DefPrecDefault))
                    rd.Selected = True
                    Ketemu = True
                End If

                If Ketemu = True Then
                    Exit For
                End If
            Next

            If Ketemu = False Then
                Dim MsgException As String = ""
                If String.IsNullOrEmpty(vBatch) = False Then
                    MsgException &= ", batch no: " & vBatch
                End If
                If String.IsNullOrEmpty(vLocation) = False Then
                    MsgException &= ", location: " & vLocation
                End If
                If String.IsNullOrEmpty(vUnitDef) = False Then
                    MsgException &= ", unit: " & vUnitDef
                End If

                MsgBox("Product code: " & vProductCode & MsgException & " not found in item DO Sales.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'hitung qty delivery
            Me.HitungQtyDelivery()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Dim ProdCode As String = ""
        Dim ProdBatch As String = ""
        Dim ProdLoc As String = ""
        Dim ProdUnit As String = ""
        Dim ProdQty As String = ""

        Try
            ProdCode = Me.dgvItem(Me.Col1CodeNo.Index, e.Row.Index).Value
            ProdBatch = Me.dgvItem(Me.Col1Batch.Index, e.Row.Index).Value
            ProdLoc = Me.dgvItem(Me.Col1LocationSource.Index, e.Row.Index).Value
            ProdUnit = Me.dgvItem(Me.Col1UnitDefault.Index, e.Row.Index).Value
            ProdQty = CDec(Me.dgvItem(Me.Col1QtyDefault.Index, e.Row.Index).Value) * -1
            Me.UpdateQtyDeliverDOSalesByBarcode(ProdCode, ProdBatch, ProdLoc, ProdUnit, ProdQty)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        Try
            If Me.TabControl1.SelectedIndex = 1 Then
                If Me.ApakahMemilikiAksesAddByBarcode = False Then
                    Me.TabControl1.SelectedIndex = 0
                    'MsgBox("You don't have authorization form this information.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim ApaShiftKey As Boolean = False
    Private Sub BtnInputBarcode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BtnInputBarcode.KeyDown
        'Try
        '    If (e.Modifiers = Keys.Shift And e.KeyCode = Keys.Tab) Then
        '        ApaShiftKey = True
        '    Else
        '        ApaShiftKey = False
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub BtnInputBarcode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnInputBarcode.Leave
        'Try
        '    'If Me.ApaShiftKey = True Then
        '    '    Me.TxtBarcodeLabel.Focus()
        '    'Else
        '    '    Me.llbLocationAccess.Focus()
        '    'End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub btnCancel_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnCancel.KeyDown
        Try
            If e.Modifiers = Keys.Shift Then
                ApaShiftKey = True
            Else
                ApaShiftKey = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Leave
        Try
            If Me.ApaShiftKey = True Then
                Me.btnOK.Focus()
            Else
                If Me.TxtLocationSource.Visible = True Then
                    Me.TxtLocationSource.Focus()
                Else
                    Me.llbLocationAccess.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class