Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing_DetailAkun
    Public vTipeTrasaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
    Public vPrimaryTipeAdjustment As Double = -1
    Public IsDebet As Boolean = True
    Public vDGVItem As DataGridView
    Public vFrmAcuan As FrmTransaksiMonthlyClosing2
    Public vPrimaryAkun As Double = -1
    Public TBLItemDetail As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    Public vPeriod As Date = Today
    Public vUserCode As String = ""
    Public vTblSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable

    Dim Report1 As New CrpTransaksiMontlyClosing_DetailAkun
    Dim Report2 As New CrpTransaksiMontlyClosing_RekapAkun
    Dim Report3 As New CrpTransaksiMontlyClosing_DetailAkun2
    Dim Report4 As New CrpTransaksiMontlyClosing_RekapAkun2

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjAkun As New AccModule.ClsMasterAkun
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
    Dim rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    Dim rowSetting As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentRow
    Dim DV As DataRowView
    Dim DS As New DseTransaksiMonthlyClosing2
    Dim TBLCetak As DseTransaksiMonthlyClosing2.TblDetailProductDataTable = DS.TblDetailProduct
    Dim rowCetak As DseTransaksiMonthlyClosing2.TblDetailProductRow
    Dim DTAAOP As New DseTransaksiMonthlyClosing2TableAdapters.TblTransAOPTableAdapter
    Dim TBLAOP As New DseTransaksiMonthlyClosing2.TblTransAOPDataTable
    Dim rowAOP As DseTransaksiMonthlyClosing2.TblTransAOPRow
    Dim TglSampai, TglDari As Date
    Dim DTASR As New DseTransaksiMonthlyClosing2TableAdapters.TblSR2TableAdapter
    Dim TBLSR As New DseTransaksiMonthlyClosing2.TblSR2DataTable
    Dim rowSR As DseTransaksiMonthlyClosing2.TblSR2Row
    Dim DTABundling As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockbundlingTableAdapter
    Dim TBLBundling As New DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingDataTable
    Dim rowBundling As DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingRow
    Dim DTAStockInOut As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockinoutTableAdapter
    Dim TBLStockInOut As New DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutDataTable
    Dim rowStockInOut As DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutRow
    Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    Dim DTAStockTake As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_stock_takingTableAdapter
    Dim TBLStockTake As New DseTransaksiMonthlyClosing2.tbltrans_stock_takingDataTable
    Dim rowStockTake As DseTransaksiMonthlyClosing2.tbltrans_stock_takingRow
    Dim DTAInvoice As New DseTransaksiMonthlyClosing2TableAdapters.TblInvoiceTableAdapter
    Dim TBLInvoice As New DseTransaksiMonthlyClosing2.TblInvoiceDataTable
    Dim rowInvoice As DseTransaksiMonthlyClosing2.TblInvoiceRow
    Dim DTAPR As New DseTransaksiMonthlyClosing2TableAdapters.TblPurchaseReturnTableAdapter
    Dim TBLPR As New DseTransaksiMonthlyClosing2.TblPurchaseReturnDataTable
    Dim rowPR As DseTransaksiMonthlyClosing2.TblPurchaseReturnRow
    Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
    Dim Aktif As Boolean = False

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_DetailAkun_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            'simpan data setting ke history
            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowFormatDetailAkun, _
            Me.cboFormat.SelectedIndex, Me.vUserCode)
            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowProductDetail, _
            Me.cboProductDetail.SelectedIndex, Me.vUserCode)
            If Me.chkShowReferenceParty.Checked = True Then
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowReferenceParty, _
                "1", Me.vUserCode)
            Else
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowReferenceParty, _
                "0", Me.vUserCode)
            End If
            If Me.chkUseSupplierInvoiceNumber.Checked = True Then
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowSupplierInvoiceNumber, _
                "1", Me.vUserCode)
            Else
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowSupplierInvoiceNumber, _
                "0", Me.vUserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_DetailAkun_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Dummy As String = ""

        Try
            Me.cboProductDetail.SelectedIndex = 0

            'isikan data dari history
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowFormatDetailAkun, Me.vUserCode)
            If Dummy = "" Then
                Me.cboFormat.SelectedIndex = 0
            Else
                Me.cboFormat.SelectedIndex = Dummy
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowProductDetail, Me.vUserCode)
            If Dummy = "" Then
                Me.cboProductDetail.SelectedIndex = 0
            Else
                Me.cboProductDetail.SelectedIndex = Dummy
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowReferenceParty, Me.vUserCode)
            If Dummy = "" Then
                Me.chkShowReferenceParty.Checked = False
            Else
                If Dummy = "1" Then
                    Me.chkShowReferenceParty.Checked = True
                Else
                    Me.chkShowReferenceParty.Checked = False
                End If
            End If
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowSupplierInvoiceNumber, Me.vUserCode)
            If Dummy = "" Then
                Me.chkUseSupplierInvoiceNumber.Checked = False
            Else
                If Dummy = "1" Then
                    Me.chkUseSupplierInvoiceNumber.Checked = True
                Else
                    Me.chkUseSupplierInvoiceNumber.Checked = False
                End If
            End If

            Me.Icon = My.Resources.desktop_icon_3_fw


            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTABundling.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAStockInOut.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAStockTake.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAInvoice.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAPR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString

            TglSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.vPeriod)
            TglDari = Me.ObjCD.GetFirstDateofTheMonth(Me.vPeriod)

            Me.LoadData()
            Me.Aktif = True

            Me.SetKontrol()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
        Dim TipeAdjustmentStr As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.chkShowReferenceParty.Enabled = False
            Me.chkUseSupplierInvoiceNumber.Enabled = False
            Me.cboProductDetail.Enabled = False
            Me.btnLoad.Enabled = False
            Me.cboFormat.Enabled = False

            'kosongkan data tabel cetak
            Me.TBLCetak.Rows.Clear()

            'fill data item detail yang memenuhi kriteria: berdasarkan: tipe item, primary akun rujukan, tipe transaksi asal
            If Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                '1. tipe item stock out, tipe transaksi asal sales
                '   debet: akun cogs
                '   kredit: akun inventory
                Me.LoadDataOutSales()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                '2. tipe item: stock out, tipe transaksi asal bundling
                '   kredit: akun inventory
                Me.LoadDataOutBundling()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                '7. tipe item: stock in, tipe transaksi asal bundling
                '   debet: akun inventory
                Me.LoadDataInBundling()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                '5. tipe item: stock out, tipe transaksi asal stock take
                '   debet: akun stock taking expense
                '   kredit: akun inventory
                Me.LoadDataOutStockTake()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                '9. tipe item: stock in, tipe transaksi asal stock take
                '   debet: akun inventory
                '   kredit: akun stock taking expense
                Me.LoadDataInStockTake()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                '6. tipe item: stock in, tipe transaksi asal sales return
                '   debet: akun inventory
                '   kredit: akun cogs
                Me.LoadDataInSalesReturn()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                '8. tipe item: stock in, tipe transaksi asal adjust
                '   debet: akun inventory
                '   kredit: akun adjustment expense
                Me.LoadDataInAdjust()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                '2. tipe item stock out, tipe trasnaksi asal adjust
                '   debet: akun adjustment expense
                '   kredit: akun inventory
                Me.LoadDataOutAdjust()
            ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                '4. tipe item: stock out, tipe transaksi asal purchase return
                '   kalo yang ini agak ribet, harus hati2
                Me.LoadDataOutPurchaseReturn()
                Exit Sub
            End If


            'tampilkan laporan
            'show report
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.Report2.SetDataSource(Me.DS)
                Me.Report2.SetParameterValue(Me.Report2.Parameter_DefPrecAmount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                Me.Report2.SetParameterValue(Me.Report2.Parameter_Period.ParameterFieldName, Me.vPeriod)
                Me.Report2.SetParameterValue(Me.Report2.Parameter_Account.ParameterFieldName, Me.ObjAkun.GetAccountCode(Me.vPrimaryAkun) + _
                " -" + Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.vPrimaryAkun))
                If Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "Out Sales")
                    If Me.IsDebet = True Then
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "Out Bundling")
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "In Bundling")
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "Out Stock Take")
                    If Me.IsDebet = True Then
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "In Stock Take")
                    If Me.IsDebet = True Then
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                    Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "In Sales Return")
                    If Me.IsDebet = True Then
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                    If Me.vPrimaryTipeAdjustment <> -1 Then
                        TipeAdjustmentStr = Me.ObjStockInOut.GetDescriptionSettingAdjustment(Me.vPrimaryTipeAdjustment)
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "In Adjust - " + TipeAdjustmentStr)
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "In Adjust")
                    End If
                    If Me.IsDebet = True Then
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                    If Me.vPrimaryTipeAdjustment <> -1 Then
                        TipeAdjustmentStr = Me.ObjStockInOut.GetDescriptionSettingAdjustment(Me.vPrimaryTipeAdjustment)
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "Out Adjust - " + TipeAdjustmentStr)
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_Transaction_Type.ParameterFieldName, "Out Adjust")
                    End If
                    If Me.IsDebet = True Then
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    Else
                        Me.Report2.SetParameterValue(Me.Report2.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    End If
                End If
                Me.crv1.ReportSource = Me.Report2
            Else
                Me.Report1.SetDataSource(Me.DS)
                Me.Report1.SetParameterValue(Me.Report1.Parameter_DefPrecAmount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                Me.Report1.SetParameterValue(Me.Report1.Parameter_Period.ParameterFieldName, Me.vPeriod)
                Me.Report1.SetParameterValue(Me.Report1.Parameter_Account.ParameterFieldName, Me.ObjAkun.GetAccountCode(Me.vPrimaryAkun) + _
                " -" + Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.vPrimaryAkun))
                If Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "Out Sales")
                    If Me.IsDebet = True Then
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "Out Bundling")
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "In Bundling")
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "Out Stock Take")
                    If Me.IsDebet = True Then
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "In Stock Take")
                    If Me.IsDebet = True Then
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                    If Me.vPrimaryTipeAdjustment <> -1 Then
                        TipeAdjustmentStr = Me.ObjStockInOut.GetDescriptionSettingAdjustment(Me.vPrimaryTipeAdjustment)
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "In Adjust - " + TipeAdjustmentStr)
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "In Adjust")
                    End If
                    If Me.IsDebet = True Then
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                    If Me.vPrimaryTipeAdjustment <> -1 Then
                        TipeAdjustmentStr = Me.ObjStockInOut.GetDescriptionSettingAdjustment(Me.vPrimaryTipeAdjustment)
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "Out Adjust - " + TipeAdjustmentStr)
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "Out Adjust")
                    End If
                    If Me.IsDebet = True Then
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    End If
                ElseIf Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                    Me.Report1.SetParameterValue(Me.Report1.Parameter_Transaction_Type.ParameterFieldName, "In Sales Return")
                    If Me.IsDebet = True Then
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                    Else
                        Me.Report1.SetParameterValue(Me.Report1.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                    End If
                End If
                Me.crv1.ReportSource = Me.Report1
            End If

            Me.crv1.Show()
            Me.SetKontrol()
            Me.cboFormat.Enabled = True
            Me.btnLoad.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub LoadDataOutAdjust()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
        Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim Lanjut As Boolean = False

        Try
            'setting parameter laporan
            TBLStockInOut.Rows.Clear()
            DTAStockInOut.FillBy1(TBLStockInOut, TglSampai, TglDari, TipeStockInOut)
            For Each rowDGV In Me.vDGVItem.Rows
                Lanjut = False
                If Me.IsDebet = True Then
                    '   debet: adjustment expense account/gain account pada tipe adjustment tertentu
                    If vPrimaryTipeAdjustment = -1 Then
                        PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryAdjustmentAccount.Index).Value
                        If PrimaryAkunDGV = Me.vPrimaryAkun Then
                            Lanjut = True
                        End If
                    Else
                        'pasti lanjut, filter selanjutnya ditentukan oleh primary tipe adjustment di item detail
                        Me.vTblSettingItem.DefaultView.RowFilter = "[Primary Adjustment] = " + Me.vPrimaryTipeAdjustment.ToString
                        Me.rowSetting = Me.vTblSettingItem.DefaultView(0).Row
                        PrimaryAkunDGV = Me.rowSetting.PrimaryLoss
                        Lanjut = True
                    End If
                Else
                    '   debet: akun inventory
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value
                    If PrimaryAkunDGV = Me.vPrimaryAkun Then
                        Lanjut = True
                    End If
                End If

                If Lanjut = True Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString + _
                    " AND [primary adjustment type] = " + Me.vPrimaryTipeAdjustment.ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row

                        TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowStockInOut = TBLStockInOut.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowStockInOut.No_Transaksi
                                    Else
                                        .No_Dokumen_Asal = rowStockInOut.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataInSalesReturn()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR

        Try
            'setting parameter laporan
            TBLSR.Rows.Clear()
            DTASR.Fill(TBLSR, TglSampai, TglDari)
            For Each rowDGV In Me.vDGVItem.Rows
                If Me.IsDebet = True Then
                    '   debet: akun inventory
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value
                Else
                    '   kredit: stock take expense account-> cogs
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryCOGSAccount.Index).Value
                End If

                If PrimaryAkunDGV = Me.vPrimaryAkun Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowSR = TBLSR.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
                                    Else
                                        .No_Dokumen_Asal = rowSR.No_SR
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataInAdjust()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim Lanjut As Boolean = False

        Try
            'setting parameter laporan
            TBLStockInOut.Rows.Clear()
            DTAStockInOut.FillBy1(TBLStockInOut, TglSampai, TglDari, TipeStockInOut)
            For Each rowDGV In Me.vDGVItem.Rows
                Lanjut = False
                If Me.IsDebet = True Then
                    '   debet: akun inventory
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value
                    If PrimaryAkunDGV = Me.vPrimaryAkun Then
                        Lanjut = True
                    End If
                Else
                    '   kredit: adjustment expense account/gain account pada tipe adjustment tertentu
                    If vPrimaryTipeAdjustment = -1 Then
                        PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryAdjustmentAccount.Index).Value
                        If PrimaryAkunDGV = Me.vPrimaryAkun Then
                            Lanjut = True
                        End If
                    Else
                        'pasti lanjut, filter selanjutnya ditentukan oleh primary tipe adjustment di item detail
                        Me.vTblSettingItem.DefaultView.RowFilter = "[Primary Adjustment] = " + Me.vPrimaryTipeAdjustment.ToString
                        Me.rowSetting = Me.vTblSettingItem.DefaultView(0).Row
                        PrimaryAkunDGV = Me.rowSetting.PrimaryGain
                        Lanjut = True
                    End If
                End If

                If Lanjut = True Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString + _
                    " AND [primary adjustment type] = " + Me.vPrimaryTipeAdjustment.ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row

                        TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowStockInOut = TBLStockInOut.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowStockInOut.No_Transaksi
                                    Else
                                        .No_Dokumen_Asal = rowStockInOut.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataInStockTake()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust

        Try
            'setting parameter laporan
            TBLStockTake.Rows.Clear()
            DTAStockTake.Fill(TBLStockTake)
            For Each rowDGV In Me.vDGVItem.Rows
                If Me.IsDebet = True Then
                    '   debet: akun inventory
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value
                Else
                    '   kredit: stock take expense account
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryStockTakingAccount.Index).Value
                End If

                If PrimaryAkunDGV = Me.vPrimaryAkun Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowStockTake = TBLStockTake.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowStockTake.No_Transaksi
                                    Else
                                        .No_Dokumen_Asal = rowStockTake.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataOutStockTake()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust

        Try
            'setting parameter laporan
            TBLStockTake.Rows.Clear()
            DTAStockTake.Fill(TBLStockTake)
            For Each rowDGV In Me.vDGVItem.Rows
                If Me.IsDebet = True Then
                    '   debet: akun expense stock taking
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryStockTakingAccount.Index).Value
                Else
                    '   kredit: inventory akun
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value
                End If

                If PrimaryAkunDGV = Me.vPrimaryAkun Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowStockTake = TBLStockTake.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowStockTake.No_Transaksi
                                    Else
                                        .No_Dokumen_Asal = rowStockTake.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataInBundling()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling

        Try
            'setting parameter laporan
            TBLBundling.Rows.Clear()
            DTABundling.Fill(TBLBundling, TglSampai, TglDari)
            For Each rowDGV In Me.vDGVItem.Rows
                '   kredit: akun inventory
                PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value

                If PrimaryAkunDGV = Me.vPrimaryAkun Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowBundling = TBLBundling.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowBundling.No_Transaksi
                                    Else
                                        .No_Dokumen_Asal = rowBundling.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataOutSales()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales

        Try
            'setting parameter laporan
            TBLInvoice.Rows.Clear()
            DTAInvoice.Fill(TBLInvoice, TglSampai, TglDari)
            For Each rowDGV In Me.vDGVItem.Rows
                If Me.IsDebet = True Then
                    '   debet: akun cogs
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryCOGSAccount.Index).Value
                Else
                    '   kredit: akun inventory
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value
                End If
                If PrimaryAkunDGV = Me.vPrimaryAkun Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TBLInvoice.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowInvoice = TBLInvoice.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowInvoice.No_Transaksi + vbCrLf + rowInvoice.Customer_Name
                                    Else
                                        .No_Dokumen_Asal = rowInvoice.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataOutPurchaseReturn()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
        Dim TipeAkun As AccModule.ClsMasterAkun.EnumTipeAkun = AccModule.ClsMasterAkun.EnumTipeAkun.Asset
        Dim SelisihPositif As Boolean = False

        Try
            'setting parameter laporan
            TBLPR.Rows.Clear()
            DTAPR.Fill(TBLPR, TglSampai, Me.TipePR, TglDari)

            TipeAkun = Me.ObjAkun.GetTipeAkun(Me.vPrimaryAkun)
            If TipeAkun = AccModule.ClsMasterAkun.EnumTipeAkun.Asset Then
                If Me.IsDebet = True Then
                    '   debet: akun inventory: selisih positif -> purchase price - cost 
                    SelisihPositif = True
                Else
                    '   debet: akun inventory: selisih negatif -> purchase price - cost 
                    SelisihPositif = False
                End If
            Else
                If Me.IsDebet = True Then
                    '   debet: akun cogs: selisih negatif -> purchase price - cost 
                    SelisihPositif = False
                Else
                    '   debet: akun cogs: selisih positif -> purchase price - cost 
                    SelisihPositif = True
                End If
            End If

            For Each rowDGV In Me.vDGVItem.Rows
                If TipeAkun = AccModule.ClsMasterAkun.EnumTipeAkun.Asset Then
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value
                Else
                    PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryCOGSAccount.Index).Value
                End If

                If PrimaryAkunDGV = Me.vPrimaryAkun Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TBLPR.DefaultView.RowFilter = "[Primary Transaksi] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowPR = TBLPR.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowPR.No_Transaksi + vbCrLf + rowPR.Supplier_Name
                                    Else
                                        .No_Dokumen_Asal = rowPR.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value

                                If SelisihPositif = True Then
                                    .Nilai_Atas = rowItemDetail.return_price
                                    .Nilai_Bawah = rowItemDetail.amount
                                Else
                                    .Nilai_Atas = rowItemDetail.amount
                                    .Nilai_Bawah = rowItemDetail.return_price
                                End If
                                .Amount = .Nilai_Atas - .Nilai_Bawah
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            With rowCetak
                                If SelisihPositif = True Then
                                    .Nilai_Atas = rowItemDetail.return_price
                                    .Nilai_Bawah = rowItemDetail.amount
                                Else
                                    .Nilai_Atas = rowItemDetail.amount
                                    .Nilai_Bawah = rowItemDetail.return_price
                                End If
                                .Amount = .Nilai_Atas - .Nilai_Bawah
                            End With
                        End If
                    Next
                End If
            Next

            'show report
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.Report4.SetDataSource(Me.DS)
                Me.Report4.SetParameterValue(Me.Report4.Parameter_DefPrecAmount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                Me.Report4.SetParameterValue(Me.Report4.Parameter_Period.ParameterFieldName, Me.vPeriod)
                Me.Report4.SetParameterValue(Me.Report4.Parameter_Account.ParameterFieldName, Me.ObjAkun.GetAccountCode(Me.vPrimaryAkun) + _
                " -" + Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.vPrimaryAkun))
                Me.Report4.SetParameterValue(Me.Report4.Parameter_Transaction_Type.ParameterFieldName, "Out Purchase Return")
                If TipeAkun = AccModule.ClsMasterAkun.EnumTipeAkun.Asset Then
                    Me.Report4.SetParameterValue(Me.Report4.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                Else
                    Me.Report4.SetParameterValue(Me.Report4.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                End If
                If SelisihPositif = True Then
                    Me.Report4.SetParameterValue(Me.Report4.Parameter_Nilai_Atas.ParameterFieldName, "Return Price")
                    Me.Report4.SetParameterValue(Me.Report4.Parameter_Nilai_Bawah.ParameterFieldName, "Cost")
                Else
                    Me.Report4.SetParameterValue(Me.Report4.Parameter_Nilai_Atas.ParameterFieldName, "Cost")
                    Me.Report4.SetParameterValue(Me.Report4.Parameter_Nilai_Bawah.ParameterFieldName, "Return Price")
                End If
                Me.crv1.ReportSource = Me.Report4
            Else
                Me.Report3.SetDataSource(Me.DS)
                Me.Report3.SetParameterValue(Me.Report3.Parameter_DefPrecAmount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
                Me.Report3.SetParameterValue(Me.Report3.Parameter_Period.ParameterFieldName, Me.vPeriod)
                Me.Report3.SetParameterValue(Me.Report3.Parameter_Account.ParameterFieldName, Me.ObjAkun.GetAccountCode(Me.vPrimaryAkun) + _
                " -" + Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.vPrimaryAkun))
                Me.Report3.SetParameterValue(Me.Report3.Parameter_Transaction_Type.ParameterFieldName, "Out Purchase Return")
                If TipeAkun = AccModule.ClsMasterAkun.EnumTipeAkun.Asset Then
                    Me.Report3.SetParameterValue(Me.Report3.Parameter_TipeAkun.ParameterFieldName, "Inventory Account")
                Else
                    Me.Report3.SetParameterValue(Me.Report3.Parameter_TipeAkun.ParameterFieldName, "Expense Account")
                End If
                If SelisihPositif = True Then
                    Me.Report3.SetParameterValue(Me.Report3.Parameter_ParamAtas.ParameterFieldName, "Return Price")
                    Me.Report3.SetParameterValue(Me.Report3.Parameter_ParamBawah.ParameterFieldName, "Cost")
                Else
                    Me.Report3.SetParameterValue(Me.Report3.Parameter_ParamAtas.ParameterFieldName, "Cost")
                    Me.Report3.SetParameterValue(Me.Report3.Parameter_ParamBawah.ParameterFieldName, "Return Price")
                End If
                Me.crv1.ReportSource = Me.Report3
            End If

            Me.crv1.Show()
            Me.SetKontrol()
            Me.cboFormat.Enabled = True
            Me.btnLoad.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataOutBundling()
        Dim rowDGV As DataGridViewRow
        Dim PrimaryAkunDGV As Double = -1
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling

        Try
            'setting parameter laporan
            TBLBundling.Rows.Clear()
            DTABundling.Fill(TBLBundling, TglSampai, TglDari)
            For Each rowDGV In Me.vDGVItem.Rows
                '   kredit: akun inventory
                PrimaryAkunDGV = rowDGV.Cells(vFrmAcuan.ColItem_PrimaryInventoryAccount.Index).Value

                If PrimaryAkunDGV = Me.vPrimaryAkun Then
                    Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
                    " AND [primary code no] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString

                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        rowBundling = TBLBundling.DefaultView(0).Row

                        If Me.cboFormat.SelectedIndex = 0 Then
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString
                        Else
                            Me.TBLCetak.DefaultView.RowFilter = "[Primary Code No] = " + rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value.ToString + _
                            " AND [Primary Document Asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                        End If

                        If TBLCetak.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            Me.rowCetak = TBLCetak.NewRow
                            With rowCetak
                                .Amount = Me.rowItemDetail.amount
                                .Primary_Code_No = rowDGV.Cells(Me.vFrmAcuan.ColPrimaryCodeNo.Index).Value
                                If Me.cboFormat.SelectedIndex = 0 Then
                                    .Primary_Document_Asal = -1
                                    .No_Dokumen_Asal = ""
                                    .Tanggal = Me.vPeriod
                                    .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                Else
                                    .Primary_Document_Asal = Me.rowItemDetail.primary_transaksi_asal
                                    If Me.chkShowReferenceParty.Checked = True Then
                                        .No_Dokumen_Asal = rowBundling.No_Transaksi
                                    Else
                                        .No_Dokumen_Asal = rowBundling.No_Transaksi
                                    End If
                                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                                    If Me.cboProductDetail.SelectedIndex = 0 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value
                                    ElseIf Me.cboProductDetail.SelectedIndex = 1 Then
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                    Else
                                        .Code_No = rowDGV.Cells(vFrmAcuan.ColCodeNo.Index).Value.ToString + vbCrLf + _
                                        rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value.ToString
                                    End If
                                End If
                                .Product_Desc = rowDGV.Cells(vFrmAcuan.ColDescription.Index).Value
                                .Qty = Me.rowItemDetail.qty
                                .DefPrecQty = rowDGV.Cells(Me.vFrmAcuan.ColItem_Decimal.Index).Value
                                .Unit = rowDGV.Cells(Me.vFrmAcuan.ColUnits.Index).Value
                            End With
                            TBLCetak.Rows.Add(Me.rowCetak)
                        Else
                            'tambahkan qty dan amount saja
                            rowCetak = TBLCetak.DefaultView(0).Row
                            rowCetak.Qty += Me.rowItemDetail.qty
                            rowCetak.Amount += Me.rowItemDetail.amount
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        If Me.Aktif = True Then
            Me.LoadData()
        End If
    End Sub

    Private Sub SetKontrol()
        Try
            If Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                Me.chkUseSupplierInvoiceNumber.Enabled = True
            Else
                Me.chkUseSupplierInvoiceNumber.Enabled = False
            End If

            If Me.cboFormat.SelectedIndex = 0 Then
                Me.cboProductDetail.Enabled = False
                Me.chkShowReferenceParty.Enabled = False
            Else
                If Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Or _
                Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Or _
                Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Or _
                Me.vTipeTrasaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                    Me.chkShowReferenceParty.Enabled = True
                Else
                    Me.chkShowReferenceParty.Enabled = False
                End If
                Me.cboProductDetail.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.Aktif = True Then
                Me.SetKontrol()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class