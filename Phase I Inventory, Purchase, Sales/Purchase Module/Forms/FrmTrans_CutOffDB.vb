Option Explicit On
Option Strict Off

Imports MySql.Data.MySqlClient
Imports MySql.Data

Public Class FrmTrans_CutOffDB
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow

    Dim CutOffVersion As String = ""
    Dim Temporary_DB As String = Core_Module.GetDBName & "_Temporary"
    Dim SumberConstring As String = Core_Module.GetConstring
    Dim TargetConstring As String = SumberConstring.Replace(Core_Module.GetDBName, Temporary_DB)

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCutOff As New ClsTrans_CutOffDB
    Dim ObjCutOff_Core As New Core_Module.ClsTrans_CutOffDB_Core
    Dim ObjJurnal As New AccModule.ClsTransaksiJournal
    Dim objSetting As New Core_Module.ClsTransaksiSetting
    Dim objLocation As New Core_Module.clsMasterStorageArea

    Dim FrmBackup As New Core_Module.FrmBackUpData
    Dim FrmRestore As New Core_Module.FrmBackUpData2

    Dim UseActiveLicense As Boolean = False
    Dim ActiveLicense As Integer = 0
    Dim ApakahControlUser As Boolean = True
    Dim StatusOnProses As String = "On Progress"
    Dim StatusOK As String = "OK"
    Dim StartDateCutOff As Date = Now
    Dim EndDateCutOff As Date = Now
    Dim counterHeader, counterItem, totalItem As Integer
    Dim RowIdx As Integer = -1
    Dim StatusCutOff As Boolean = False

    Dim TblListTransfer As New DseTrans_CutOffDB.ListTransferDataTable
    Dim RowListTransfer As DseTrans_CutOffDB.ListTransferRow
    Dim DTA_Beginning_Balance As New DseTrans_CutOffDBTableAdapters.tbl_beginning_balanceTableAdapter
    Dim TBL_Beginning_Balance As New DseTrans_CutOffDB.tbl_beginning_balanceDataTable
    Dim DTA_Trial_Balance As New DseTrans_CutOffDBTableAdapters.Trial_BalanceTableAdapter
    Dim TBL_Trial_Balance As New DseTrans_CutOffDB.Trial_BalanceDataTable

    Dim DtaCurr As New Core_Module.dseMasterCurrencyTableAdapters.tblmaster_currency1TableAdapter
    Dim TblCurr As New Core_Module.dseMasterCurrency.tblmaster_currency1DataTable
    Dim DtaDec As New DseTrans_CutOffDBTableAdapters.tblmaster_productdecimalTableAdapter
    Dim TblDec As New DseTrans_CutOffDB.tblmaster_productdecimalDataTable

    Dim TblDo_Out As New DseTrans_CutOffDB.tbltrans_doDataTable

    Dim TblPrimary_SO, TblPrimary_DO, TblPrimary_SOS, TblPrimary_SISR, TblPrimary_SIP, TblPrimary_GRN, TblPrimary_SR, TblPrimary_Inv, TblPrimary_CD As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
    Dim TblPrimary_PR, TblPrimary_AOP, TblPrimary_PO, TblPrimary_PL As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
    Dim TblPrimary_Taking, TblPrimary_InAdj, TblPrimary_OutAdj, TblPrimary_Bund As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable

    Public Enum EnumSiklusPurchase As Integer
        GetPOBData
        TransferPurchaseData
    End Enum

    Public Enum EnumSiklusSales As Integer
        GetSOBData
        TransferSalesData
    End Enum

    Dim StatusShowLabel As Integer = EnumLabel.LabelUtama
    Public Enum EnumLabel As Integer
        LabelUtama
        LabelPOB
        LabelSOB
    End Enum

    Enum EnumProdQtyNStockCard As Integer
        SetSumber
        SetTarget
    End Enum

    Private Sub SetDefaultConnectionString()
        Try
            AccModule.SetConstring_AccModule(Me.SumberConstring)
            Core_Module.SetConstring(Me.SumberConstring)
            'CustomSMESS.SetConstring_CustomModule(Me.SumberConstring)
            E_mail_Module.SetConstring_EmailModule(Me.SumberConstring)
            Inventory_Module.SetConstring_InventoryModule(Me.SumberConstring)
            Purchase_Module.SetConstring_Purchase_Module(Me.SumberConstring)
            Sales_Module.SetConstring_SalesModule(Me.SumberConstring)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTargetConnectionString()
        Try
            AccModule.SetConstring_AccModule(Me.TargetConstring)
            Core_Module.SetConstring(Me.TargetConstring)
            'CustomSMESS.SetConstring_CustomModule(Me.TargetConstring)
            E_mail_Module.SetConstring_EmailModule(Me.TargetConstring)
            Inventory_Module.SetConstring_InventoryModule(Me.TargetConstring)
            Purchase_Module.SetConstring_Purchase_Module(Me.TargetConstring)
            Sales_Module.SetConstring_SalesModule(Me.TargetConstring)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_CutOffDB_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If ApakahControlUser = True Then
                Me.ObjCutOff.ControlUserLogin(Me.UseActiveLicense, Me.ActiveLicense, True)
            End If

            If Me.btnStartCancel.Text = "Cancel" Then
                If Me.CancelProsesCutOff() = False Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If e.Cancel = False Then
                Me.SetDefaultConnectionString()
            End If
        End Try
    End Sub

    Private Sub FrmTrans_CutOffDB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Me.TxtDB.Visible = False
            'Me.LinkLabel1.Visible = False
            Me.LinkLabel1.Text = Me.status_status & " | " & Core_Module.GetDBName
            Me.TxtDB.Text = Core_Module.GetDBName

            Me.Label1.Text = ""
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Visible = False
            Me.Bgw1.WorkerSupportsCancellation = True

            Me.StartDateCutOff = CDate("1/1/1950")
            Me.EndDateCutOff = Me.ObjCD.GetLastDateOfTheYear(Me.dtpCutOffPeriod.Value)
            Me.dgvProgress.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'pastikan hanya ada 1 user yang online
            If Me.ObjCutOff.CekUserLogin() = False Then
                MsgBox("Can't start cut off database, make sure only you users who login smess", MsgBoxStyle.Information, "Information")
                Me.btnStartCancel.Enabled = False
                Me.ApakahControlUser = False
                Exit Sub
            End If

            '2. set ke maintenance mode, sehingga tidak ada user lain yang bisa masuk ketika cut off db
            Me.ObjCutOff.ControlUserLogin(Me.UseActiveLicense, Me.ActiveLicense)

            Me.LoadCutOff_AllModule()

            'MsgBox("Harus bisa reversible", MsgBoxStyle.Information, "Information")

            Me.DtaCurr.Fill(Me.TblCurr)
            Me.DtaDec.Fill(Me.TblDec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles Bgw1.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.btnStartCancel.Text = "Cancel"
            Me.dtpCutOffPeriod.Enabled = False
            Me.ProgressBar1.Visible = True

            If Me.CutOff_AllModule() = True Then
                Me.StatusCutOff = True
            Else
                Me.StatusCutOff = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: Bgw1_DoWork")
        End Try
    End Sub

    Private Sub Bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Bgw1.RunWorkerCompleted
        Try
            If Me.StatusCutOff = True Then
                Me.StockCardVSProductQty(EnumProdQtyNStockCard.SetTarget)
                MsgBox("Cut Off Successed", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Cut Off Canceled", MsgBoxStyle.Information, "Information")
            End If

            Me.btnStartCancel.Text = "Start"
            Me.dtpCutOffPeriod.Enabled = True
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Visible = False
            Me.Label1.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: Bgw1_RunWorkerCompleted")
        End Try
    End Sub

    Private Sub btnStartCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartCancel.Click
        Try
            If Me.btnStartCancel.Text = "Cancel" Then
                Me.CancelProsesCutOff()
            Else
                If Me.ApakahDataBenar = True Then
                    '1. back up data
                    Me.BackAndRestoreDB()
                    FrmTrans_CutOffDB.CheckForIllegalCrossThreadCalls = False
                    Me.Bgw1.RunWorkerAsync()
                    'Me.CutOff_AllModule()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpCutOffPeriod_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpCutOffPeriod.ValueChanged
        Try
            Me.EndDateCutOff = Me.ObjCD.GetLastDateOfTheYear(Me.dtpCutOffPeriod.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RowAdd()
        Dim number_ As New List(Of Integer)
        Dim rowDGV As DataGridViewRow

        Try
            Me.TblListTransfer.DefaultView.RowFilter = Nothing
            Me.TblListTransfer.DefaultView.Sort = "[No]"
            For Each dv As DataRowView In Me.TblListTransfer.DefaultView
                Me.RowListTransfer = dv.Row
                If number_.Contains(Me.RowListTransfer.No) = False Then
                    number_.Add(Me.RowListTransfer.No)
                    Me.dgvProgress.Rows.Add()
                    rowDGV = Me.dgvProgress.Rows(Me.dgvProgress.RowCount - 1)
                    rowDGV.Cells(Me.ColDesc.Index).Value = Me.RowListTransfer.Description
                    rowDGV.Cells(Me.ColModule.Index).Value = Me.RowListTransfer._Module
                    rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColStatus.Index).Value = "On Queue"
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: RowAdd")
        End Try
    End Sub

    Private Sub RowUpdate(ByVal vDesc As String, ByVal vStatus As String)
        Dim rowDGV As DataGridViewRow
        Dim rowIdx As Integer = -1

        Try
            Me.TblListTransfer.DefaultView.RowFilter = "[Description] = '" & vDesc.ToString & "'"
            If Me.TblListTransfer.DefaultView.Count > 0 Then
                Me.RowListTransfer = Me.TblListTransfer.DefaultView(0).Row
                rowIdx = Me.RowListTransfer.No
            Else
                Exit Sub
            End If

            If rowIdx < Me.dgvProgress.RowCount - 1 Then
                rowDGV = Me.dgvProgress.Rows(rowIdx)
                rowDGV.Cells(Me.ColStatus.Index).Value = vStatus
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: RowUpdate")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim Dta As New DseTrans_CutOffDBTableAdapters.tbltrans_accounting_closingTableAdapter
        Dim Tbl As New DseTrans_CutOffDB.tbltrans_accounting_closingDataTable
        Dim row As DseTrans_CutOffDB.tbltrans_accounting_closingRow

        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ClosingMethod_ As String = ""

        Try
            '1. pastikan cut off period ada setelah convertion date 
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl)
            If Tbl.Count > 0 Then
                row = Tbl.Rows(0)
                If Me.EndDateCutOff <= row.Convertion_Date Then
                    MsgBox("Cut off period must be bigger than convertion date", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox("Cut off period must be bigger than convertion dat", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2. cocokan versi db dengan versi cut off
            'sementara di pending dulu

            '---Tambahan Pastika Product Qty dengan Stok Card di Database sumber sama
            If Me.StockCardVSProductQty(EnumProdQtyNStockCard.SetSumber) = False Then
                MsgBox("Product qty are different from stock card", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. back up data terlebih dahulu
            Me.FrmDia.Title = "Please specify where to save back up file"
            Me.FrmDia.Filter = "*.smess|smess back up"
            Me.FrmDia.FilterIndex = 0
            Me.FrmDia.DefaultExt = ".smess"

            Me.FrmDia.OverwritePrompt = True
            If Me.FrmDia.ShowDialog <> DialogResult.OK Then
                Return False
            End If
            If Me.FrmDia.FileName = "" Then
                Return False
            End If

            Me.FrmBackup.strConn = Core_Module.GetConstring
            Me.FrmBackup.vFileName = FrmDia.FileName
            Me.FrmBackup.DBServer = Core_Module.GetDBServer
            Me.FrmBackup.Database = Core_Module.GetDBName
            Me.FrmBackup.DBUser = Core_Module.GetDBUserID
            Me.FrmBackup.DBPass = Core_Module.GetDBPassword
            Me.FrmBackup.DBPort = Core_Module.GetDBPort
            Me.FrmBackup.vFileName2 = ""
            Me.FrmBackup.vBackUpDesc = "SMESS Auto Back Up"
            Me.FrmBackup.vCompress = True
            Me.FrmBackup.vShutdown = False
            Me.FrmBackup.vEncrypt = True
            Me.FrmBackup.vEncryptPass = Core_Module.GetEncryptPass

            '4. jika LIFO, keluarkan logic error
            ClosingMethod_ = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
            If ClosingMethod_ = "2" Then    '\\ nilai value 0 = average, 1 = fifo, 2 = lifo.
                MsgBox("Logic error, contact your provider when you see this message", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5. pastikan sudah end of year jurnal

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: ApakahDataBenar")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            ObjSetting = Nothing
        End Try
    End Function

    Public Function BackAndRestoreDB() As Boolean
        Dim comm As New MySql.Data.MySqlClient.MySqlCommand
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection
        Dim FileName_ As String = ""

        Try
            Me.RowUpdate("Back up database", Me.StatusOnProses)

            'Backup database
            FileName_ = FrmDia.FileName.Substring(0, FrmDia.FileName.Length - 6) + " - " + _
            Format(Now, "dd MMM yyyy - HH_mm") + FrmDia.FileName.Substring(FrmDia.FileName.Length - 6)
            Me.FrmBackup.ShowDialog()

            'Create new database
            Conn.ConnectionString = Core_Module.GetConstring
            Conn.Open()
            comm.Connection = Conn
            comm.CommandText = "DROP DATABASE IF EXISTS " & Me.Temporary_DB
            comm.ExecuteNonQuery()
            comm.CommandText = "CREATE DATABASE " & Me.Temporary_DB
            comm.ExecuteNonQuery()
            Conn.Close()

            'Restore database
            Me.FrmRestore.vFileName = FileName_
            Me.FrmRestore.vConStr = Me.TargetConstring
            Me.FrmRestore.vCompress = True
            Me.FrmRestore.vUpdateData = False
            Me.FrmRestore.vUpdateStructure = True
            Me.FrmRestore.ShowDialog()

            Me.RowUpdate("Back up database", Me.StatusOK)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: BackAndRestoreDB")
            Return False
        Finally
            Conn = Nothing
            comm = Nothing
        End Try
    End Function

    Private Function CancelProsesCutOff() As Boolean
        Dim comm As New MySql.Data.MySqlClient.MySqlCommand
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection
        Dim NewDBName As String = Core_Module.GetDBName & Me.Temporary_DB
        Dim hasil, hasil2 As MsgBoxResult

        Try
            hasil = MsgBox("Cancel proccess cut of database?", MsgBoxStyle.OkCancel, "Confir")
            If hasil = MsgBoxResult.Ok Then
                If Me.Bgw1.IsBusy = True Then
                    Me.Bgw1.CancelAsync()
                End If

                hasil2 = MsgBox("Drop Database Cut Off?", MsgBoxStyle.OkCancel, "Confir")
                If hasil2 = MsgBoxResult.Ok Then
                    Conn.ConnectionString = Core_Module.GetConstring
                    Conn.Open()
                    comm.Connection = Conn
                    comm.CommandText = "DROP DATABASE IF EXISTS " & Me.Temporary_DB
                    comm.ExecuteNonQuery()
                    Conn.Close()
                End If

                Me.Label1.Text = ""
                Me.ProgressBar1.Value = 0
                Me.btnStartCancel.Text = "Start"
                Me.dtpCutOffPeriod.Enabled = True

                For Each rowD As DataGridViewRow In Me.dgvProgress.Rows
                    rowD.Cells(Me.ColStatus.Index).Value = "On Queue"
                Next

                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CancelProsesCutOff")
            Return False
        Finally
            Conn = Nothing
            comm = Nothing
        End Try
    End Function

    'LOAD MODULE
    Private Sub LoadCutOff_AllModule()
        Try
            Me.LoadCutOff_CoreModule()
            Me.LoadCutOff_PurchaseModule()
            Me.LoadCutOff_SalesModule()
            Me.LoadCutOff_InventoryModule()
            Me.LoadCutOff_HRModule()
            Me.LoadCutOff_AccountingModule()

            'Load Monthly Closing
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Monthly Closing", "Transfer Monthly Closing Data", "tbltrans_purchase_closing")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Monthly Closing", "Transfer Monthly Closing Data", "tbltrans_purchase_closing_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Monthly Closing", "Transfer Monthly Closing Data", "tbltrans_purchase_closing_item_akun")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Monthly Closing", "Transfer Monthly Closing Data", "tbltrans_purchase_closing_item_detail")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Monthly Closing", "Transfer Monthly Closing Data", "tbltrans_purchase_closing_item_inadjustment")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Monthly Closing", "Transfer Monthly Closing Data", "tbltrans_purchase_closing_item_tipe_adjustment")

            Me.LoadCutOff_CustomModule()

            Me.RowAdd()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CancelProsesCutOff")
        End Try
    End Sub

    Private Sub LoadCutOff_CoreModule()
        Try
            Me.RowIdx += 1 'back up
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Back up database", "")

            'master data
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Code Data", "tblmaster_usergroup")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Code Data", "tblmaster_usergroup_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Code Data", "tblmaster_usergroup_item_turunan")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Code Data", "tbluser")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Code Data", "tblmaster_user_temppass")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Code Data", "tblmaster_user_temppass_item")
            Me.RowIdx += 1 '   - access management
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Access Management Data", "tblmasterhakakses")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Access Management Data", "tblmasterhakkhusus")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Access Management Data", "tblantarauserhakakses")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer User Access Management Data", "tblantarauserhakkhusus")
            Me.RowIdx += 1 '   - db autoback up
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Auto Back Up Data", "tblmaster_backup_db")
            Me.RowIdx += 1 '   - coa + update balance value
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Chart Of Account and Begining Balance Data", "tblmaster_akun")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Chart Of Account and Begining Balance Data", "tblmaster_asset_group")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Chart Of Account and Begining Balance Data", "tblmaster_asset")
            Me.RowIdx += 1 '   - budget
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Budget Data", "tblmaster_budget")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Budget Data", "tblmaster_budget_item")
            Me.RowIdx += 1 '   - customer group
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Group Data", "tblmaster_customer_group")
            Me.RowIdx += 1 '   - customer type
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Type Data", "tblmaster_customer_type")
            Me.RowIdx += 1 '   - customer sales area
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Sales Area Data", "tblmaster_customer_sales_area")
            Me.RowIdx += 1 '   - customer market segmentation
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Segmentation Data", "tblmaster_customer_segmentation")
            Me.RowIdx += 1 '   - sales person group
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Person Group Data", "tblmaster_salesperson_group")
            Me.RowIdx += 1 '   - sales person
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Person Data", "tblmaster_salesperson")
            Me.RowIdx += 1 '   - sales person budget
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Person Budget", "tbltrans_salesbudget")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Person Budget", "tbltrans_salesbudget_item")
            Me.RowIdx += 1 '   - customer
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Data", "tblmaster_customer")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Data", "tblmaster_customer_item_address")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Data", "tblmaster_customer_item_bank")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Customer Data", "tblmaster_customer_itemcp")
            Me.RowIdx += 1 '   - supplier
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Supplier", "tblmaster_supplier")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Supplier", "tblmaster_supplier_item_address")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Supplier", "tblmaster_supplier_item_bank")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Sales Supplier", "tblmaster_supplier_itemcp")
            Me.RowIdx += 1 '   - currency
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Currency Data", "tblmaster_currency")
            Me.RowIdx += 1 '   - Department
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Department Data", "tblmaster_departemen")
            Me.RowIdx += 1 '   - unit of measurement
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Unit Of Measurement Data", "tblmaster_master_unit")
            Me.RowIdx += 1 '   - warehouse
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Warehouse Data", "tblmaster_warehouse")
            Me.RowIdx += 1 '   - location
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Location Data", "tblmaster_storage_area")
            Me.RowIdx += 1 '   - down time
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Downtime Data", "tblmaster_downtime")
            Me.RowIdx += 1 '   - machine group
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Machine Group Data", "tblmaster_machinegroup")
            Me.RowIdx += 1 '   - process
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Process Data", "tblmaster_proses")
            Me.RowIdx += 1 '   - reject
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Reject Data", "tblmaster_reject")
            Me.RowIdx += 1 '   - machine
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Machine Data", "tblmaster_machinedata")
            Me.RowIdx += 1 '   - work unit
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Work Unit Data", "tblmaster_workunit")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Work Unit Data", "tblmaster_workunit_item_employee")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Work Unit Data", "tblmaster_workunit_item_machine")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Work Unit Data", "tblmaster_workunit_item_setuptime")
            Me.RowIdx += 1 '   - product group
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Group, Brand, Sub Brand Data", "tblmaster_product_group")
            '   - brand
            'Me.TblListTransfer.Rows.Add(27, "Master data", "Transfer Product Brand Data", "")
            '   - sub brand
            'Me.TblListTransfer.Rows.Add(28, "Master data", "Transfer Product Sub Brand Data", "")
            Me.RowIdx += 1 '- category
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Category Data", "tblmaster_category")
            Me.RowIdx += 1 '   - product custom value
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Custom Value Data", "tblmaster_kompetensi_product")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Custom Value Data", "tblmaster_kompetensi_product_item")
            Me.RowIdx += 1 '   - product
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_product")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_productqty")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_product_item_bom")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_product_item_multiunit")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_product_item_workunit")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_product_item_kompetensi")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_product_item_attachment")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Data", "tblmaster_product_pricing")
            Me.RowIdx += 1 '   - price list
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Price List Data", "tblmaster_product_pricelist")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Price List Data", "tblmaster_product_pricelist_group")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Price List Data", "tblmaster_product_pricelist_customer")
            Me.RowIdx += 1 '   - product account group
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Product Account Group Data", "tblmaster_product_account_group")
            Me.RowIdx += 1 '  - product material planning group
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_group")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_product")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_product_group")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_inv_exc")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_itemsupplier")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_adjustment")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_bundling")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_invoice")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Material Planning Group Data", "tbltrans_material_planning_group_item_sr")

            'master HR
            Me.RowIdx += 1 '   - cuti
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Leave Data", "tblmaster_cuti")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Leave Data", "tblmaster_cuti_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Leave Data", "tblmaster_cuti_item_desc")
            Me.RowIdx += 1 '   - job position
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Job Position Data", "tblmaster_jobposition")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Job Position Data", "tblmaster_jobposition_item_kompetensi")
            Me.RowIdx += 1 '   - karyawan
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Employee Data", "tblmaster_karyawan")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Employee Data", "tblmaster_karyawan_custom_field")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Employee Data", "tblmaster_karyawan_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Employee Data", "tblmaster_karyawan_item_kompetensi")
            Me.RowIdx += 1 '   - kompetensi
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Competence Data", "tblmaster_kompetensi")
            Me.RowIdx += 1 '   - payroll access
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Payroll Access Data", "tblmaster_payroll_access")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Payroll Access Data", "tblmaster_payroll_access_item_salaryformula")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Payroll Access Data", "tblmaster_payroll_access_item_user")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Payroll Access Data", "tblmaster_priority_level")
            Me.RowIdx += 1 '   - workshift
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Workshift Data", "tblmaster_workshift")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Workshift Data", "tblmaster_workshift_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Workshift Data", "tblmaster_workshift_itemallocate")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Workshift Data", "tblmaster_workshift_itemholiday")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Workshift Data", "tblmaster_workshift_itemparameters")
            Me.RowIdx += 1 '   - slip desc
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Slip Description Data", "tblmaster_salary_slip_description")
            Me.RowIdx += 1 '   - salary formula
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_component_detailemployee")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_otherdeduction")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_otherdeduction_constraint")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_otherdeduction_employee")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_otherdeduction_variable")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_parameter")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_salarycomponent")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_salarycomponent_constraint")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Salary Formula Data", "tblmaster_salary_formula_salarycomponent_variable")

            'setting:
            Me.RowIdx += 1 '   - global setting + tbltrans_setting + tbltrans_setting_user
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbllistedsoftware")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_adjustment")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_auxilary")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_direct_labor")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_divisi")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_factory_overhead")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_fg")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_rm")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_trading")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_cogs_wip")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_leave")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_setting_update")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblappsetting")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_ledger_acces")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_ledger_access_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_product_group")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_inventorymovement")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_inventorymovement_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_stockbundling")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_stockbundling_item_filter_process")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_stockbundling_item_filter_product_group")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_stockbundling_item_filter_workshift")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_stockbundling_item_filter_workunit")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_report_stockbundling_item_sort")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_transf_aop")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_transf_aopr")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_transf_inv")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_transf_invdp")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_transf_sr")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_user")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_setting_warehouse_exclusion")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_accounting_closing")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_kode_negara")

            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_fingerprint_server")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_fingerprint_server_machines")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_cuti_pengajuan_setting")


            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_numbering")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_numbering_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_smess")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbl_log_user")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_version_control")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_version_operating")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_dbversion")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_emailsettings")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_emailsettings_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_emailsettings_log")
            '---
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_ptkp")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_tarif_pph")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_tarif_pph_final_pesangon")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_tipe")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tblmaster_tipe_otorisasi")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_salesinvoice_taxrate")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_salesinvoice_taxrate_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Global Setting Data", "tbltrans_salesinvoice_tax_invoice_setting")
            Me.RowIdx += 1 '   - company data setting
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Company Data Settings", "tbltrans_company_data")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Master data", "Transfer Company Data Settings", "tbltrans_company_data_item")
            ''   - default exchange rate
            'Me.TblListTransfer.Rows.Add(35, "Master data", "Transfer Default Exchange Rate Data", "")

            Me.RowIdx += 1 'begining balance aop
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Begining Balance", "Transfer Purchase & Purchase Return Begining Balance Data", "")
            'Me.RowIdx += 1 'begining balance pr
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Begining Balance", "Transfer Purchase Return Begining Balance Data", "")
            Me.RowIdx += 1 'begining balance inv
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Begining Balance", "Transfer Sales & Sales Return Begining Balance Data", "")
            'Me.RowIdx += 1 'begining balance sr
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Begining Balance", "Transfer Sales Return Begining Balance Data", "")
            Me.RowIdx += 1 'begining balance qty -> ini hati2 ketika avg, fifo, atau lifo berbeda prosedur input datanya
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Begining Balance", "Transfer Product Qty Begining Balance Data", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCutOff_InventoryModule()
        Try
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Stock In Adjust Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Stock Out Adjust Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Stock Bundling Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Stock Taking Data", "")
            'Me.RowIdx += 1
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Inventory Check Tag Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Location Transfer Data", "")

            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_item_adv_exclusion")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_item_location_exclusion")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_item_qty")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_detailotorisasipo")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_itemdo")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_itemgit")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_itempob")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_itemsob")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Material Planning Data", "tbltrans_material_planning_itempurch_req")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Shipping Forecast Data", "tbltrans_shipping_forecast")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Shipping Forecast Data", "tbltrans_shipping_forecast_item_pob")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Shipping Forecast Data", "tbltrans_shipping_forecast_itemdelivery")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Inventory Module", "Transfer Shipping Forecast Data", "tbltrans_shipping_forecast_itemstock")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCutOff_PurchaseModule()

        Try
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Purchase Order Balance Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Purchase Order Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Packing List Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Stock In Purchase Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Purchase Return Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Stock Out Purchase Return Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Quick Purchase Return Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Downpayment Approval Of Purchase Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Approval Of Purchase Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer Quick AOP Data")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Transfer PO Cancellation Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Purchase Module", "Set Purchase Order Balance", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCutOff_SalesModule()
        Dim number_ As New List(Of Integer)
        Try
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Quotation Data", "tbltrans_quotation")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Quotation Data", "tbltrans_quotation_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Quotation Data", "tbltrans_quotation_itemdelivery")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Sales Order Balance Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Sales Order Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Delivery Order Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Stock Out Sales Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Stock In Sales Return Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Sales Return Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Quick Sales Return Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Downpayment Invoice Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Sales Invoice Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer Quick Invoice Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Transfer SO Cancellation Data", "")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Sales Module", "Set Sales Order Balance", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCutOff_HRModule()
        Try
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Leave Application Data", "tbltrans_cuti_pengajuan")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Leave Application Data", "tbltrans_cuti_pengajuan_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Leave Application  Data", "tbltrans_cuti_pengajuan_setting")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Attendance Data", "tbltrans_attendance")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Attendance Data", "tbltrans_attendance_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Attendance Data", "tbltrans_attendancelog")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Loan Application Data", "tbltrans_loan")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Loan Application Data", "tbltrans_loan_item")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Loan Payment Data", "tbltrans_loan_payment")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Loan Payment Data", "tbltrans_loan_payment_tem")
            'Me.RowIdx += 1
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Other Payment Data")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_itemtax")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_item_iteration")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_item_other")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_item_otherdeduction")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_employee_d_igi")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_hitungpph21")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_laporanincometax_igi")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_override_amount")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Data", "tbltrans_salary_override_amount_item")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Payment Data", "tbltrans_payment")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Payment Data", "tbltrans_payment_item_employee")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Payment Data", "tbltrans_payment_item_employee_d")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Payment Data", "tbltrans_payment_item_jurnal")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Salary Payment Data", "tbltrans_payment_setting_cashortransfer")
            'Me.RowIdx += 1
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "HR Module", "Transfer Collection Letter Data")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCutOff_AccountingModule()
        Try
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer General Journal & Cash Bank Data", "tbltrans_jurnal_quick_add")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer General Journal & Cash Bank Data", "tbltrans_jurnal_quick_add_item")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer General Journal & Cash Bank Data", "tbltrans_jurnal")
            'Me.RowIdx += 1
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer Cash And Bank Data")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer Credit & Debit Note Data")
            'Me.RowIdx += 1
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer Debit Note Data")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer Purchase Payment Journal Data")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer Sales Receivement Journal Data")
            Me.RowIdx += 1
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Accounting Module", "Transfer Collection Letter Data")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCutOff_CustomModule()
        Try
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Driver Data", "tbltrans_custom_driver")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Expedition Data", "tbltrans_custom_expedition")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Delivery Area Data", "tbltrans_custom_delivery_area")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Collection Order Data", "tbltrans_custom_collection_order")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Collection Order Data", "tbltrans_custom_collection_order_item")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Collection Log Data", "tbltrans_custom_collection_log")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Collection Log Data", "tbltrans_custom_collection_log_item")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Delivery Operation Data", "tbltrans_custom_delivery_operation")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Delivery Operation Data", "tbltrans_custom_delivery_operation_item")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Delivery Operation Data", "tbltrans_custom_delivery_operation_expedition")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Sales Receipt Data", "tbltrans_custom_sales_receipt")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Sales Receipt Data", "tbltrans_custom_sales_receipt_item")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Surat Tagihan Data", "tbltrans_custom_surattagihan")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Surat Tagihan Data", "tbltrans_custom_surattagihan_item")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Loan Appllication Data", "tbltrans_custom_loanapplication")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Unit Cost Data", "tbltrans_custom_unitcost")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Unit Cost Data", "tbltrans_custom_unitcost_item")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Unit Cost Data", "tbltrans_custom_unitcost_data")
            'Me.RowIdx += 1 '   - user code
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Kwitansi Data", "tbltrans_custom_kwitansi")
            Me.RowIdx += 1 '   - user code
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Setting Data", "tbltrans_customtrans")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Setting Data", "tbltrans_customreport")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Setting Data", "tbltrans_custom_identifier")
            Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer Custom Setting Data", "tbltrans_custom_update_bundling")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer User Data", "tbltrans_custom_so_print_option")
            'Me.TblListTransfer.Rows.Add(Me.RowIdx, "Custom Module", "Transfer User Data", "tbltrans_custom_stocktakingprep")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'CUTOFF MODULE
    Private Function CutOff_AllModule() As Boolean
        'Me.GetTblPrimaryUpdate()   

        Try
            'karena saat restructure auto increment primary tidak mulai dari 1
            'maka terlebih dahulu di truncate
            Me.TruncateAllTable()

            'urutan proses
            '2. start cut off -> ketika cut off batasan stock minus, pob minus, dan sob minus tidak berlaku
            If Me.CutOff_CoreModule() = False Then
                Return False
            End If
            If Me.CutOff_PurchaseModule() = False Then
                Return False
            End If
            If Me.CutOff_SalesModule() = False Then
                Return False
            End If
            If Me.CutOff_InventoryModule() = False Then
                Return False
            End If
            If Me.CutOff_HRModule() = False Then
                Return False
            End If
            If Me.CutOff_AccountingModule() = False Then
                Return False
            End If

            'tambahan kerjaan
            '1. closing inventory
            If Me.TransferMonthlyClosing() = False Then
                Return False
            End If
            ''2. revaluasi
            ''3. journal locking

            ''cut off custom
            If Me.CutOff_CustomModule() = False Then
                Return False
            End If

            'review ulang semuanya
            '1. semua laporan database target dan sumber harus sama
            '3. 

            'identified problems:
            '1. apa yang terjadi ketika ada barang yang minus (per lot/location, bukan secara keseluruhan) pada tanggal cut off
            '2. 

            'result:
            '1. script database asal
            '2. script database hasil
            '3. script database asal - transaksi stlh cut off -> ini digunakan untuk ngereverse proses

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CutOff_CoreModule() As Boolean
        Try
            'Transfer semua data master
            If Me.TransferMaster() = False Then
                Return False
            End If
            'Transfer AP balance
            If Me.TransferAPBalance() = False Then
                Return False
            End If
            'Transfer AR balance
            If Me.TransferARBalance() = False Then
                Return False
            End If
            'Transfer begining balance
            If Me.TransferBeginingBalance() = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CutOff_CoreModule")
        End Try
    End Function

    Private Function CutOff_InventoryModule() As Boolean
        Try
            'in adjust
            If Me.TransferStockInAdjust() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_inadj", , , Me.TblPrimary_InAdj, , False)
            'out adjust
            If Me.TransferStockOutAdjust() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_ouadj", , , Me.TblPrimary_OutAdj, , False)
            'bundling
            If Me.TransferStockBundling() = False Then
                Return False
            End If
            'stock taking
            If Me.TransferStockTaking() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_taking", , , Me.TblPrimary_Taking, , False)
            'location transfer
            If Me.TransferStockTransfer() = False Then
                Return False
            End If
            'Material Planning & Shipping Forecast
            If Me.TransferMP_SF() = False Then
                Return False
            End If
            'setelah semua proses transfer yang mempengaruhi product qty
            'maka hapus tblmaster_productqty dan transfer ulang datanya dari db sumber.
            Me.TrancuteNTransfer("tblmaster_productqty")

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CutOff_InventoryModule")
            Return False
        End Try
    End Function

    Private Function CutOff_PurchaseModule() As Boolean
        Try
            'purchase order /get pob ada di dalam fungsi transfer po
            If Me.TransferPO() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_po", , , Me.TblPrimary_PO, , False)
            'packing list
            If Me.TransferPackingList() = False Then
                Return False
            End If
            'stock in purchase
            If Me.TransferStockInPurch() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_sip", , , Me.TblPrimary_SIP, , False)
            'Me.TransferMaster_UsingScript("a_grn", , , Me.TblPrimary_GRN, , False)
            'purchase return 
            If Me.TransferPurchReturn() = False Then
                Return False
            End If
            'stock out purchase return
            If Me.TransferStockOutPurchReturn() = False Then
                Return False
            End If
            'quick pr
            If Me.TransferQuickPurchReturn() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_pr", , , Me.TblPrimary_PR, , False)
            'aop /transfer aop dp ada di dalam fungsi transfer aop
            If Me.TransferAOP() = False Then
                Return False
            End If
            'quick aop
            If Me.TransferQuickAOP() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_aop", , , Me.TblPrimary_AOP, , False)
            'poc
            If Me.TransferPOC() = False Then
                Return False
            End If
            Me.SetOrderBalancePurchase()

            'MsgBox("belum untuk siklus purchase request")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CutOff_PurchaseModule")
            Return False
        End Try
    End Function

    Private Function CutOff_SalesModule() As Boolean

        Try
            If Me.TransferQuotation() = False Then
                Return False
            End If
            If Me.TransferSO() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_so", , , Me.TblPrimary_SO, , False)
            If Me.TransferDO() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_do", , , Me.TblPrimary_DO, , False)
            ''MsgBox("Backup DO")
            If Me.TransferStockOutSales() = False Then
                Return False
            End If
            If Me.TransferStockInSalesReturn() = False Then
                Return False
            End If
            If Me.TransferSalesReturn() = False Then
                Return False
            End If
            If Me.TransferQuickSalesReturn() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_sisr", , , Me.TblPrimary_SISR, , False)
            'Me.TransferMaster_UsingScript("a_sr", , , Me.TblPrimary_SR, , False)
            'Transfer invoice dp ada di dalam fungsi transfer sales invoice
            If Me.TransferSalesInvoice() = False Then
                Return False
            End If
            If Me.TransferQuickSalesInvoice() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_inv", , , Me.TblPrimary_Inv, , False)
            If Me.TransferSOC() = False Then
                Return False
            End If
            'update primary sales invoice pada saler return, karena saat transfer sales return
            'sales invoice belum di transfer jadi primary invoice masih -1
            Me.UpdatePrimaryInv_Salesreturn()
            Me.TrancuteNTransfer("tbltrans_salesinvoice_tax_invoice_setting")
            Me.SetOrderBalanceSales()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CutOff_SalesModule")
            Return False
        End Try
    End Function

    Private Function CutOff_HRModule() As Boolean
        Dim Desc_ As String = ""

        Try
            For i As Integer = 1 To 6
                If i = 1 Then
                    Desc_ = "Transfer Leave Application Data"
                ElseIf i = 2 Then
                    Desc_ = "Transfer Attendance Data"
                ElseIf i = 3 Then
                    Desc_ = "Transfer Loan Application Data"
                ElseIf i = 4 Then
                    Desc_ = "Transfer Loan Payment Data"
                ElseIf i = 5 Then
                    Desc_ = "Transfer Salary Data"
                ElseIf i = 6 Then
                    Desc_ = "Transfer Salary Payment Data"
                End If

                Me.RowUpdate(Desc_, Me.StatusOnProses)
                Me.TblListTransfer.DefaultView.RowFilter = Nothing
                Me.TblListTransfer.DefaultView.RowFilter = "[Description] = '" & Desc_.ToString & "'"
                For Each dv As DataRowView In Me.TblListTransfer.DefaultView
                    Me.RowListTransfer = dv.Row
                    Me.ProgressBar1.Value = 0
                    If Me.TransferMaster_UsingScript(Me.RowListTransfer.Tabel_Name) = False Then
                        Return False
                    End If
                Next
                Me.RowUpdate(Desc_, Me.StatusOK)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function CutOff_AccountingModule() As Boolean

        Try
            If Me.TransferGeneralJournal_CashBank() = False Then
                Return False
            End If
            If Me.TransferCreditDebitNote() = False Then
                Return False
            End If
            'Me.TransferMaster_UsingScript("a_cd", , , Me.TblPrimary_CD, , False)
            If Me.TransferPurchasePayment() = False Then
                Return False
            End If
            If Me.TransferSalesReceivement() = False Then
                Return False
            End If
            If Me.TransferCollectionLetter() = False Then
                Return False
            End If

            'TransferSalesReceivement_UsingScriptItemPayment_Repair()
            'MsgBox("Hati2 u/ revaluasi, harus dipikirkan ulang caranya")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CutOff_AccountingModule")
            Return False
        End Try
    End Function

    Private Function CutOff_CustomModule() As Boolean
        Dim Tbl As New DseTrans_CutOffDB.ListTransferDataTable
        Dim Row As DseTrans_CutOffDB.ListTransferRow
        Dim Tbl1, Tbl2, Tbl3 As New DataTable
        Dim row1, row2, row3 As DataRow

        Dim DelRow As Boolean = False
        Dim ListPrimary As New List(Of Double)

        Try
            Me.TblListTransfer.DefaultView.RowFilter = Nothing
            Me.TblListTransfer.DefaultView.RowFilter = "[Module] = 'Custom Module'"
            Tbl.Merge(Me.TblListTransfer.DefaultView.ToTable)
            For Each Row In Tbl.Rows
                Me.RowUpdate(Row.Description, Me.StatusOnProses)
                Me.TblListTransfer.DefaultView.RowFilter = Nothing
                Me.TblListTransfer.DefaultView.RowFilter = "[Tabel Name] = '" & Row.Tabel_Name & "'"
                If Me.TblListTransfer.DefaultView.Count > 0 Then
                    Me.RowListTransfer = Me.TblListTransfer.DefaultView(0).Row
                    Me.ProgressBar1.Value = 0

                    Tbl1 = New DataTable
                    Tbl2 = New DataTable
                    Tbl3 = New DataTable

                    Me.Label1.Text = Me.RowListTransfer.Description & " "

                    If Me.RowListTransfer.Description = "Transfer Custom Collection Order Data" Then
                        ListPrimary.Clear()

                        If Me.CutOff_CustomModule_Transfer("tbltrans_custom_collection_order", "tbltrans_custom_collection_order_item", _
                        "", Tbl1, Tbl2, Tbl3, ListPrimary, True) = False Then
                            Return False
                        End If
                    ElseIf Me.RowListTransfer.Description = "Transfer Custom Collection Log Data" Then
                        If Me.CutOff_CustomModule_Transfer("tbltrans_custom_collection_log", "tbltrans_custom_collection_log_item", _
                        "", Tbl1, Tbl2, Tbl3, ListPrimary) = False Then
                            Return False
                        End If
                    ElseIf Me.RowListTransfer.Description = "Transfer Custom Delivery Operation Data" Then
                        ListPrimary.Clear()

                        If Me.CutOff_CustomModule_Transfer("tbltrans_custom_delivery_operation", "tbltrans_custom_delivery_operation_item", _
                        "tbltrans_custom_delivery_operation_expedition", Tbl1, Tbl2, Tbl3, ListPrimary) = False Then
                            Return False
                        End If
                    ElseIf Me.RowListTransfer.Description = "Transfer Custom Sales Receipt Data" Then
                        ListPrimary.Clear()

                        If Me.CutOff_CustomModule_Transfer("tbltrans_custom_sales_receipt", "tbltrans_custom_sales_receipt_item", _
                        "", Tbl1, Tbl2, Tbl3) = False Then
                            Return False
                        End If
                    ElseIf Me.RowListTransfer.Description = "Transfer Custom Surat Tagihan Data" Then
                        ListPrimary.Clear()

                        If Me.CutOff_CustomModule_Transfer("tbltrans_custom_surattagihan", "tbltrans_custom_surattagihan_item", _
                        "", Tbl1, Tbl2, Tbl3) = False Then
                            Return False
                        End If
                    ElseIf Me.RowListTransfer.Description = "Transfer Custom Unit Cost Data" Then
                        ListPrimary.Clear()
                        If Me.CutOff_CustomModule_Transfer("tbltrans_custom_unitcost", "tbltrans_custom_unitcost_item", _
                        "tbltrans_custom_unitcost_data", Tbl1, Tbl2, Tbl3) = False Then
                            Return False
                        End If
                    Else
                        If Me.TransferMaster_UsingScript(Me.RowListTransfer.Tabel_Name) = False Then
                            Return False
                        End If
                    End If

                    Tbl1.Dispose()
                    Tbl2.Dispose()
                    Tbl3.Dispose()
                End If
                Me.RowUpdate(Row.Description, Me.StatusOK)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: CutOff_AccountingModule")
            Return False
        Finally
            Tbl.Dispose()
            Tbl1.Dispose()
            Tbl2.Dispose()
            Tbl3.Dispose()
        End Try
    End Function

    Private Function CutOff_CustomModule_Transfer(ByVal Tbl1_name As String, ByVal Tbl2_Name As String, ByVal Tbl3_Name As String, _
     ByRef Tbl1 As DataTable, ByRef Tbl2 As DataTable, ByRef Tbl3 As DataTable, Optional ByVal ListPrimary As List(Of Double) = Nothing, _
     Optional ByVal GunakanList As Boolean = False) As Boolean
        Dim Dta As New MySqlDataAdapter
        Dim cmd As New MySqlCommand
        Dim ConnSumber As New MySql.Data.MySqlClient.MySqlConnection(Me.SumberConstring)

        Dim row1, row2, row3 As DataRow
        Dim DelRow As Boolean = False
        Dim LblText As String = Me.Label1.Text

        Try
            cmd.Connection = ConnSumber

            cmd.CommandText = "select * from " & Tbl1_name & ""
            Dta.SelectCommand = cmd
            Dta.Fill(Tbl1)
            cmd.CommandText = "select * from " & Tbl2_Name & ""
            Dta.SelectCommand = cmd
            Dta.Fill(Tbl2)
            If Tbl3_Name <> "" Then
                cmd.CommandText = "select * from " & Tbl3_Name & ""
                Dta.SelectCommand = cmd
                Dta.Fill(Tbl3)
            End If

            If Tbl1_name <> "tbltrans_custom_unitcost" Then
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = Tbl1.Rows.Count
                For i As Integer = Tbl1.Rows.Count - 1 To 0 Step -1
                    row1 = Tbl1.Rows(i)
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.ProgressBar1.Value += 1
                    Me.Label1.Text = LblText & " " & ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100, 2, _
                    MidpointRounding.AwayFromZero), 2) + "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + _
                    ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)

                    DelRow = False
                    If Tbl1_name = "tbltrans_custom_collection_log" Then
                        If row1.Item("Payment Method") = 0 Then
                            'jika collection order yang di rujuk sudah di hapus maka collection log nya tidak usah di transfer
                            If ListPrimary.Contains(row1.Item("Primary Collection Order")) = True Then
                                Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & row1.Item("Primary")
                                For iCO As Integer = Tbl2.DefaultView.Count - 1 To 0 Step -1
                                    row2 = Tbl2.DefaultView(iCO).Row
                                    Tbl2.Rows.Remove(row2)
                                Next
                                Tbl1.Rows.Remove(row1)
                                GoTo Back2
                            End If
                        End If
                    End If
Back1:
                    Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & row1.Item("Primary")
                    For i2 As Integer = Tbl2.DefaultView.Count - 1 To 0 Step -1
                        row2 = Tbl2.DefaultView(i2).Row
                        If DelRow = True Then
                            Tbl2.Rows.Remove(row2)
                        Else
                            If Tbl1_name = "tbltrans_custom_delivery_operation" Then
                                If row2.Item("delivery description") = 0 Or row2.Item("delivery description") = 1 Or _
                                row2.Item("delivery description") = 2 Then 'Goods, Invoice_Exchange, Collection
                                    GoTo PrimaryInvoice
                                ElseIf row2.Item("delivery description") = 4 Then 'Goods_DO_Only
                                    row2.Item("Primary do") = Me.ObjCutOff_Core.GetPrimaryUpdate(row2.Item("Primary do"), Me.TblPrimary_DO)
                                    If row2.Item("Primary do") < 0 Then
                                        DelRow = True
                                        GoTo back1
                                    End If
                                End If
                            Else
PrimaryInvoice:
                                row2.Item("Primary Invoice") = Me.ObjCutOff_Core.GetPrimaryUpdate(row2.Item("Primary Invoice"), Me.TblPrimary_Inv)
                                If row2.Item("Primary Invoice") < 0 Then
                                    DelRow = True
                                    GoTo back1
                                End If
                            End If
                        End If
                    Next
Back2:
                    If DelRow = True Then
                        If GunakanList = True Then
                            If ListPrimary.Contains(CDbl(row1.Item("Primary"))) = False Then
                                ListPrimary.Add(CDbl(row1.Item("Primary")))
                            End If
                        End If
                        If Tbl3_Name <> "" Then
                            Tbl3.DefaultView.RowFilter = "[Primary Transaksi] = " & row1.Item("Primary")
                            For iDO As Integer = Tbl3.DefaultView.Count - 1 To 0 Step -1
                                row3 = Tbl3.DefaultView(iDO).Row
                                Tbl3.Rows.Remove(row3)
                            Next
                        End If
                        Tbl1.Rows.Remove(row1)
                    End If
                Next
            Else
                Dim Tbl1_T, Tbl2_T, Tbl3_T As New DataTable
                Dim row1_T As DataRow

                Tbl1_T = Tbl1.Copy
                Tbl2_T = Tbl2.Copy
                Tbl3_T = Tbl3.Copy
                Tbl1.Clear()
                Tbl2.Clear()
                Tbl3.Clear()

                Tbl1.DefaultView.RowFilter = "[Transaction Date] > #" & Me.EndDateCutOff & "#"
                For Each dvUC As DataRowView In Tbl1_T.DefaultView
                    row1_T = dvUC.Row
                    row1_T.Item("Primary AOP") = Me.ObjCutOff_Core.GetPrimaryUpdate(row1_T.Item("Primary AOP"), Me.TblPrimary_AOP)
                    If row1_T.Item("Primary AOP") > 0 Then
                        row1 = Tbl1.NewRow
                        For i As Integer = 0 To Tbl1_T.Columns.Count - 1
                            row1.Item(i) = row1_T.Item(i)
                        Next
                        Tbl1.Rows.Add(row1)

                        Tbl2_T.DefaultView.RowFilter = "[Primary Transaction] = " & row1_T.Item("Primary")
                        If Tbl2_T.DefaultView.Count > 0 Then
                            Tbl2.Merge(Tbl2_T.DefaultView.ToTable)
                        End If
                        Tbl3_T.DefaultView.RowFilter = "[Primary Transaction] = " & row1_T.Item("Primary")
                        If Tbl3_T.DefaultView.Count > 0 Then
                            Tbl3.Merge(Tbl3_T.DefaultView.ToTable)
                        End If
                    End If
                Next
            End If

            If Me.TransferMaster_UsingScript(Tbl1_name, Nothing, "", Tbl1, False, False) = False Then
                Return False
            End If
            If Me.TransferMaster_UsingScript(Tbl2_Name, Nothing, "", Tbl2, False, False) = False Then
                Return False
            End If
            If Tbl3_Name <> "" Then
                If Me.TransferMaster_UsingScript(Tbl3_Name, Nothing, "", Tbl3, False, False) = False Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    'TRANSFER CORE MODULE
    Private Function TransferMaster() As Boolean
        Dim ListPrimary As New List(Of Double)
        Dim count As Integer = 0

        Try
            Me.TblListTransfer.DefaultView.RowFilter = Nothing
            Me.TblListTransfer.DefaultView.RowFilter = "[Module] = 'Master data'"
            Me.TblListTransfer.DefaultView.Sort = "[No]"
            Me.RowListTransfer = Me.TblListTransfer.Rows(Me.TblListTransfer.DefaultView.Count - 1)
            For i As Integer = 1 To Me.RowListTransfer.No
                Me.dgvProgress.Item(Me.ColStatus.Index, i).Value = Me.StatusOnProses
                ListPrimary.Clear()

                count = 0
                Me.TblListTransfer.DefaultView.RowFilter = "[No] = " & i
                For Each dv As DataRowView In Me.TblListTransfer.DefaultView
                    Me.RowListTransfer = dv.Row
                    Me.ProgressBar1.Value = 0
                    count += 1
                    If Me.RowListTransfer.No > 0 And Me.RowListTransfer.Tabel_Name <> "" Then
                        If Me.TransferMaster_UsingScript(Me.RowListTransfer.Tabel_Name, ListPrimary) = False Then
                            Return False
                        End If
                    End If
                Next

                Me.dgvProgress.Item(Me.ColStatus.Index, i).Value = Me.StatusOK
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster")
            Return False
        End Try
    End Function

    Private Function TransferMaster_UsingScript(ByVal vTblName As String, Optional ByRef vListPrimary As List(Of Double) = Nothing, _
    Optional ByVal vTblName2 As String = "", Optional ByVal Tbl2 As DataTable = Nothing, Optional ByVal UpdatePrimary As Boolean = False, _
    Optional ByVal FillUlang As Boolean = True, Optional ByVal vTblName3 As String = "", Optional ByVal Tbl3 As DataTable = Nothing) As Boolean
        Dim Dta As New MySqlDataAdapter
        Dim Tbl As New DataTable
        Dim Row As DataRow

        Dim cmd As New MySqlCommand
        Dim ConnSumber As New MySql.Data.MySqlClient.MySqlConnection(Me.SumberConstring)
        Dim ConnTarget As New MySql.Data.MySqlClient.MySqlConnection(Me.TargetConstring)
        Dim tblDummy, tblDummy2, tblDummy3 As New Core_Module.dsFormat.tblFormatDataTable
        Dim rowsCommand As MySqlCommand = ConnTarget.CreateCommand()
        Dim rowsReader As MySqlDataReader

        Dim NamaKolom As String = ""
        Dim NamaKolom2 As String = ""
        Dim NamaKolom3 As String = ""
        Dim Value_, ValueInsert_, ValueInsert2_, ValueInsert3_, QueryInsert, QueryInsert2, QueryInsert3 As String
        Dim Sql As String = ""
        Dim Script As New MySql.Data.MySqlClient.MySqlScript(ConnTarget, Sql)
        Dim StartRowIndex As Double = 0
        Dim BatchQty As Double = 500
        Dim PrimaryMax As Double = -1
        Dim IdxColumn As Integer = 0

        Try
            If FillUlang = True Then
                cmd.Connection = ConnSumber
                cmd.CommandText = "select * from " + vTblName
                Dta.SelectCommand = cmd
                Dta.Fill(Tbl)
            Else
                Tbl = Tbl2.Copy
            End If

            Me.TransferMaster_DataUpdate(vTblName, Tbl, vListPrimary, PrimaryMax, Tbl2, Tbl3)

            'jika update primary header maka kolom primary pada table header tidak di hapus
            'karena masih di pakai untuk find & update primary transaksi pada table item nya.
            If UpdatePrimary = True Then
                IdxColumn = 1
            Else
                IdxColumn = 0
            End If

            If Tbl.Rows.Count > 0 Then
                ConnTarget.Open()
                tblDummy.Rows.Clear()
                rowsCommand.CommandText = "DESCRIBE " & vTblName
                rowsReader = rowsCommand.ExecuteReader()
                'dapatkan nama kolom
                For i As Integer = IdxColumn To Tbl.Columns.Count - 1
                    rowsReader.Read()
                    If rowsReader(0).ToString <> Tbl.Columns(i).ColumnName Then
                        rowsReader.Read()
                    End If
                    tblDummy.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, i)
                    If i <> Tbl.Columns.Count - 1 Then
                        NamaKolom &= "`" & Tbl.Columns(i).ColumnName & "`" & ","
                    Else
                        NamaKolom &= "`" & Tbl.Columns(i).ColumnName & "`"
                    End If
                Next
                ConnTarget.Close()

                If vTblName2 <> "" Then
                    ConnTarget.Open()
                    tblDummy2.Rows.Clear()
                    rowsCommand.CommandText = "DESCRIBE " & vTblName2
                    rowsReader = rowsCommand.ExecuteReader()
                    'dapatkan nama kolom
                    For i As Integer = 0 To Tbl2.Columns.Count - 1
                        rowsReader.Read()
                        If rowsReader(0).ToString <> Tbl2.Columns(i).ColumnName Then
                            rowsReader.Read()
                        End If
                        tblDummy2.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, i)
                        If i <> Tbl2.Columns.Count - 1 Then
                            NamaKolom2 &= "`" & Tbl2.Columns(i).ColumnName & "`" & ","
                        Else
                            NamaKolom2 &= "`" & Tbl2.Columns(i).ColumnName & "`"
                        End If
                    Next
                    ConnTarget.Close()
                End If

                If vTblName3 <> "" Then
                    ConnTarget.Open()
                    tblDummy3.Rows.Clear()
                    rowsCommand.CommandText = "DESCRIBE " & vTblName3
                    rowsReader = rowsCommand.ExecuteReader()
                    'dapatkan nama kolom
                    For i As Integer = 0 To Tbl3.Columns.Count - 1
                        rowsReader.Read()
                        If rowsReader(0).ToString <> Tbl3.Columns(i).ColumnName Then
                            rowsReader.Read()
                        End If
                        tblDummy3.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, i)
                        If i <> Tbl3.Columns.Count - 1 Then
                            NamaKolom3 &= "`" & Tbl3.Columns(i).ColumnName & "`" & ","
                        Else
                            NamaKolom3 &= "`" & Tbl3.Columns(i).ColumnName & "`"
                        End If
                    Next
                    ConnTarget.Close()
                End If

                ConnTarget.Open()
                Me.SetProgressBarNLabel("", BatchQty, True, Tbl.Rows.Count)
                While StartRowIndex < Tbl.Rows.Count
                    ValueInsert_ = ""
                    ValueInsert2_ = ""
                    ValueInsert3_ = ""

                    For i As Double = 0 To BatchQty - 1 Step 1
                        If Tbl.Rows.Count > StartRowIndex + i Then
                            If Me.Bgw1.CancellationPending Then
                                Return False
                            End If

                            Value_ = ""
                            Row = Tbl.Rows(StartRowIndex + i)
                            For i2 As Integer = IdxColumn To Tbl.Columns.Count - 1
                                If i2 <> IdxColumn Then
                                    Value_ &= ","
                                End If
                                If IdxColumn = 1 And IdxColumn > 0 Then
                                    Value_ &= Me.FrmBackup.FormatValue(tblDummy, (i2 - IdxColumn), Tbl.Rows(Tbl.Rows.IndexOf(Row))(i2))
                                Else
                                    Value_ &= Me.FrmBackup.FormatValue(tblDummy, i2, Tbl.Rows(Tbl.Rows.IndexOf(Row))(i2))
                                End If
                            Next

                            ValueInsert_ &= "(" & Value_ & "), "

                            If vTblName2 <> "" Then
                                'jika primary max = -1 maka tidak ada update primary transaksi
                                If PrimaryMax <> -1 Then
                                    PrimaryMax += 1
                                End If

                                Me.TransferMaster_UsingScript2(Tbl2, vTblName2, ValueInsert2_, Row.Item("Primary"), PrimaryMax, FrmBackUp, tblDummy2)
                                If vTblName3 <> "" Then
                                    Me.TransferMaster_UsingScript2(Tbl3, vTblName3, ValueInsert3_, Row.Item("Primary"), PrimaryMax, FrmBackUp, tblDummy3)
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next

                    ValueInsert_ = ValueInsert_.Remove(ValueInsert_.Length - 2, 2) & ";"
                    QueryInsert = "INSERT INTO " & vTblName & "(" & NamaKolom & ") VALUES " & vbCrLf & ValueInsert_ & vbCrLf & vbCrLf & "COMMIT;"
                    Sql = QueryInsert
                    Script.Query = Sql
                    Script.Execute()
                    If ValueInsert2_ <> "" Then
                        ValueInsert2_ = ValueInsert2_.Remove(ValueInsert2_.Length - 2, 2) & ";"
                        QueryInsert2 = "INSERT INTO " & vTblName2 & "(" & NamaKolom2 & ") VALUES " & vbCrLf & ValueInsert2_ & vbCrLf & vbCrLf & "COMMIT;"
                        Sql = QueryInsert2
                        Script.Query = Sql
                        Script.Execute()
                    End If
                    If ValueInsert3_ <> "" Then
                        ValueInsert3_ = ValueInsert3_.Remove(ValueInsert3_.Length - 2, 2) & ";"
                        QueryInsert3 = "INSERT INTO " & vTblName3 & "(" & NamaKolom3 & ") VALUES " & vbCrLf & ValueInsert3_ & vbCrLf & vbCrLf & "COMMIT;"
                        Sql = QueryInsert3
                        Script.Query = Sql
                        Script.Execute()
                    End If

                    Me.SetProgressBarNLabel(vTblName, BatchQty)

                    StartRowIndex += BatchQty
                End While

                ConnTarget.Close()
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster_UsingScript")
            Return False
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            ConnSumber = Nothing
            ConnTarget = Nothing
            Script = Nothing
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferMaster_UsingScript2(ByVal Tbl2 As DataTable, ByVal vTblName As String, _
    ByRef ValueInsert_ As String, ByVal vPrimaryTransaksi As Double, ByVal vPrimaryMax As Double, _
    ByVal FrmBackUp As Core_Module.FrmBackUpData, ByVal tblDummy As Core_Module.dsFormat.tblFormatDataTable) As Boolean
        Dim Value_ As String = ""
        Dim Row2 As DataRow
        Dim PrimaryString As String = ""

        Try
            If vTblName = "tbltrans_surattagihan_item" Then
                PrimaryString = "Primary Transaksi"
            Else
                PrimaryString = "Primary Jurnal"
            End If

            Tbl2.DefaultView.RowFilter = "[" & PrimaryString & "] = " & vPrimaryTransaksi
            If Tbl2.DefaultView.Count > 0 Then
                For ii As Integer = 0 To Tbl2.DefaultView.Count - 1
                    Value_ = ""
                    Row2 = Tbl2.DefaultView(ii).Row
                    For i As Integer = 0 To Tbl2.Columns.Count - 1
                        If i <> 0 Then
                            Value_ &= ","
                        End If

                        If Tbl2.Columns(i).ColumnName = PrimaryString Then
                            If vPrimaryMax <> -1 Then
                                Value_ &= vPrimaryMax.ToString
                            Else
                                Value_ &= Me.FrmBackup.FormatValue(tblDummy, i, Tbl2.DefaultView(ii)(i))
                            End If
                        Else
                            Value_ &= Me.FrmBackup.FormatValue(tblDummy, i, Tbl2.DefaultView(ii)(i))
                        End If
                    Next

                    ValueInsert_ &= "(" & Value_ & "), "
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster_UsingScript")
            Return False
        End Try
    End Function

    Private Sub TransferMaster_DataUpdate(ByVal vTblName As String, ByRef vTbl As DataTable, _
    ByRef vListPrimary As List(Of Double), ByRef PrimaryMax As Double, ByRef vTbl2 As DataTable, ByRef vTbl3 As DataTable)
        Dim Tbl2 As New DataTable
        Dim Row, Row2, Row3 As DataRow

        Try
            If vTblName = "tbltrans_salesbudget" Or vTblName = "tbltrans_material_planning_group" Or _
            vTblName = "tbltrans_material_planning" Or vTblName = "tbltrans_shipping_forecast" Or _
            vTblName = "tbltrans_quotation" Then
                Tbl2 = vTbl.Copy
                vTbl.Rows.Clear()

                Dim DateFilter_ As String = ""
                If vTblName = "tbltrans_salesbudget" Then
                    DateFilter_ = "period"
                ElseIf vTblName = "tbltrans_material_planning_group" Then
                    DateFilter_ = "tanggal"
                ElseIf vTblName = "tbltrans_material_planning" Or vTblName = "tbltrans_shipping_forecast" Then
                    DateFilter_ = "date"
                ElseIf vTblName = "tbltrans_quotation" Then
                    DateFilter_ = "Tanggal Transaksi"
                End If

                Tbl2.DefaultView.RowFilter = "[" & DateFilter_ & "] > #" & Me.EndDateCutOff & "#"
                If Tbl2.DefaultView.Count > 0 Then
                    For Each dv As DataRowView In Tbl2.DefaultView
                        Row2 = dv.Row
                        Row = vTbl.NewRow
                        For ICol As Integer = 0 To Tbl2.Columns.Count - 1
                            Row.Item(ICol) = Row2.Item(ICol)
                        Next
                        vTbl.Rows.Add(Row)

                        vListPrimary.Add(Row2.Item(0))
                    Next
                End If
            ElseIf vTblName = "tbltrans_salesbudget_item" Or vTblName = "tbltrans_quotation_item" Or vTblName = "tbltrans_quotation_itemdelivery" Or _
            vTblName = "tbltrans_material_planning_group_item_group" Or vTblName = "tbltrans_material_planning_group_item_product" Or _
            vTblName = "tbltrans_material_planning_group_item_product_group" Or vTblName = "tbltrans_material_planning_group_itemsupplier" Or _
            vTblName = "tbltrans_material_planning_group_item_adjustment" Or vTblName = "tbltrans_material_planning_group_item_bundling" Or _
            vTblName = "tbltrans_material_planning_group_item_invoice" Or vTblName = "tbltrans_material_planning_group_item_sr" Or _
            vTblName = "tbltrans_material_planning_item" Or vTblName = "tbltrans_material_planning_item_qty" Or _
            vTblName = "tbltrans_material_planning_itemdo" Or vTblName = "tbltrans_material_planning_itemgit" Or _
            vTblName = "tbltrans_material_planning_itempob" Or vTblName = "tbltrans_material_planning_itemsob" Or _
            vTblName = "tbltrans_material_planning_itempurch_req" Or vTblName = "tbltrans_shipping_forecast_itemstock" Or _
            vTblName = "tbltrans_shipping_forecast_item_pob" Or vTblName = "tbltrans_shipping_forecast_itemdelivery" Then
                Dim PrimaryAcuan As String = ""
                Dim ApakahSet As Boolean = True
                If vTblName = "tbltrans_material_planning_group_item_group" Then 'MP ITEM GROUP
                    Me.Label1.Text = "Get Material Planning Item Group Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_group_item_product" Then 'MP ITEM PRODUCT
                    Me.Label1.Text = "Get Material Planning Item Product Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_group_item_product_group" Then 'MP ITEM PRODUCT GROUP
                    Me.Label1.Text = "Get Material Planning Item Product Group Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_group_itemsupplier" Then 'MP ITEM SUPPLIER
                    Me.Label1.Text = "Get Material Planning Item Supplier Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_group_item_adjustment" Then 'MP ITEM ADJUSMENT
                    PrimaryAcuan = "primary adj"
                    Me.Label1.Text = "Get Material Planning Adjusment Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_group_item_bundling" Then 'MP ITEM BUNDLING
                    PrimaryAcuan = "primary bundling"
                    Me.Label1.Text = "Get Material Planning Bundling Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_group_item_invoice" Then 'MP ITEM INVOICE
                    PrimaryAcuan = "primary invoice"
                    Me.Label1.Text = "Get Material Planning Invoice Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_group_item_sr" Then 'MP ITEM SR
                    PrimaryAcuan = "primary sr"
                    Me.Label1.Text = "Get Material Planning Sales Return Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_item" Then 'MP ITEM
                    Me.Label1.Text = "Get Material Planning Item Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_item_qty" Then 'MP ITEM QTY
                    Me.Label1.Text = "Get Material Planning Item Qty Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_itemdo" Then 'MP ITEM DO
                    PrimaryAcuan = "primary do rec"
                    Me.Label1.Text = "Get Material Planning Delivery Order Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_itemgit" Then 'MP ITEM GIT
                    PrimaryAcuan = "primary packing list rec"
                    Me.Label1.Text = "Get Material Planning GIT Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_itempob" Then 'MP ITEM POB
                    PrimaryAcuan = "primary po rec"
                    Me.Label1.Text = "Get Material Planning Purchase Order Balance Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_itemsob" Then 'MP ITEM SOB
                    PrimaryAcuan = "primary so rec"
                    Me.Label1.Text = "Get Material Planning Sales Order Balance Data, Please Wait..."
                ElseIf vTblName = "tbltrans_material_planning_itempurch_req" Then 'MP ITEM PURCH REQ
                    PrimaryAcuan = "primary pr rec"
                    Me.Label1.Text = "Get Material Planning Purchase Request Data, Please Wait..."
                ElseIf vTblName = "tbltrans_shipping_forecast_item_pob" Then 'SF ITEM POB
                    PrimaryAcuan = "primary po"
                    Me.Label1.Text = "Get Shipping Forecast Purchase Order Balance Data, Please Wait..."
                ElseIf vTblName = "tbltrans_shipping_forecast_itemdelivery" Then 'SF ITEM DELIVERY
                    PrimaryAcuan = "primary so"
                    Me.Label1.Text = "Get Shipping Forecast Sales Order Balance Data, Please Wait..."
                ElseIf vTblName = "tbltrans_shipping_forecast_itemstock" Then 'SF ITEM STOCK
                    Me.Label1.Text = "Get Material Planning Bundling Data, Please Wait..."
                Else
                    ApakahSet = False
                End If

                If ApakahSet = True Then
                    vTbl.Columns.RemoveAt(0)
                End If

                Tbl2 = vTbl.Copy
                vTbl.Rows.Clear()

                If vListPrimary.Count > 0 Then
                    If vTblName = "tbltrans_quotation_item" Then
                        Dim PrimaryMin_ As Double = vListPrimary.Item(0)
                        Dim PrimaryMax_ As Double = vListPrimary.Item(vListPrimary.Count - 1)
                        Tbl2.DefaultView.RowFilter = "[Primary Transaksi] >= " & PrimaryMin_.ToString & " AND [Primary Transaksi] <= " & PrimaryMax_.ToString
                        If Tbl2.DefaultView.Count > 0 Then
                            vTbl.Merge(Tbl2.DefaultView.ToTable)
                            For i As Integer = vTbl.Rows.Count - 1 To 0 Step -1
                                Row = vTbl.Rows(i)
                                If vListPrimary.Contains(Row.Item("Primary Transaksi")) = False Then
                                    vTbl.Rows.Remove(Row)
                                End If
                            Next
                        End If
                    Else
                        For Each Primary_ As Double In vListPrimary
                            Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & Primary_
                            If Tbl2.DefaultView.Count > 0 Then
                                vTbl.Merge(Tbl2.DefaultView.ToTable)
                            End If
                        Next
                    End If
                End If
            ElseIf vTblName = "tbltrans_material_planning_detailotorisasipo" Then
                Dim PrimaryPO_New As Double = -1
                For i As Integer = vTbl.Rows.Count - 1 To 0 Step -1
                    Row = vTbl.Rows(i)
                    PrimaryPO_New = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary po"), Me.TblPrimary_PO)
                    If PrimaryPO_New > 0 Then
                        Row.Item("primary po") = PrimaryPO_New
                    Else
                        vTbl.Rows.Remove(Row)
                    End If
                Next
            ElseIf vTblName = "tblmaster_akun" Then
                Dim rowBB As DseTrans_CutOffDB.Trial_BalanceRow
                Me.TBL_Beginning_Balance.Rows.Clear()
                Me.TBL_Trial_Balance.Rows.Clear()
                Me.GetBeginningBalance(Me.EndDateCutOff, Me.TBL_Beginning_Balance)
                For Each Row In vTbl.Rows
                    Me.TBL_Trial_Balance.DefaultView.RowFilter = "[Primary Akun] = " & Row.Item("Primary")
                    If Me.TBL_Trial_Balance.DefaultView.Count > 0 Then
                        rowBB = Me.TBL_Trial_Balance.DefaultView(0).Row
                        If rowBB.Saldo_Normal = 0 Then
                            Row.Item("Beginning Balance") = rowBB.Debet
                            Row.Item("Beggining Balance Original Cur") = rowBB.Debet_Original
                        Else
                            Row.Item("Beginning Balance") = rowBB.Credit
                            Row.Item("Beggining Balance Original Cur") = rowBB.Credit_Original
                        End If
                    End If
                Next
            ElseIf vTblName = "tbltrans_accounting_closing" Then
                Row = vTbl.Rows(0)
                If Row.Item("Convertion Date") < Me.EndDateCutOff Then
                    Row.Item("Convertion Date") = Me.EndDateCutOff.AddDays(1)
                End If
                If Row.Item("Convertion Date Payroll") < Me.EndDateCutOff Then
                    Row.Item("Convertion Date Payroll") = Me.EndDateCutOff.AddDays(1)
                End If
            ElseIf vTblName = "tbltrans_jurnal" Then
                Me.Label1.Text = "Get Journal Data, Please Wait..."
                Tbl2 = vTbl.Copy
                vTbl.Rows.Clear()

                'fill general journal
                Tbl2.DefaultView.RowFilter = "[Tipe Journal] = " & AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal & _
                " AND [Tipe Jurnal 2] = " & AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping & _
                " AND [Date] > #" & Me.EndDateCutOff & "#"
                vTbl.Merge(Tbl2.DefaultView.ToTable)
                'fill cash and bank
                Tbl2.DefaultView.RowFilter = "[Tipe Journal] = " & AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal & _
                " AND [Tipe Jurnal 2] = " & AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Cash_n_Bank & _
                " AND [Date] > #" & Me.EndDateCutOff & "#"
                vTbl.Merge(Tbl2.DefaultView.ToTable)
                'fill hr
                Tbl2.DefaultView.RowFilter = "[Tipe Journal] = " & AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal & _
                " AND ([Tipe Jurnal 2] = " & AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Payment & _
                " OR [Tipe Jurnal 2] = " & AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan & _
                " OR [Tipe Jurnal 2] = " & AccModule.ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan_Payment & ")"
                vTbl.Merge(Tbl2.DefaultView.ToTable)
                'fill end of year
                Tbl2.DefaultView.RowFilter = "[Tipe Journal] = " & AccModule.ClsTransaksiJournal.EnumTipeJournal.ClosingJournal & _
               " AND [Tipe Jurnal 2] = " & AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping & _
               " AND [Date] > #" & Me.EndDateCutOff & "#"
                vTbl.Merge(Tbl2.DefaultView.ToTable)
                'filter jurnal item
                vTbl.DefaultView.Sort = "[Primary] ASC"
                If vTbl.DefaultView.Count > 0 Then
                    Dim PrimaryJurnalMin As Double = 1
                    Dim PrimaryJurnalMax As Double = 999999999
                    Row = vTbl.DefaultView(0).Row
                    PrimaryJurnalMin = Row.Item("Primary")
                    Row = vTbl.DefaultView(vTbl.DefaultView.Count - 1).Row
                    PrimaryJurnalMax = Row.Item("Primary")

                    Dim Tbl_ItemJur2 As New DataTable
                    Tbl_ItemJur2 = vTbl2.Copy
                    vTbl2.Clear()
                    Tbl_ItemJur2.DefaultView.RowFilter = "[Primary Jurnal] >= " & PrimaryJurnalMin & " AND [Primary Jurnal] <= " & PrimaryJurnalMax
                    If Tbl_ItemJur2.DefaultView.Count > 0 Then
                        vTbl2.Merge(Tbl_ItemJur2.DefaultView.ToTable)
                    End If
                    Tbl_ItemJur2.Clear()

                    Dim ListPrimaryJur As New List(Of Double)
                    For Each Row In vTbl.Rows
                        If ListPrimaryJur.Contains(Row.Item("Primary")) = False Then
                            ListPrimaryJur.Add(Row.Item("Primary"))
                        End If
                    Next

                    For iii As Integer = vTbl2.Rows.Count - 1 To 0 Step -1
                        Row2 = vTbl2.Rows(iii)
                        If ListPrimaryJur.Contains(Row2.Item("Primary Jurnal")) = False Then
                            vTbl2.Rows.Remove(Row2)
                        End If
                    Next

                    For iii As Integer = vTbl3.Rows.Count - 1 To 0 Step -1
                        Row3 = vTbl3.Rows(iii)
                        If ListPrimaryJur.Contains(Row3.Item("Primary Jurnal")) = False Then
                            vTbl3.Rows.Remove(Row3)
                        End If
                    Next
                End If

                PrimaryMax = Me.ObjCutOff.GetPrimaryMaxJurnal(Me.TargetConstring)
            ElseIf vTblName = "tbltrans_purchase_closing" Then
                Tbl2 = vTbl.Copy
                vTbl.Rows.Clear()
                Dim DateFilter As Date = Me.ObjCD.GetFirstDateofTheMonth(Me.EndDateCutOff)
                Dim ObjJurnal As New AccModule.ClsTransaksiJournal
                Dim PrimaryJurnalUpdate As Double = -1
                Dim NoTransJurnal As String = ""

                Me.SetTargetConnectionString()
                Tbl2.DefaultView.RowFilter = "[Closing Period] >= #" & DateFilter & "#"
                If Tbl2.DefaultView.Count > 0 Then
                    For Each dv As DataRowView In Tbl2.DefaultView
                        Row2 = dv.Row
                        NoTransJurnal = "COGS-" & Format(Row2.Item("Closing Period"), "MMM yyyy")
                        Row2.Item("primary jurnal") = ObjJurnal.GetPrimary(NoTransJurnal)
                        Row = vTbl.NewRow
                        For ICol As Integer = 0 To Tbl2.Columns.Count - 1
                            Row.Item(ICol) = Row2.Item(ICol)
                        Next
                        vTbl.Rows.Add(Row)

                        vListPrimary.Add(Row2.Item(0))
                    Next
                End If
                Me.SetDefaultConnectionString()
                ObjJurnal = Nothing
            ElseIf vTblName = "tbltrans_purchase_closing_item" Or vTblName = "tbltrans_purchase_closing_item_akun" Or _
            vTblName = "tbltrans_purchase_closing_item_detail" Or vTblName = "tbltrans_purchase_closing_item_inadjustment" Or _
            vTblName = "tbltrans_purchase_closing_item_tipe_adjustment" Then
                Tbl2 = vTbl.Copy
                vTbl.Rows.Clear()
                For Each Primary_ As Double In vListPrimary
                    Tbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & Primary_
                    If Tbl2.DefaultView.Count > 0 Then
                        vTbl.Merge(Tbl2.DefaultView.ToTable)
                    End If
                Next

                'Untuk tbltrans_purchase_closing_item & tbltrans_purchase_closing_item_akun transfer biasa
                If vTblName = "tbltrans_purchase_closing_item_detail" Then
                    Me.Label1.Text = "Get Purchase Closing Detail Data. Please wait..."
                    Me.ProgressBar1.Value = 0
                    Me.ProgressBar1.Maximum = vTbl.Rows.Count
                    For Each Row In vTbl.Rows
                        'Me.ProgressBar1.Value += 1
                        'Me.Label1.Text = "Get Purchase Closing Detail Data " & ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100, 2, _
                        'MidpointRounding.AwayFromZero), 2) + "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + _
                        'ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)

                        If Row.Item("tipe item detail") <> 0 And Row.Item("tipe item detail") <> 3 Then
                            If Row.Item("tipe transaksi asal") = 0 Then  'InPurchase -> primary aop
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_AOP)
                            ElseIf Row.Item("tipe transaksi asal") = 1 Then  'InSR -> primary stock in sales return
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_SISR)
                            ElseIf Row.Item("tipe transaksi asal") = 2 Then  'InAdjust
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_InAdj)
                            ElseIf Row.Item("tipe transaksi asal") = 3 Then  'InStockTakeAdjust -> primary stock taking
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_Taking)
                            ElseIf Row.Item("tipe transaksi asal") = 4 Then  'OutSales - > primary invoice
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_Inv)
                            ElseIf Row.Item("tipe transaksi asal") = 5 Then  'OutPR
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_PR)
                            ElseIf Row.Item("tipe transaksi asal") = 6 Then  'OutStockTakeAdjust -> primary stock taking
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_Taking)
                            ElseIf Row.Item("tipe transaksi asal") = 7 Then  'OutAdjust
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_OutAdj)
                            ElseIf Row.Item("tipe transaksi asal") = 8 Then  'OutBundling
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_Bund)
                            ElseIf Row.Item("tipe transaksi asal") = 9 Then  'InBundling
                                Row.Item("primary transaksi asal") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("primary transaksi asal"), Me.TblPrimary_Bund)
                            End If
                        End If
                    Next
                ElseIf vTblName = "tbltrans_purchase_closing_item_inadjustment" Then
                    Me.Label1.Text = "Get Purchase Closing Adjusment Data. Please Wait..."
                    Me.ProgressBar1.Value = 0
                    Me.ProgressBar1.Maximum = vTbl.Rows.Count
                    For Each Row In vTbl.Rows
                        'Me.ProgressBar1.Value += 1
                        'Me.Label1.Text = "Get Purchase Closing Adjusment Data " & ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100, 2, _
                        'MidpointRounding.AwayFromZero), 2) + "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + _
                        'ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)

                        If Row.Item("Tipe Dokumen Acuan") = 0 Then  'StockInAdjust
                            Row.Item("Primary Dokumen Acuan") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("Primary Dokumen Acuan"), Me.TblPrimary_InAdj)
                        ElseIf Row.Item("Tipe Dokumen Acuan") = 1 Then  'StockTakingInAdjustment
                            Row.Item("Primary Dokumen Acuan") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("Primary Dokumen Acuan"), Me.TblPrimary_Taking)
                        ElseIf Row.Item("Tipe Dokumen Acuan") = 2 Then  'StockBundlingAdjustment
                            Row.Item("Primary Dokumen Acuan") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("Primary Dokumen Acuan"), Me.TblPrimary_Bund)
                        ElseIf Row.Item("Tipe Dokumen Acuan") = 3 Then  'StockInSalesReturnPeriodeSama -> primary stock in sales return
                            Row.Item("Primary Dokumen Acuan") = Me.ObjCutOff_Core.GetPrimaryUpdate(Row.Item("Primary Dokumen Acuan"), Me.TblPrimary_SISR)
                        End If
                    Next
                End If
            ElseIf vTblName = "tbltrans_surattagihan" Then
                Me.Label1.Text = "Get Colletion Letter Data, Please Wait..."
                vListPrimary.Clear()

                For Each Row2 In vTbl2.Rows
                    If vListPrimary.Contains(Row2.Item("Primary Transaksi")) = False Then
                        If Me.ObjCutOff_Core.GetPrimaryUpdate(Row2.Item("Primary Invoice"), Me.TblPrimary_Inv) = -1 Then
                            vListPrimary.Add(Row2.Item("Primary Transaksi"))
                        End If
                    End If
                Next

                For i As Integer = vTbl.Rows.Count - 1 To 0 Step -1
                    Row = vTbl.Rows(i)
                    If vListPrimary.Contains(Row.Item("Primary")) = False Then
                        vTbl2.DefaultView.RowFilter = "[Primary Transaksi] = " & Row.Item("Primary")
                        If vTbl2.DefaultView.Count > 0 Then
                            For i2 As Integer = vTbl2.DefaultView.Count - 1 To 0 Step -1
                                Row2 = vTbl2.DefaultView(i2).Row
                                vTbl2.Rows.Remove(Row2)
                            Next
                        End If
                        vTbl.Rows.Remove(Row)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster_DataUpdate")
        Finally
            Tbl2.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Sub

    Private Sub GetBeginningBalance(ByVal vEndDate As Date, ByVal TBL As DseTrans_CutOffDB.tbl_beginning_balanceDataTable)
        Dim rowBeginningBalance As DseTrans_CutOffDB.tbl_beginning_balanceRow
        Dim row As DseTrans_CutOffDB.Trial_BalanceRow
        Dim filter As String = ""
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim CYE As Decimal = 0
        Dim PrimaryAkunCYE As Decimal = 0

        Try
            Me.DTA_Beginning_Balance.Connection.ConnectionString = Me.SumberConstring
            Me.DTA_Beginning_Balance.FillBy(Me.TBL_Beginning_Balance, vEndDate)
            Me.DTA_Trial_Balance.Connection.ConnectionString = Me.SumberConstring
            Me.DTA_Trial_Balance.Fill(Me.TBL_Trial_Balance)

            '2.a. Tambahkan akun current year earning, jika pada bulan tersebut tidak ada jurnal closing
            If Me.ObjJurnal.GetJml(AccModule.ClsTransaksiJournal.EnumTipeJournal.ClosingJournal, _
            Me.ObjCD.GetFisrtDateOfTheYear(vEndDate), Me.ObjCD.GetLasDateOfTheMonth(vEndDate)) = 0 Then
                CYE = Me.ObjJurnal.GetCurrentYearEarning1(vEndDate, -1, -1)
                PrimaryAkunCYE = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning)

                For Each rowBeginningBalance In Me.TBL_Beginning_Balance.Rows
                    If rowBeginningBalance.Primary = PrimaryAkunCYE Then
                        If CYE > 0 Then
                            rowBeginningBalance.Kredit += Math.Abs(CYE)
                        Else
                            rowBeginningBalance.Debet += Math.Abs(CYE)
                        End If
                    End If
                Next
            End If

            For Each row In Me.TBL_Trial_Balance.Rows
                filter = "Primary = " + row.Primary_Akun.ToString
                Me.TBL_Beginning_Balance.DefaultView.RowFilter = filter
                If Me.TBL_Beginning_Balance.DefaultView.Count > 0 Then
                    rowBeginningBalance = Me.TBL_Beginning_Balance.DefaultView(0).Row
                    row.Debet += rowBeginningBalance.Debet
                    row.Credit += rowBeginningBalance.Kredit
                    row.Debet_Original += rowBeginningBalance.Debet_Original
                    row.Credit_Original += rowBeginningBalance.Credit_Original

                    If row.Saldo_Normal = 0 Then
                        row.Debet = row.Beginning_Balance + row.Debet - row.Credit
                        row.Credit = 0
                        row.Debet_Original = row.Beginning_Balance + row.Debet_Original - row.Credit_Original
                        row.Credit_Original = 0
                    Else
                        row.Credit = row.Beginning_Balance + row.Credit - row.Debet
                        row.Debet = 0
                        row.Credit_Original = row.Beginning_Balance + row.Credit_Original - row.Debet_Original
                        row.Debet_Original = 0
                    End If
                End If
            Next

            '3. Hitung sum untuk setiap data header 
            Me.ObjCutOff.UpdateSaldoHeader(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: GetBeginningBalance")
        End Try
    End Sub

    Private Function TransferAPBalance() As Boolean
        Dim listPurchase As New List(Of String)

        Dim frmLaporanAPSummary_ As New AccModule.FrmLaporanAPSummary
        Dim Tbl_Supplier As New AccModule.DseLaporanAPSummary.tblmaster_supplierDataTable
        Dim row_Supplier As AccModule.DseLaporanAPSummary.tblmaster_supplierRow

        'AOP Otomatis
        Dim dtaAOP As New Purchase_Module.DseTrans_CutOffDBTableAdapters.tbltrans_aopTableAdapter
        Dim tblAOP As New Purchase_Module.DseTrans_CutOffDB.tbltrans_aopDataTable
        Dim rowAOP As Purchase_Module.DseTrans_CutOffDB.tbltrans_aopRow
        'AOP Manual
        Dim dtaAOP2 As New AccModule.DseTransaksiPurchaseBalanceTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim tblAOP2 As New AccModule.DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualDataTable
        Dim rowAOP2 As AccModule.DseTransaksiPurchaseBalance.tbltrans_historicalpurchasebalancemanualRow
        'Purchase Return Otomatis
        Dim dtaAOPReturn As New Purchase_Module.DseTrans_CutOffDBTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim tblAOPReturn As New Purchase_Module.DseTrans_CutOffDB.tbltrans_purchasereturnDataTable
        Dim rowAOPReturn As Purchase_Module.DseTrans_CutOffDB.tbltrans_purchasereturnRow
        'Purchase Return Manual
        Dim dtaAOPReturn2 As New Purchase_Module.DseTrans_CutOffDBTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim tblAOPReturn2 As New Purchase_Module.DseTrans_CutOffDB.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim rowAOPReturn2 As Purchase_Module.DseTrans_CutOffDB.tbltrans_historicalpurchasereturnbalancemanualRow
        'Credit - Debit Note
        Dim dtaCD_Note As New AccModule.DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim tblCD_Note As New AccModule.DseTransaksiCreditNote.tbltrans_credit_noteDataTable
        Dim rowCD_Note As AccModule.DseTransaksiCreditNote.tbltrans_credit_noteRow

        Dim frmTransaksiPurchaseBalanceManual_ As New AccModule.FrmTransaksiPurchaseBalanceManual
        Dim frmTransaksiPurchaseReturnBalanceManual_ As New AccModule.FrmTransaksiPurchaseReturnBalanceManual

        Dim rowIdx As Integer = 0
        Dim ApakahTrans As Boolean = False
        Dim Filter As String = ""
        Dim _NoTransaksi As String = ""
        Dim _TglTransaksi As Date = Now
        Dim _DueDate As Date = Now
        Dim _PrimaryAkunAP As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Purchase & Purchase Return Begining Balance Data", Me.StatusOnProses)
            Me.Label1.Text = "Get Purchase And Purchase Return Begining Balance Data, Please Wait..."

            Me.SetDefaultConnectionString()
            dtaAOP.Connection.ConnectionString = Me.SumberConstring
            dtaAOPReturn.Connection.ConnectionString = Me.SumberConstring
            dtaAOP.Fill(tblAOP, Me.EndDateCutOff)
            dtaAOP2.Fill(tblAOP2)
            dtaAOPReturn.Fill(tblAOPReturn, Me.EndDateCutOff)
            dtaAOPReturn2.Fill(tblAOPReturn2)
            dtaCD_Note.Fill(tblCD_Note)
            frmLaporanAPSummary_.FromCutOff_APSummary(Tbl_Supplier, Me.EndDateCutOff.Date, Me.EndDateCutOff.Date)
            Me.SetTargetConnectionString()

            'Transfer Purchase Begining Balance Data
            Me.SetProgressBarNLabel("", BatchProgBar, True, Tbl_Supplier.Count)
            For Each row_Supplier In Tbl_Supplier.Rows
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Purchase And Purchase Return Begining Balance Data", BatchProgBar)

                ApakahTrans = False
                If row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Then
                    tblAOP.DefaultView.RowFilter = "[Primary] = " & row_Supplier.Primary_Transaksi_Asal.ToString
                    If tblAOP.DefaultView.Count > 0 Then
                        rowAOP = tblAOP.DefaultView(0).Row
                        _NoTransaksi = rowAOP.No_Transaksi
                        _TglTransaksi = rowAOP.Tanggal_Transaksi
                        If IsNumeric(rowAOP.Payment_Term) Then
                            _DueDate = rowAOP.Tanggal_Transaksi.AddDays(rowAOP.Payment_Term)
                        Else
                            _DueDate = rowAOP.Tanggal_Transaksi
                        End If
                        _PrimaryAkunAP = rowAOP.primary_akun_ap
                        ApakahTrans = True
                    End If
                ElseIf row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual Then
                    tblAOP2.DefaultView.RowFilter = "[Primary] = " & row_Supplier.Primary_Transaksi_Asal.ToString
                    If tblAOP2.DefaultView.Count > 0 Then
                        rowAOP2 = tblAOP2.DefaultView(0).Row
                        _NoTransaksi = rowAOP2.AOP_Number
                        _TglTransaksi = rowAOP2.AOP_Date
                        _DueDate = rowAOP2.Due_Date
                        _PrimaryAkunAP = rowAOP2.primary_ap_account
                        ApakahTrans = True
                    End If
                ElseIf row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Then
                    tblAOPReturn.DefaultView.RowFilter = "[Primary] = " & row_Supplier.Primary_Transaksi_Asal.ToString
                    If tblAOPReturn.DefaultView.Count > 0 Then
                        rowAOPReturn = tblAOPReturn.DefaultView(0).Row
                        _NoTransaksi = rowAOPReturn.No_Transaksi
                        _TglTransaksi = rowAOPReturn.Tanggal_Transaksi
                        If IsNumeric(rowAOPReturn.Payment_Term) Then
                            _DueDate = rowAOPReturn.Tanggal_Transaksi.AddDays(rowAOPReturn.Payment_Term)
                        Else
                            _DueDate = rowAOPReturn.Tanggal_Transaksi
                        End If
                        ApakahTrans = True
                    End If
                ElseIf row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Then
                    tblAOPReturn2.DefaultView.RowFilter = "[Primary] = " & row_Supplier.Primary_Transaksi_Asal.ToString
                    If tblAOPReturn2.DefaultView.Count > 0 Then
                        rowAOPReturn2 = tblAOPReturn2.DefaultView(0).Row
                        _NoTransaksi = rowAOPReturn2.AOPR_Number
                        _TglTransaksi = rowAOPReturn2.AOPR_Date
                        _DueDate = rowAOPReturn2.Due_Date
                        _PrimaryAkunAP = rowAOPReturn2.primary_ap_account
                        ApakahTrans = True
                    End If
                ElseIf row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Or _
                row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                    If row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                        Filter = "Primary = " & row_Supplier.Primary_Transaksi_Asal.ToString & _
                                 " And [tipe transaksi] = " & AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                                 " And [tipe kartu] = " & AccModule.ClsTransaksiCreditNote.EnumTipeKartu.Supplier
                    ElseIf row_Supplier.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                        Filter = "Primary = " & row_Supplier.Primary_Transaksi_Asal.ToString & _
                                 " And [tipe transaksi] = " & AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                                 " And [tipe kartu] = " & AccModule.ClsTransaksiCreditNote.EnumTipeKartu.Supplier
                    End If
                    tblCD_Note.DefaultView.RowFilter = Nothing
                    tblCD_Note.DefaultView.RowFilter = Filter
                    If tblCD_Note.DefaultView.Count > 0 Then
                        rowCD_Note = tblCD_Note.DefaultView(0).Row
                        _NoTransaksi = rowCD_Note.no_transksi
                        _TglTransaksi = rowCD_Note.tgl
                        _DueDate = rowCD_Note.due_date
                        _PrimaryAkunAP = rowCD_Note.primary_akun_ar
                        ApakahTrans = True
                    End If
                End If

                If ApakahTrans = True Then
                    If row_Supplier.Ending_Balance > 0 Then
                        If frmTransaksiPurchaseBalanceManual_.FromCutOff_PurchBalanceManual(Me.UserCode, row_Supplier.Primary_Supplier, _
                        _NoTransaksi, _TglTransaksi, _DueDate, row_Supplier.Ending_Balance, _PrimaryAkunAP) = False Then
                            MsgBox("Failed to save purchase balance: " & row_Supplier.No_Transaksi_Asal, MsgBoxStyle.Critical, "Error")
                            Return False
                        End If
                    ElseIf row_Supplier.Ending_Balance < 0 Then
                        If frmTransaksiPurchaseReturnBalanceManual_.FromCutOff_PurchBalanceReturnManual(Me.UserCode, row_Supplier.Primary_Supplier, _
                        _NoTransaksi, _TglTransaksi, _DueDate, _PrimaryAkunAP, row_Supplier.Ending_Balance * -1) = False Then
                            MsgBox("Failed to save purchase return balance: " & row_Supplier.No_Transaksi_Asal, MsgBoxStyle.Critical, "Error")
                            Return False
                        End If
                    End If
                Else
                    If row_Supplier.No_Transaksi_Asal <> "" Then
                        If listPurchase.Contains(row_Supplier.No_Transaksi_Asal) = False Then
                            listPurchase.Add(row_Supplier.No_Transaksi_Asal)
                        End If
                    End If
                End If
            Next

            If listPurchase.Count > 0 Then
                Dim frmTrans_CutOffDB_BalanceException_ As New FrmTrans_CutOffDB_BalanceException
                If listPurchase.Count > 0 Then
                    frmTrans_CutOffDB_BalanceException_.ListBox1.Items.Add("The following transaction number doesn't exist:")
                    For Each i As String In listPurchase
                        frmTrans_CutOffDB_BalanceException_.ListBox1.Items.Add(i)
                    Next
                    frmTrans_CutOffDB_BalanceException_.ListBox1.Items.Add("")
                End If

                frmTrans_CutOffDB_BalanceException_.ShowDialog()
            End If

            Me.RowUpdate("Transfer Purchase & Purchase Return Begining Balance Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferAPBalance")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.Label1.Text = ""
            Me.ProgressBar1.Value = 0
            dtaAOP.Dispose()
            tblAOP.Dispose()
            dtaAOP2.Dispose()
            tblAOP2.Dispose()
            dtaAOPReturn.Dispose()
            tblAOPReturn.Dispose()
            dtaAOPReturn2.Dispose()
            tblAOPReturn2.Dispose()
            dtaCD_Note.Dispose()
            tblCD_Note.Dispose()
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Private Function TransferARBalance() As Boolean
        Dim frmLaporanARSummary_ As New AccModule.FrmLaporanARSummary
        Dim Tbl_Customer As New AccModule.DseLaporanARSummary.tblmaster_customerDataTable
        Dim row_Customer As AccModule.DseLaporanARSummary.tblmaster_customerRow

        'Invoice Otomatis
        Dim dtaInvoice As New Purchase_Module.DseTrans_CutOffDBTableAdapters.tbltrans_salesinvTableAdapter
        Dim tblInvoice As New Purchase_Module.DseTrans_CutOffDB.tbltrans_salesinvDataTable
        Dim rowInvoice As Purchase_Module.DseTrans_CutOffDB.tbltrans_salesinvRow
        'Invoice Manual
        Dim dtaInvoice2 As New AccModule.DseTransaksiSalesBalanceTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim tblInvoice2 As New AccModule.DseTransaksiSalesBalance.tbltrans_historicalsalesDataTable
        Dim rowInvoice2 As AccModule.DseTransaksiSalesBalance.tbltrans_historicalsalesRow
        'Invoice Return Otomatis
        Dim dtaInvoiceReturn As New Purchase_Module.DseTrans_CutOffDBTableAdapters.tbltrans_salesreturnTableAdapter
        Dim tblInvoiceReturn As New Purchase_Module.DseTrans_CutOffDB.tbltrans_salesreturnDataTable
        Dim rowInvoiceReturn As Purchase_Module.DseTrans_CutOffDB.tbltrans_salesreturnRow
        'Invoice Return Manual
        Dim dtaInvoiceReturn2 As New AccModule.DseTransaksiSalesReturnBalanceTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim tblInvoiceReturn2 As New AccModule.DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualDataTable
        Dim rowInvoiceReturn2 As AccModule.DseTransaksiSalesReturnBalance.tbltrans_historicalsalesreturn_manualRow
        'Credit-Debit Note
        Dim dtaCD_Note As New AccModule.DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim tblCD_Note As New AccModule.DseTransaksiCreditNote.tbltrans_credit_noteDataTable
        Dim rowCD_Note As AccModule.DseTransaksiCreditNote.tbltrans_credit_noteRow

        Dim frmTransaksiSalesBalanceManual_ As New AccModule.FrmTransaksiSalesBalanceManual
        Dim frmTransaksiSalesReturnBalanceManual_ As New AccModule.FrmTransaksiSalesReturnBalanceManual

        Dim dbName As String = Core_Module.GetDBName
        Dim rowIdx As Integer = 0
        Dim ApakahTrans As Boolean = False
        Dim Filter As String = ""
        Dim _NoTransaksi As String = ""
        Dim _TglTransaksi As Date = Now
        Dim _DueDate As Date = Now
        Dim _PrimaryAkunAR As Double = -1
        Dim BatchProgBar As Integer = 1
        Dim ListSales As New List(Of String)

        Try
            Me.RowUpdate("Transfer Sales & Sales Return Begining Balance Data", Me.StatusOnProses)
            Me.Label1.Text = "Get Sales And Sales Return Begining Balance Data, Please Wait... "

            Me.SetDefaultConnectionString()
            dtaInvoice.Connection.ConnectionString = Me.SumberConstring
            dtaInvoiceReturn.Connection.ConnectionString = Me.SumberConstring
            dtaInvoice.Fill(tblInvoice, Me.EndDateCutOff)
            dtaInvoice2.Fill(tblInvoice2)
            dtaInvoiceReturn.Fill(tblInvoiceReturn, Me.EndDateCutOff)
            dtaInvoiceReturn2.Fill(tblInvoiceReturn2)
            dtaCD_Note.Fill(tblCD_Note)
            frmLaporanARSummary_.FromCutOff_ARSummary(Tbl_Customer, Me.EndDateCutOff.Date, Me.EndDateCutOff.Date)
            Me.SetTargetConnectionString()

            'Transfer Sales Begining Balance Data
            Me.SetProgressBarNLabel("", BatchProgBar, True, Tbl_Customer.Count)
            For Each row_Customer In Tbl_Customer.Rows
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Sales And Sales Return Begining Balance Data", BatchProgBar)

                ApakahTrans = False
                If row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
                    tblInvoice.DefaultView.RowFilter = "[Primary] = " & row_Customer.Primary_Transaksi_Asal.ToString
                    If tblInvoice.DefaultView.Count > 0 Then
                        rowInvoice = tblInvoice.DefaultView(0).Row
                        _NoTransaksi = rowInvoice.No_Transaksi
                        _TglTransaksi = rowInvoice.Tanggal_Transaksi
                        If IsNumeric(rowInvoice.Payment_Term) Then
                            _DueDate = rowInvoice.Tanggal_Transaksi.AddDays(rowInvoice.Payment_Term)
                        Else
                            _DueDate = rowInvoice.Tanggal_Transaksi
                        End If
                        _PrimaryAkunAR = rowInvoice.primary_akun_ar
                        ApakahTrans = True
                    End If
                ElseIf row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual Then
                    tblInvoice2.DefaultView.RowFilter = "[Primary] = " & row_Customer.Primary_Transaksi_Asal.ToString
                    If tblInvoice2.DefaultView.Count > 0 Then
                        rowInvoice2 = tblInvoice2.DefaultView(0).Row
                        _NoTransaksi = rowInvoice2.Invoice_Number
                        _TglTransaksi = rowInvoice2.Invoice_Date
                        _DueDate = rowInvoice2.Due_Date
                        _PrimaryAkunAR = rowInvoice2.primary_ar_account
                        ApakahTrans = True
                    End If
                ElseIf row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
                    tblInvoiceReturn.DefaultView.RowFilter = "[Primary] = " & row_Customer.Primary_Transaksi_Asal.ToString
                    If tblInvoiceReturn.DefaultView.Count > 0 Then
                        rowInvoiceReturn = tblInvoiceReturn.DefaultView(0).Row
                        _NoTransaksi = rowInvoiceReturn.No_Transaksi
                        _TglTransaksi = rowInvoiceReturn.Tanggal_Transaksi
                        If IsNumeric(rowInvoiceReturn.Payment_Term) Then
                            _DueDate = rowInvoiceReturn.Tanggal_Transaksi.AddDays(rowInvoiceReturn.Payment_Term)
                        Else
                            _DueDate = rowInvoiceReturn.Tanggal_Transaksi
                        End If
                        _PrimaryAkunAR = rowInvoiceReturn.primary_ar_account
                        ApakahTrans = True
                    End If
                ElseIf row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Then
                    tblInvoiceReturn2.DefaultView.RowFilter = "[Primary] = " & row_Customer.Primary_Transaksi_Asal.ToString
                    If tblInvoiceReturn2.DefaultView.Count > 0 Then
                        rowInvoiceReturn2 = tblInvoiceReturn2.DefaultView(0).Row
                        _NoTransaksi = rowInvoiceReturn2.AOSR_Number
                        _TglTransaksi = rowInvoiceReturn2.AOSR_Date
                        _DueDate = rowInvoiceReturn2.Due_Date
                        _PrimaryAkunAR = rowInvoiceReturn2.primary_ar_account
                        ApakahTrans = True
                    End If
                ElseIf row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Or _
                row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                    If row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                        Filter = "Primary = " & row_Customer.Primary_Transaksi_Asal.ToString & _
                                 " And [tipe transaksi] = " & AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                                 " And [tipe kartu] = " & AccModule.ClsTransaksiCreditNote.EnumTipeKartu.Customer
                    ElseIf row_Customer.Tipe_Transaksi_Asal = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                        Filter = "Primary = " & row_Customer.Primary_Transaksi_Asal.ToString & _
                                 " And [tipe transaksi] = " & AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                                 " And [tipe kartu] = " & AccModule.ClsTransaksiCreditNote.EnumTipeKartu.Customer
                    End If
                    tblCD_Note.DefaultView.RowFilter = Nothing
                    tblCD_Note.DefaultView.RowFilter = Filter
                    If tblCD_Note.DefaultView.Count > 0 Then
                        rowCD_Note = tblCD_Note.DefaultView(0).Row
                        _NoTransaksi = rowCD_Note.no_transksi
                        _TglTransaksi = rowCD_Note.tgl
                        _DueDate = rowCD_Note.due_date
                        _PrimaryAkunAR = rowCD_Note.primary_akun_ar
                        ApakahTrans = True
                    End If
                End If

                If ApakahTrans = True Then
                    If row_Customer.Ending_Balance > 0 Then
                        If frmTransaksiSalesBalanceManual_.FromCutOff_SalesBalanceManual(Me.UserCode, row_Customer.Primary, _NoTransaksi, _
                        _TglTransaksi, _DueDate, row_Customer.Ending_Balance, _PrimaryAkunAR) = False Then
                            MsgBox("Failed to save sales balance: " & row_Customer.No_Transaksi_Asal, MsgBoxStyle.Critical, "Error")
                            Return False
                        End If
                    ElseIf row_Customer.Ending_Balance < 0 Then
                        If frmTransaksiSalesReturnBalanceManual_.FromCutOff_SalesBalanceReturnManual(Me.UserCode, row_Customer.Primary, _NoTransaksi, _
                       _TglTransaksi, _DueDate, row_Customer.Ending_Balance * -1, _PrimaryAkunAR) = False Then
                            MsgBox("Failed to save sales balance: " & row_Customer.No_Transaksi_Asal, MsgBoxStyle.Critical, "Error")
                            Return False
                        End If
                    End If
                Else
                    If row_Customer.No_Transaksi_Asal <> "" Then
                        If ListSales.Contains(row_Customer.No_Transaksi_Asal) = False Then
                            ListSales.Add(row_Customer.No_Transaksi_Asal)
                        End If
                    End If
                End If
            Next

            If ListSales.Count > 0 Then
                Dim frmTrans_CutOffDB_BalanceException_ As New FrmTrans_CutOffDB_BalanceException
                If ListSales.Count > 0 Then
                    frmTrans_CutOffDB_BalanceException_.ListBox1.Items.Add("The following invoice number doesn't exist:")
                    For Each i As String In ListSales
                        frmTrans_CutOffDB_BalanceException_.ListBox1.Items.Add(i)
                    Next
                    frmTrans_CutOffDB_BalanceException_.ListBox1.Items.Add("")
                End If

                frmTrans_CutOffDB_BalanceException_.ShowDialog()
            End If
            Me.RowUpdate("Transfer Sales & Sales Return Begining Balance Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferARBalance")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.Label1.Text = ""
            Me.ProgressBar1.Value = 0
            Tbl_Customer.Dispose()
            dtaInvoice.Dispose()
            tblInvoice.Dispose()
            dtaInvoice2.Dispose()
            tblInvoice2.Dispose()
            dtaInvoiceReturn.Dispose()
            tblInvoiceReturn.Dispose()
            dtaInvoiceReturn2.Dispose()
            tblInvoiceReturn2.Dispose()
            dtaCD_Note.Dispose()
            tblCD_Note.Dispose()
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Private Function TransferBeginingBalance() As Boolean
        Dim frmLaporanStockDetail2_ As New Inventory_Module.FrmLaporanStockDetail2
        Dim frmTransaksiStockInAdj_ As New Inventory_Module.FrmTransaksiStockInAdj
        Dim listOfProduct As New List(Of Double)

        Dim dtaAmount As New DseTrans_CutOffDBTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim tblAmount As New DseTrans_CutOffDB.tbltrans_purchase_closing_item_detailDataTable
        Dim TBLCetak As New Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable
        Dim rowQty As Inventory_Module.DseLaporanStockDetail2.TblCetakRow

        Dim costingTipe As Integer = -1
        Dim totalBefore As Decimal = 0
        Dim MaxRow As Integer = 500000
        Dim Count As Integer = 0
        Dim BuatBaru, SimpanInAdj As Boolean
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Product Qty Begining Balance Data", Me.StatusOnProses)
            Me.Label1.Text = "Get Product Qty Begining Balance Data, Please Wait... "
            Me.SetDefaultConnectionString()
            costingTipe = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
            frmLaporanStockDetail2_.FromCutOff_BeginingBalance(TBLCetak, Me.EndDateCutOff)
            dtaAmount.Connection.ConnectionString = Me.SumberConstring
            dtaAmount.Fill(tblAmount, Me.ObjCD.GetFirstDateofTheMonth(Me.EndDateCutOff), Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem)

            If Not TBLCetak.Columns.Contains("Amount Final") Then
                TBLCetak.Columns.Add("Amount Final", System.Type.GetType("System.Decimal")).DefaultValue = 0
            End If

            ''isi amount, ambil dari purchase_closing_item_detail
            For Each rowQty In TBLCetak.Rows
                If Not listOfProduct.Contains(rowQty.Primary_Product) Then
                    listOfProduct.Add(rowQty.Primary_Product)
                End If
            Next

            For Each primaryProduct As Double In listOfProduct
                TBLCetak.DefaultView.RowFilter = "[Primary_Product] = " & primaryProduct.ToString
                For i As Integer = 0 To TBLCetak.DefaultView.Count - 1
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    rowQty = TBLCetak.DefaultView(i).Row
                    If i < TBLCetak.DefaultView.Count - 1 Then
                        If costingTipe = 0 Then
                            rowQty.Item("Amount Final") = Me.ObjCD.DigitGrouping(Me.ObjCutOff.TransferBeginingBalance_GetAmount(tblAmount, rowQty.Primary_Product, rowQty.QtyDec, ClsTrans_CutOffDB.EnumClosingMethod.Average), 0)
                        ElseIf costingTipe = 1 Then
                            rowQty.Item("Amount Final") = Me.ObjCD.DigitGrouping(Me.ObjCutOff.TransferBeginingBalance_GetAmount(tblAmount, rowQty.Primary_Product, rowQty.QtyDec, ClsTrans_CutOffDB.EnumClosingMethod.FIFO), 0)
                        End If
                        totalBefore += rowQty.Item("Amount Final")
                    Else
                        If costingTipe = 0 Then
                            rowQty.Item("Amount Final") = Me.ObjCD.DigitGrouping(Me.ObjCutOff.TransferBeginingBalance_GetTotalAmount(tblAmount, rowQty.Primary_Product, ClsTrans_CutOffDB.EnumClosingMethod.Average) - totalBefore, 0)
                        ElseIf costingTipe = 1 Then
                            rowQty.Item("Amount Final") = Me.ObjCD.DigitGrouping(Me.ObjCutOff.TransferBeginingBalance_GetTotalAmount(tblAmount, rowQty.Primary_Product, ClsTrans_CutOffDB.EnumClosingMethod.FIFO) - totalBefore, 0)
                        End If
                        totalBefore = 0
                    End If
                Next
            Next

            ''input pada inventory begining balance
            If TBLCetak.Rows.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TBLCetak.Count)
                Me.SetTargetConnectionString()

                Dim rowIdx As Integer = 0
                Me.TblListTransfer.DefaultView.RowFilter = "[Description] = 'Transfer Product Qty Begining Balance Data'"
                If Me.TblListTransfer.DefaultView.Count > 0 Then
                    Me.RowListTransfer = Me.TblListTransfer.DefaultView(0).Row
                    rowIdx = Me.RowListTransfer.No
                End If

                Me.SetProgressBarNLabel("Transfer Product Qty Begining Balance Data", BatchProgBar, True, TBLCetak.Rows.Count)
                For i As Integer = 0 To TBLCetak.Rows.Count - 1
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Product Qty Begining Balance Data", BatchProgBar)

                    rowQty = TBLCetak.Rows(i)
                    'karena tidak boleh ada product yang amountnya <= 0 maka. beri peringatan dan di simpan nilai amountnya 1 dulu
                    If rowQty.Item("Amount Final") <= 0 Then
                        rowQty.Item("Amount Final") = 1
                    End If
                    BuatBaru = False
                    SimpanInAdj = False
                    Count += 1
                    If Count = 1 Then
                        BuatBaru = True
                    ElseIf Count = MaxRow Or i = TBLCetak.Rows.Count - 1 Then
                        SimpanInAdj = True
                        Count = 0
                    End If

                    frmTransaksiStockInAdj_.FromCutOff_BeginningBalance(Me.UserCode, rowQty, Me.EndDateCutOff, BuatBaru, SimpanInAdj)
                Next
                Me.RowUpdate("Transfer Product Qty Begining Balance Data", Me.StatusOK)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferBeginingBalance")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.Label1.Text = ""
            Me.ProgressBar1.Value = 0
            dtaAmount.Dispose()
            tblAmount.Dispose()
            TBLCetak.Dispose()
            Me.SetDefaultConnectionString()
        End Try
    End Function

    'TRANSFER PURCHASE MODULE
    Public Function TransferPOB(ByRef TblPO As DseTrans_CutOffDB.tbltrans_poDataTable, _
    ByRef TblPO_Item As DseTrans_CutOffDB.tbltrans_po_itemDataTable, _
    ByRef TblPO_ItemDel As DseTrans_CutOffDB.tbltrans_po_itemdeliveryDataTable) As Boolean
        Dim vTblDec_PO As New DseTrans_CutOffDB.tblmaster_productdecimalDataTable
        Dim rowPO As DseTrans_CutOffDB.tbltrans_poRow
        Dim rowPO_Item As DseTrans_CutOffDB.tbltrans_po_itemRow
        Dim rowPO_ItemDel As DseTrans_CutOffDB.tbltrans_po_itemdeliveryRow
        Dim DtaPL_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim TblPL_Item_OutSumQty, TblPL_Item_OutSumQty_Filter As New DseTrans_CutOffDB.tbltrans_packinglist_itemDataTable
        Dim DtaPOC_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_poc_itemTableAdapter
        Dim TblPOC_Item_OutSumQty, TblPOC_Item_OutSumQty_Filter As New DseTrans_CutOffDB.tbltrans_poc_itemDataTable
        Dim DtaSIO_Item As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TblSIP_Item_OutSumQty, TblSIP_Item_OutSumQty_Filter As New DseTrans_CutOffDB.tbltransaksi_stockinout_itemDataTable
        Dim DtaPRet_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_purchasereturn_itemTableAdapter
        Dim TblPRet_Item_OutSumQty, TblPRet_Item_OutSumQty_Filter As New DseTrans_CutOffDB.tbltrans_purchasereturn_itemDataTable

        Dim AddPO_ As Boolean = False
        Dim rowIdx As Integer = 0
        Dim DefPrec As Integer = 0
        Dim TransPOItem As Boolean = True
        Dim PrimaryPO_Max As Double = -1
        Dim PrimaryBatch As Double = 500
        Dim BatchProgBar As Integer = 1

        Try
            'Cari purchase order dalam period cut off yang masih mempunyai balance
            Me.RowUpdate("Transfer Purchase Order Balance Data", Me.StatusOnProses)
            DtaPL_Item.Connection.ConnectionString = Me.SumberConstring
            DtaPRet_Item.Connection.ConnectionString = Me.SumberConstring
            DtaSIO_Item.Connection.ConnectionString = Me.SumberConstring
            DtaPOC_Item.Connection.ConnectionString = Me.SumberConstring
            DtaPL_Item.FillBy_OutCutOffSumQty(TblPL_Item_OutSumQty, Me.EndDateCutOff)
            DtaPRet_Item.FillBy_QutCutOffSumQty(TblPRet_Item_OutSumQty, Me.EndDateCutOff)
            DtaPOC_Item.FillBy_OutCutOffSumQty(TblPOC_Item_OutSumQty, Me.EndDateCutOff)
            DtaSIO_Item.FillBy_OutCutOffSumQty(TblSIP_Item_OutSumQty, Me.EndDateCutOff, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase)

            TblPO.DefaultView.RowFilter = "[Primary Quick Acuan] = -1 AND [Tanggal Transaksi] <= #" & Me.EndDateCutOff & "#"
            TblPO.DefaultView.Sort = "[Primary]"
            If TblPO.DefaultView.Count > 0 Then
                vTblDec_PO = TblDec.Copy
                Me.SetProgressBarNLabel("Transfer Purchase Order Balance Data", BatchProgBar, True, TblPO.DefaultView.Count)
                For i As Integer = TblPO.DefaultView.Count - 1 To 0 Step -1
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    AddPO_ = False
                    rowPO = TblPO.DefaultView(i).Row
                    Me.SetProgressBarNLabel("Transfer Purchase Order Balance Data", BatchProgBar)

                    If PrimaryPO_Max < rowPO.Primary Then
                        PrimaryPO_Max = rowPO.Primary + PrimaryBatch
                        Me.ObjCutOff.GetTblPOB_Filter(TblPL_Item_OutSumQty, TblPL_Item_OutSumQty_Filter, TblPOC_Item_OutSumQty, _
                        TblPOC_Item_OutSumQty_Filter, TblPRet_Item_OutSumQty, TblPRet_Item_OutSumQty_Filter, _
                        TblSIP_Item_OutSumQty, TblSIP_Item_OutSumQty_Filter, rowPO.Primary, PrimaryPO_Max)
                    End If

                    TblPO_Item.DefaultView.RowFilter = "[Primary Transaksi] = " & rowPO.Primary
                    If TblPO_Item.DefaultView.Count > 0 Then
                        For i2 As Integer = TblPO_Item.DefaultView.Count - 1 To 0 Step -1
                            rowPO_Item = TblPO_Item.DefaultView(i2).Row
                            'cek apakah item dalam po masih membunyai balance, jika tidak hapus baris itemnya
                            'jika semua item dalam po tidak mempunyai balance maka hapus transaksi po nya.
                            TransPOItem = Me.ObjCutOff.GetBalancePO(TblPL_Item_OutSumQty_Filter, TblPOC_Item_OutSumQty_Filter, _
                            TblSIP_Item_OutSumQty_Filter, TblPRet_Item_OutSumQty_Filter, rowPO_Item, vTblDec_PO)
                            If TransPOItem = False Then
DelPO_ItemDel:
                                TblPO_ItemDel.DefaultView.RowFilter = "[identifier] = '" & rowPO_Item.identifier & "'"
                                If TblPO_ItemDel.DefaultView.Count > 0 Then
                                    For Each dvDel As DataRowView In TblPO_ItemDel.DefaultView
                                        rowPO_ItemDel = dvDel.Row
                                        TblPO_ItemDel.Rows.Remove(rowPO_ItemDel)
                                    Next
                                End If
                                If TransPOItem = False Then
                                    TblPO_Item.Rows.Remove(rowPO_Item)
                                End If
                            Else
                                AddPO_ = True
                                If rowPO_Item.Qty_Multi = 0 Then
                                    GoTo DelPO_ItemDel
                                End If
                            End If
                        Next

                        If AddPO_ = False Then
                            TblPO.Rows.Remove(rowPO)
                        End If
                    End If
                Next
            End If

            Me.RowUpdate("Transfer Purchase Order Balance Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferPOB")
            Return False
        Finally
            Me.Label1.Text = ""
            Me.ProgressBar1.Value = 0
            DtaPL_Item.Dispose()
            TblPL_Item_OutSumQty.Dispose()
            TblPL_Item_OutSumQty_Filter.Dispose()
            DtaPOC_Item.Dispose()
            TblPOC_Item_OutSumQty.Dispose()
            TblPOC_Item_OutSumQty_Filter.Dispose()
            DtaSIO_Item.Dispose()
            TblSIP_Item_OutSumQty.Dispose()
            TblSIP_Item_OutSumQty_Filter.Dispose()
            DtaPRet_Item.Dispose()
            TblPRet_Item_OutSumQty.Dispose()
            TblPRet_Item_OutSumQty_Filter.Dispose()
        End Try
    End Function

    Public Function TransferPO() As Boolean
        Dim DtaPO As New DseTrans_CutOffDBTableAdapters.tbltrans_poTableAdapter
        Dim TblPO As New DseTrans_CutOffDB.tbltrans_poDataTable
        Dim DtaPO_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_po_itemTableAdapter
        Dim TblPO_Item, TblPO_Item_Filter As New DseTrans_CutOffDB.tbltrans_po_itemDataTable
        Dim DtaPO_ItemDel As New DseTrans_CutOffDBTableAdapters.tbltrans_po_itemdeliveryTableAdapter
        Dim TblPO_ItemDel, TblPO_ItemDel_Filter As New DseTrans_CutOffDB.tbltrans_po_itemdeliveryDataTable
        Dim rowPO As DseTrans_CutOffDB.tbltrans_poRow
        Dim FrmPO As New FrmTransaksiPO

        Dim PrimaryPO_Max As Double = -1
        Dim PrimaryBatch As Double = -1
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Purchase Order Data", Me.StatusOnProses)
            Me.Label1.Text = "Get Purchase Order Balance Data, Please Wait..."

            DtaPO.Connection.ConnectionString = Me.SumberConstring
            DtaPO_Item.Connection.ConnectionString = Me.SumberConstring
            DtaPO_ItemDel.Connection.ConnectionString = Me.SumberConstring
            DtaPO.Fill_All(TblPO)
            DtaPO_Item.FillBy_PeriodCutOff(TblPO_Item, Me.EndDateCutOff)
            DtaPO_ItemDel.FillBy_PeriodCutOff(TblPO_ItemDel, Me.EndDateCutOff)

            'purchase order balance
            Me.TransferPOB(TblPO, TblPO_Item, TblPO_ItemDel)

            DtaPO_Item.Connection.ConnectionString = Me.SumberConstring
            DtaPO_ItemDel.Connection.ConnectionString = Me.SumberConstring
            DtaPO_Item.ClearBeforeFill = False
            DtaPO_ItemDel.ClearBeforeFill = False
            DtaPO_Item.FillBy_OutCutOff(TblPO_Item, Me.EndDateCutOff)
            DtaPO_ItemDel.FillBy_OutCutOff(TblPO_ItemDel, Me.EndDateCutOff)

            Me.SetTargetConnectionString()

            ''Transfer PO period cutoff yang masih mempunyai sisa balance dan transfer semua PO di luar period cutoff
            TblPO.DefaultView.RowFilter = "[Primary Quick Acuan] = -1"
            TblPO.DefaultView.Sort = "[Primary]"
            If TblPO.DefaultView.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblPO.DefaultView.Count)
                For Each dv As DataRowView In TblPO.DefaultView
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    rowPO = dv.Row
                    FrmPO.CutOff_ForBgw = False
                    Me.SetProgressBarNLabel("Transfer Purchase Order Data", BatchProgBar)

                    If PrimaryPO_Max < rowPO.Primary Then
                        PrimaryPO_Max = rowPO.Primary + PrimaryBatch
                        Me.ObjCutOff.GetTblPO_Filter(TblPO_Item, TblPO_Item_Filter, _
                        TblPO_ItemDel, TblPO_ItemDel_Filter, rowPO.Primary, PrimaryPO_Max)
                    End If

                    PrimaryUpdate = FrmPO.FromCutOff_PO(rowPO, TblPO_Item, TblPO_ItemDel)
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowPO.Primary, PrimaryUpdate, Me.TblPrimary_PO)
                Next
            End If
            Me.RowUpdate("Transfer Purchase Order Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferPO")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.Label1.Text = ""
            Me.ProgressBar1.Value = 0
            Me.SetDefaultConnectionString()
            FrmPO.Dispose()
            TblPO.Dispose()
            TblPO_Item.Dispose()
            TblPO_ItemDel.Dispose()
        End Try
    End Function

    Public Function TransferPackingList() As Boolean
        Dim DtaPL As New DseTrans_CutOffDBTableAdapters.tbltrans_packinglistTableAdapter
        Dim TblPL_Out As New DseTrans_CutOffDB.tbltrans_packinglistDataTable
        Dim rowPL As DseTrans_CutOffDB.tbltrans_packinglistRow
        Dim FrmPL As New FrmTransaksiPackingList

        Dim BatchProgBar As Integer = 1
        Dim PrimaryUpdate As Double = -1

        Try
            Me.RowUpdate("Transfer Packing List Data", Me.StatusOnProses)
            DtaPL.Connection.ConnectionString = Me.SumberConstring
            DtaPL.FillBy_OutCutOff(TblPL_Out, Me.EndDateCutOff)

            ''Transfer semua Packing list di luar period cutoff
            Me.SetProgressBarNLabel("", BatchProgBar, True, TblPL_Out.Count)
            For Each rowPL In TblPL_Out.Rows
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Packing List Data", BatchProgBar)

                'Get Primary PO New
                Me.SetDefaultConnectionString()
                FrmPL.FromCutOff_PackingList_Create(rowPL.Primary, Me.TblPrimary_PO, Me.ObjCutOff_Core)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmPL.FromCutOff_PackingList_Save()
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowPL.Primary, PrimaryUpdate, Me.TblPrimary_PL)
            Next
            Me.RowUpdate("Transfer Packing List Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferPackingList")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmPL.Dispose()
            DtaPL.Dispose()
            TblPL_Out.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferStockInPurch() As Boolean
        Dim DtaSIO As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TblSIP_Out As New DseTrans_CutOffDB.tbltransaksi_stockinoutDataTable
        Dim rowSIP As DseTrans_CutOffDB.tbltransaksi_stockinoutRow
        Dim DtaGrn_Old, DtaGrn_New As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbltrans_grnTableAdapter
        Dim TblGrn_Old, TblGrn_New As New Core_Module.DseTrans_CutOffDB_Core.tbltrans_grnDataTable
        Dim row As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateRow
        Dim FrmSIP As New Inventory_Module.FrmTransaksiStockInPurchase2

        Dim PrimaryOld As Double = -1
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock In Purchase Data", Me.StatusOnProses)
            DtaSIO.Connection.ConnectionString = Me.SumberConstring
            DtaSIO.FillBy_OutCutOff(TblSIP_Out, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, Me.EndDateCutOff)

            ''Transfer semua stock in purchase di luar period cutoff
            TblSIP_Out.DefaultView.RowFilter = "[Primary Quick Acuan] = -1"
            If TblSIP_Out.DefaultView.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblSIP_Out.DefaultView.Count)
                For Each dv As DataRowView In TblSIP_Out.DefaultView
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    rowSIP = dv.Row
                    Me.SetProgressBarNLabel("Transfer Stock In Purchase Data", BatchProgBar)

                    'Get Primary packing List New & Primary PO New
                    Me.SetDefaultConnectionString()
                    FrmSIP.FromCutOff_StockInPurch_Create(rowSIP.Primary, Me.TblPrimary_PO, Me.TblPrimary_PL, Me.ObjCutOff_Core)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmSIP.FromCutOff_StockInPurch_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowSIP.Primary, PrimaryUpdate, Me.TblPrimary_SIP)
                Next

                Me.SetDefaultConnectionString()
                DtaGrn_Old.FillBy_OutCutOff(TblGrn_Old, Me.EndDateCutOff)
                Me.SetTargetConnectionString()
                DtaGrn_New.FillBy_OutCutOff(TblGrn_New, Me.EndDateCutOff)

                For Each row In Me.TblPrimary_SIP.Rows
                    Me.SetProgressBarNLabel("Update GRN Data", BatchProgBar)

                    PrimaryOld = Me.ObjCutOff_Core.GetPrimaryGrn(row.Primary_Old, TblGrn_Old)
                    PrimaryUpdate = Me.ObjCutOff_Core.GetPrimaryGrn(row.Primary_New, TblGrn_New)
                    Me.ObjCutOff_Core.AddPrimaryUpdate(PrimaryOld, PrimaryUpdate, Me.TblPrimary_GRN)
                Next
            End If
            Me.RowUpdate("Transfer Stock In Purchase Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockInPurch")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmSIP.Dispose()
            DtaSIO.Dispose()
            TblSIP_Out.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferPurchReturn() As Boolean
        Dim DtaPRet As New DseTrans_CutOffDBTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim TblPRet_Out As New DseTrans_CutOffDB.tbltrans_purchasereturn1DataTable
        Dim rowPR As DseTrans_CutOffDB.tbltrans_purchasereturn1Row
        Dim FrmPR As New FrmTransaksiPurchaseReturn

        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Purchase Return Data", Me.StatusOnProses)
            DtaPRet.Connection.ConnectionString = Me.SumberConstring
            DtaPRet.FillBy_OutCutOff(TblPRet_Out, Me.EndDateCutOff)

            TblPRet_Out.DefaultView.RowFilter = "[Quick PR] = 0"
            Me.SetProgressBarNLabel("", BatchProgBar, True, TblPRet_Out.DefaultView.Count)
            For Each dv As DataRowView In TblPRet_Out.DefaultView
                rowPR = dv.Row
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Purchase Return Data", BatchProgBar)

                Me.SetDefaultConnectionString()
                FrmPR.FromCutOff_PReturn_Create(rowPR.Primary, rowPR.primary_supplier, Me.TblPrimary_GRN, Me.TblPrimary_PO, Me.ObjCutOff_Core)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmPR.FromCutOff_PReturn_Save()
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowPR.Primary, PrimaryUpdate, Me.TblPrimary_PR)
            Next
            Me.RowUpdate("Transfer Purchase Return Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferPurchReturn")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            DtaPRet.Dispose()
            TblPRet_Out.Dispose()
            FrmPR.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferStockOutPurchReturn() As Boolean
        Dim DtaSIO As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TblSOPR_Out As New DseTrans_CutOffDB.tbltransaksi_stockinoutDataTable
        Dim rowSOPR As DseTrans_CutOffDB.tbltransaksi_stockinoutRow
        Dim FrmSOPR As New Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock Out Purchase Return Data", Me.StatusOnProses)
            DtaSIO.Connection.ConnectionString = Me.SumberConstring
            DtaSIO.FillBy_OutCutOff(TblSOPR_Out, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn, Me.EndDateCutOff)

            TblSOPR_Out.DefaultView.RowFilter = "[Primary Quick Acuan] = -1"
            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSOPR_Out.DefaultView.Count)
            For Each dv As DataRowView In TblSOPR_Out.DefaultView
                rowSOPR = dv.Row
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Stock Out Purchase Return Data", BatchProgBar)

                Me.SetDefaultConnectionString()
                FrmSOPR.FromCutOff_SOPR_Create(rowSOPR.Primary, Me.TblPrimary_PR, Me.TblPrimary_PO, Me.ObjCutOff_Core)
                Me.SetTargetConnectionString()
                FrmSOPR.FromCutOff_SOPR_Save()
            Next
            Me.RowUpdate("Transfer Stock Out Purchase Return Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockOutPurchReturn")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            DtaSIO.Dispose()
            TblSOPR_Out.Dispose()
            FrmSOPR.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferQuickPurchReturn() As Boolean
        Dim DtaPRet As New DseTrans_CutOffDBTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim TblPRet_Out As New DseTrans_CutOffDB.tbltrans_purchasereturn1DataTable
        Dim rowPR_Q As DseTrans_CutOffDB.tbltrans_purchasereturn1Row
        Dim Dta_SOPRNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_stockinout_noTableAdapter
        Dim Tbl_SOPRNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noDataTable

        Dim FrmPRQuick As New FrmTransaksiPurchaseReturn_Quick
        Dim NoStockOutPR As String = ""
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Quick Purchase Return Data", Me.StatusOnProses)
            Me.SetDefaultConnectionString()
            DtaPRet.Connection.ConnectionString = Me.SumberConstring
            DtaPRet.FillBy_OutCutOff(TblPRet_Out, Me.EndDateCutOff)
            Dta_SOPRNo.Fill(Tbl_SOPRNo, Me.EndDateCutOff, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn)

            TblPRet_Out.DefaultView.RowFilter = "[Quick PR] = 1"
            Me.SetProgressBarNLabel("", BatchProgBar, True, TblPRet_Out.DefaultView.Count)
            For Each dv As DataRowView In TblPRet_Out.DefaultView
                rowPR_Q = dv.Row
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Quick Purchase Return Data", BatchProgBar)

                NoStockOutPR = Me.ObjCutOff_Core.GetNoInOut(Tbl_SOPRNo, rowPR_Q.Primary)
                Me.SetDefaultConnectionString()
                FrmPRQuick.FromCutOff_PReturn_Quick_Create(rowPR_Q.Primary, rowPR_Q.primary_supplier)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmPRQuick.FromCutOff_PReturn_Quick_Save(NoStockOutPR)
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowPR_Q.Primary, PrimaryUpdate, Me.TblPrimary_PR)
            Next
            Me.RowUpdate("Transfer Quick Purchase Return Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferQuickPurchReturn")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            DtaPRet.Dispose()
            TblPRet_Out.Dispose()
            Dta_SOPRNo.Dispose()
            Tbl_SOPRNo.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferAOP_DP(ByVal TblAOP_Out As DseTrans_CutOffDB.tbltrans_aop1DataTable) As Boolean
        Dim DtaAOP As New DseTrans_CutOffDBTableAdapters.tbltrans_aop1TableAdapter
        Dim TblAOP_DP_Out As New DseTrans_CutOffDB.tbltrans_aop1DataTable
        Dim rowAOP_DP As DseTrans_CutOffDB.tbltrans_aop1Row
        Dim DtaAOP_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_aop_itemTableAdapter
        Dim TblAOP_Item_Out As New DseTrans_CutOffDB.tbltrans_aop_itemDataTable
        Dim DtaAOP_ItemDP As New DseTrans_CutOffDBTableAdapters.tbltrans_aop_itemdpTableAdapter
        Dim TblAOP_ItemDP_Out As New DseTrans_CutOffDB.tbltrans_aop_itemdpDataTable
        Dim FrmAOP_DP As New FrmTransaksiAOPDP

        Dim TransJurnal As Boolean = False
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Downpayment Approval Of Purchase Data", Me.StatusOnProses)
            DtaAOP.Connection.ConnectionString = Me.SumberConstring
            DtaAOP_Item.Connection.ConnectionString = Me.SumberConstring
            DtaAOP.Fill_All(TblAOP_DP_Out, ClsTransaksiAOP.EnumTipeAop.AOP_DP)
            DtaAOP_Item.FillBy_OutCutOff(TblAOP_Item_Out, Me.EndDateCutOff)
            DtaAOP_ItemDP.FillBy_OutCutOff(TblAOP_ItemDP_Out, Me.EndDateCutOff)

            'AOP DP yang di transfer adalah semua AOP DP yang berada di luar periode cut off dan
            'AOP DP yang ada di periode cut off apabila digunakan oleh transaksi AOP yang berada di luar periode cut off.
            'Me.TblAOP_DP.DefaultView.RowFilter = Nothing
            TblAOP_DP_Out.DefaultView.RowFilter = "[Tanggal Transaksi] <= #" & Me.EndDateCutOff & "#"
            For Each dv As DataRowView In TblAOP_DP_Out.DefaultView
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                rowAOP_DP = dv.Row
                TblAOP_ItemDP_Out.DefaultView.RowFilter = "[Primary AOP DP] = " & rowAOP_DP.Primary
                If TblAOP_ItemDP_Out.DefaultView.Count = 0 Then
                    TblAOP_DP_Out.Rows.Remove(rowAOP_DP)
                Else
                    TblAOP_Out.DefaultView.RowFilter = "[Primary] = " & rowAOP_DP.Primary.ToString & " AND [Tanggal Transaksi] <= #" & Me.EndDateCutOff & "#"
                    If TblAOP_Out.DefaultView.Count = 0 Then
                        TblAOP_DP_Out.Rows.Remove(rowAOP_DP)
                    End If
                End If
            Next

            TblAOP_DP_Out.DefaultView.RowFilter = Nothing
            Me.SetProgressBarNLabel("Transfer Downpayment Approval Of Purchase Data", BatchProgBar, True, TblAOP_DP_Out.Count)
            For Each rowAOP_DP In TblAOP_DP_Out.Rows
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Downpayment Approval Of Purchase Data", BatchProgBar)

                'Get Primary PO New
                If rowAOP_DP.Tanggal_Transaksi <= Me.EndDateCutOff Then
                    TransJurnal = False
                Else
                    TransJurnal = True
                End If

                Me.SetDefaultConnectionString()
                FrmAOP_DP.FromCutOff_AOPDP_Create(rowAOP_DP.Primary, Me.TblPrimary_PO, Me.ObjCutOff_Core)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmAOP_DP.FromCutOff_AOPDP_Save(TransJurnal)
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowAOP_DP.Primary, PrimaryUpdate, Me.TblPrimary_AOP)
            Next

            Me.RowUpdate("Transfer Downpayment Approval Of Purchase Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferAOP_DP")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmAOP_DP.Dispose()
            DtaAOP.Dispose()
            TblAOP_DP_Out.Dispose()
            DtaAOP_Item.Dispose()
            TblAOP_Item_Out.Dispose()
            DtaAOP_ItemDP.Dispose()
            TblAOP_ItemDP_Out.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferAOP() As Boolean
        Dim DtaAOP As New DseTrans_CutOffDBTableAdapters.tbltrans_aop1TableAdapter
        Dim TblAOP_Out As New DseTrans_CutOffDB.tbltrans_aop1DataTable
        Dim rowAOP As DseTrans_CutOffDB.tbltrans_aop1Row
        Dim FrmAOP As New FrmTransaksiAOP2

        Dim TblPrimary_AOPDP As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Approval Of Purchase Data", Me.StatusOnProses)
            DtaAOP.Connection.ConnectionString = Me.SumberConstring
            DtaAOP.FillBy_OutCutOff(TblAOP_Out, ClsTransaksiAOP.EnumTipeAop.AOP, Me.EndDateCutOff, 0)

            'transfer dulu aop dp
            Me.TransferAOP_DP(TblAOP_Out)

            If TblAOP_Out.Count > 0 Then
                TblPrimary_AOPDP = TblPrimary_AOP.Copy
                Me.SetProgressBarNLabel("Transfer Downpayment Approval Of Purchase Data", BatchProgBar, True, TblAOP_Out.Count)
                For Each rowAOP In TblAOP_Out.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Approval Of Purchase Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    FrmAOP.FromCutOff_AOP_Create(rowAOP.Primary, Me.TblPrimary_PO, Me.TblPrimary_GRN, TblPrimary_AOPDP, Me.TblPrimary_PR, Me.ObjCutOff_Core)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmAOP.FromCutOff_AOP_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowAOP.Primary, PrimaryUpdate, Me.TblPrimary_AOP)
                Next
            End If
            Me.RowUpdate("Transfer Approval Of Purchase Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferAOP")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmAOP.Dispose()
            DtaAOP.Dispose()
            TblAOP_Out.Dispose()
            TblPrimary_AOPDP.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferQuickAOP() As Boolean
        Dim DtaAOP As New DseTrans_CutOffDBTableAdapters.tbltrans_aop1TableAdapter
        Dim TblAOP_Out As New DseTrans_CutOffDB.tbltrans_aop1DataTable
        Dim rowAOP_Q As DseTrans_CutOffDB.tbltrans_aop1Row
        Dim DtaPO_no As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_po_noTableAdapter
        Dim TblPO_no As New Core_Module.DseTrans_CutOffDB_Core.tbl_po_noDataTable
        Dim DtaSIP_no As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_stockinout_noTableAdapter
        Dim TblSIP_no As New Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noDataTable
        Dim FrmAOPQuick As New FrmTransaksiAOP_Quick2

        Dim NoPO As String = ""
        Dim NoSIP As String = ""
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Quick AOP Data", Me.StatusOnProses)
            DtaAOP.Connection.ConnectionString = Me.SumberConstring
            DtaPO_no.Connection.ConnectionString = Me.SumberConstring
            DtaSIP_no.Connection.ConnectionString = Me.SumberConstring
            DtaAOP.FillBy_OutCutOff(TblAOP_Out, ClsTransaksiAOP.EnumTipeAop.AOP, Me.EndDateCutOff, 1)
            DtaPO_no.Fill(TblPO_no, Me.EndDateCutOff)
            DtaSIP_no.Fill(TblSIP_no, Me.EndDateCutOff, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase)

            Me.SetProgressBarNLabel("Transfer Quick AOP Data", BatchProgBar, True, TblAOP_Out.Count)
            For Each rowAOP_Q In TblAOP_Out.Rows
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Quick AOP Data", BatchProgBar)
                Me.ObjCutOff.GetNoPO_SIPByAOPQuick(TblPO_no, TblSIP_no, rowAOP_Q.Primary, NoPO, NoSIP)

                Me.SetDefaultConnectionString()
                FrmAOPQuick.FromCutOff_AOPQuick_Create(rowAOP_Q.Primary, rowAOP_Q.No_Transaksi)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmAOPQuick.FromCutOff_AOPQuick_Save(NoPO, NoSIP)
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowAOP_Q.Primary, PrimaryUpdate, Me.TblPrimary_AOP)
            Next
            Me.RowUpdate("Transfer Quick AOP Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferQuickAOP")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            DtaAOP.Dispose()
            TblAOP_Out.Dispose()
            DtaPO_no.Dispose()
            TblPO_no.Dispose()
            DtaSIP_no.Dispose()
            TblSIP_no.Dispose()
            FrmAOPQuick.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferPOC() As Boolean
        Dim DtaPOC As New DseTrans_CutOffDBTableAdapters.tbltrans_pocTableAdapter
        Dim TblPOC_Out As New DseTrans_CutOffDB.tbltrans_pocDataTable
        Dim rowPOC As DseTrans_CutOffDB.tbltrans_pocRow
        Dim FrmPOC As New FrmTransaksiPOC
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer PO Cancellation Data", Me.StatusOnProses)
            DtaPOC.Connection.ConnectionString = Me.SumberConstring
            DtaPOC.FillBy_OutCutOff(TblPOC_Out, Me.EndDateCutOff)
            Me.SetProgressBarNLabel("Transfer PO Cancellation Data", BatchProgBar, True, TblPOC_Out.Count)
            For Each rowPOC In TblPOC_Out.Rows
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer PO Cancellation Data", BatchProgBar)

                'Get Primary PO New
                Me.SetDefaultConnectionString()
                FrmPOC.FromCutOff_POC_Create(rowPOC.Primary, Me.TblPrimary_PO, Me.ObjCutOff_Core)
                Me.SetTargetConnectionString()
                FrmPOC.FromCutOff_POC_Save()
            Next
            Me.RowUpdate("Transfer PO Cancellation Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferPOC")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            DtaPOC.Dispose()
            TblPOC_Out.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function SetOrderBalancePurchase() As Boolean
        Me.SetTargetConnectionString()
        Dim DtaPO As New DseTrans_CutOffDBTableAdapters.tbltrans_poTableAdapter
        Dim TblPO As New DseTrans_CutOffDB.tbltrans_poDataTable
        Dim rowPO As DseTrans_CutOffDB.tbltrans_poRow
        Dim ObjPOItem As New Purchase_Module.ClsTransaksiPOItem

        Try
            DtaPO.Connection.ConnectionString = Me.TargetConstring
            DtaPO.Fill_All(TblPO)

            Me.RowUpdate("Set Purchase Order Balance", Me.StatusOnProses)
            If TblPO.DefaultView.Count > 0 Then
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = TblPO.Rows.Count
                For Each rowPO In TblPO.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.ProgressBar1.Value += 1
                    Me.Label1.Text = "Set Purchase Order Balance " & ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100, 2, _
                    MidpointRounding.AwayFromZero), 2) + "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)

                    ObjPOItem.SetOrderBalance(rowPO.Primary)
                Next
            End If

            Me.RowUpdate("Set Purchase Order Balance", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
            Return False
        Finally
            DtaPO.Dispose()
            TblPO.Dispose()
            ObjPOItem = Nothing
            Me.SetDefaultConnectionString()
        End Try
    End Function

    'TRANSFER SALES MODULE
    Public Function TransferQuotation() As Boolean
        Dim ListPrimary As New List(Of Double)

        Try
            Me.RowUpdate("Transfer Quotation Data", Me.StatusOnProses)
            Me.TblListTransfer.DefaultView.RowFilter = Nothing
            Me.TblListTransfer.DefaultView.RowFilter = "[Description] = 'Transfer Quotation Data'"
            For Each dv As DataRowView In Me.TblListTransfer.DefaultView
                Me.RowListTransfer = dv.Row
                Me.ProgressBar1.Value = 0
                If Me.TransferMaster_UsingScript(Me.RowListTransfer.Tabel_Name, ListPrimary) = False Then
                    Return False
                End If
            Next
            Me.RowUpdate("Transfer Quotation Data", Me.StatusOK)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferSO() As Boolean
        Dim DtaSO As New DseTrans_CutOffDBTableAdapters.tbltrans_soTableAdapter
        Dim TblSO As New DseTrans_CutOffDB.tbltrans_soDataTable
        Dim DtaSO_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_so_itemTableAdapter
        Dim TblSO_Item, TblSO_Item_Filter As New DseTrans_CutOffDB.tbltrans_so_itemDataTable
        Dim DtaSO_ItemDel As New DseTrans_CutOffDBTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TblSO_ItemDel, TblSO_ItemDel_Filter As New DseTrans_CutOffDB.tbltrans_so_itemdeliveryDataTable
        Dim rowSO As DseTrans_CutOffDB.tbltrans_soRow
        Dim FrmSO As New Sales_Module.FrmTransaksiSO

        Dim PrimaryUpdate As Double = -1
        Dim PrimarySO_Max As Double = -1
        Dim PrimaryBatch As Double = 500
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Sales Order Data", Me.StatusOnProses)
            Me.Label1.Text = "Get Sales Order Balance Data, Please Wait..."
            DtaSO.Connection.ConnectionString = Me.SumberConstring
            DtaSO_Item.Connection.ConnectionString = Me.SumberConstring
            DtaSO_ItemDel.Connection.ConnectionString = Me.SumberConstring

            DtaSO.Fill_All(TblSO)
            DtaSO_Item.FillBy_PeriodCutOff(TblSO_Item, Me.EndDateCutOff)
            DtaSO_ItemDel.FillBy_PeriodCutOff(TblSO_ItemDel, Me.EndDateCutOff)

            'Get SOB
            Me.TransferSOB(TblSO, TblSO_Item, TblSO_ItemDel)

            DtaSO_Item.Connection.ConnectionString = Me.SumberConstring
            DtaSO_ItemDel.Connection.ConnectionString = Me.SumberConstring
            DtaSO_Item.ClearBeforeFill = False
            DtaSO_ItemDel.ClearBeforeFill = False
            DtaSO_Item.FillBy_OutCutOff(TblSO_Item, Me.EndDateCutOff)
            DtaSO_ItemDel.FillBy_OutCutOff(TblSO_ItemDel, Me.EndDateCutOff)

            Me.SetTargetConnectionString()
            TblSO.DefaultView.RowFilter = "[Primary Quick Invoice] = -1"
            TblSO.DefaultView.Sort = "[Primary]"
            If TblSO.DefaultView.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblSO.DefaultView.Count)
                For Each dv As DataRowView In TblSO.DefaultView
                    rowSO = dv.Row
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    If PrimarySO_Max < rowSO.Primary Then
                        PrimarySO_Max = rowSO.Primary + PrimaryBatch
                        Me.ObjCutOff.GetTblSO_Filter(TblSO_Item, TblSO_Item_Filter, _
                        TblSO_ItemDel, TblSO_ItemDel_Filter, rowSO.Primary, PrimarySO_Max)
                    End If

                    Me.SetProgressBarNLabel("Transfer Sales Order Data", BatchProgBar)

                    PrimaryUpdate = FrmSO.FromCutOff_SO(rowSO, TblSO, TblSO_Item_Filter, TblSO_ItemDel_Filter, TblDec, rowSO.Created_By)
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowSO.Primary, PrimaryUpdate, Me.TblPrimary_SO)
                Next
            End If

            Me.RowUpdate("Transfer Sales Order Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
            Return False
        Finally
            DtaSO.Dispose()
            DtaSO_Item.Dispose()
            DtaSO_ItemDel.Dispose()
            TblSO.Dispose()
            TblSO_Item.Dispose()
            TblSO_ItemDel.Dispose()
            TblSO_Item_Filter.Dispose()
            TblSO_ItemDel_Filter.Dispose()
            FrmSO.Dispose()
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferSOB(ByRef TblSO As DseTrans_CutOffDB.tbltrans_soDataTable, _
    ByRef TblSO_Item As DseTrans_CutOffDB.tbltrans_so_itemDataTable, _
    ByRef TblSO_ItemDel As DseTrans_CutOffDB.tbltrans_so_itemdeliveryDataTable) As Boolean
        Dim DtaDO As New DseTrans_CutOffDBTableAdapters.tbltrans_doTableAdapter
        Dim TblDo3 As New DseTrans_CutOffDB.tbltrans_doDataTable
        Dim DtaDoItem As New DseTrans_CutOffDBTableAdapters.tbltrans_do_itemTableAdapter
        Dim TblDoItem_OutSumQty, TblDoItem_OutSumQty_Filter, TblDoItem3 As New DseTrans_CutOffDB.tbltrans_do_itemDataTable
        Dim DtaSOCItem As New DseTrans_CutOffDBTableAdapters.tbltrans_soc_itemTableAdapter
        Dim TblSOCItem_OutSumQty, TblSOCItem_OutSumQty_Filter As New DseTrans_CutOffDB.tbltrans_soc_itemDataTable
        Dim DtaSISRItem As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TblSISRItem_OutSumQty As New DseTrans_CutOffDB.tbltransaksi_stockinout_itemDataTable
        Dim DtaSRItem As New DseTrans_CutOffDBTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim TblSRItem_OutSumQty, TblSRItem_OutSumQty_filter As New DseTrans_CutOffDB.tbltrans_salesreturn_itemDataTable
        Dim vTblDec_SO As New DseTrans_CutOffDB.tblmaster_productdecimalDataTable
        Dim rowSO As DseTrans_CutOffDB.tbltrans_soRow
        Dim rowSO_Item As DseTrans_CutOffDB.tbltrans_so_itemRow
        Dim rowSO_ItemDel As DseTrans_CutOffDB.tbltrans_so_itemdeliveryRow

        Dim AddSO_ As Boolean = False
        Dim PrimarySO_Max As Double = 1
        Dim PrimaryBatch As Integer = 500
        Dim TransSOItem As Boolean = False
        Dim BatchProgBar As Integer = 1

        Try
            DtaDO.Connection.ConnectionString = Me.SumberConstring
            DtaDoItem.Connection.ConnectionString = Me.SumberConstring
            DtaSRItem.Connection.ConnectionString = Me.SumberConstring
            DtaSISRItem.Connection.ConnectionString = Me.SumberConstring
            DtaSOCItem.Connection.ConnectionString = Me.SumberConstring

            Me.TblDo_Out = New DseTrans_CutOffDB.tbltrans_doDataTable
            DtaDO.FillBy_OutCutOff(Me.TblDo_Out, Me.EndDateCutOff)
            DtaDO.FillBy_PeriodCutOff(TblDo3, Me.EndDateCutOff)
            DtaDoItem.FillBy_PeriodCutOffSumQty(TblDoItem3, Me.EndDateCutOff)
            DtaDoItem.FillBy_OutCutOffSumQty(TblDoItem_OutSumQty, Me.EndDateCutOff)
            DtaSRItem.FillBy_OutCutOffSumQty(TblSRItem_OutSumQty, Me.EndDateCutOff)
            DtaSOCItem.FillBy_OutCutOffSumQty(TblSOCItem_OutSumQty, Me.EndDateCutOff)
            DtaSISRItem.FillBy_OutCutOffSumQty(TblSISRItem_OutSumQty, Me.EndDateCutOff, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn)

            Me.GetTblDo_FromStockOutSales(TblDo3, TblDoItem3, TblDoItem_OutSumQty)

            'Cari sales order dalam period cut off yang masih mempunyai balance
            Me.RowUpdate("Transfer Sales Order Balance Data", Me.StatusOnProses)
            TblSO.DefaultView.RowFilter = "[Primary Quick Invoice] = -1 AND [Tanggal Transaksi] <= #" & Me.EndDateCutOff & "#"
            TblSO.DefaultView.Sort = "[Primary] DESC"
            If TblSO.DefaultView.Count > 0 Then
                vTblDec_SO = TblDec.Copy
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblSO.DefaultView.Count)
                For i As Integer = TblSO.DefaultView.Count - 1 To 0 Step -1
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    AddSO_ = False
                    rowSO = TblSO.DefaultView(i).Row
                    Me.SetProgressBarNLabel("Transfer Sales Order Balance Data", BatchProgBar)

                    If PrimarySO_Max < rowSO.Primary Then
                        PrimarySO_Max = rowSO.Primary + PrimaryBatch
                        Me.ObjCutOff.GetTblSOB_Filter(TblDoItem_OutSumQty, TblDoItem_OutSumQty_Filter, TblSOCItem_OutSumQty, _
                        TblSOCItem_OutSumQty_Filter, TblSRItem_OutSumQty, TblSRItem_OutSumQty_filter, rowSO.Primary, PrimarySO_Max)
                    End If

                    TblSO_Item.DefaultView.RowFilter = "[Primary Transaksi] = " & rowSO.Primary
                    If TblSO_Item.DefaultView.Count > 0 Then
                        For i2 As Integer = TblSO_Item.DefaultView.Count - 1 To 0 Step -1
                            rowSO_Item = TblSO_Item.DefaultView(i2).Row
                            'cek apakah item dalam so masih mempunyai balance, jika tidak hapus baris itemnya
                            'jika semua item dalam so tidak mempunyai balance maka hapus transaksi so nya.
                            TransSOItem = Me.ObjCutOff.GetBalanceSO(TblDoItem_OutSumQty_Filter, TblSOCItem_OutSumQty_Filter, TblSRItem_OutSumQty_filter, rowSO_Item, vTblDec_SO, TblSISRItem_OutSumQty)
                            If TransSOItem = False Then
DeleteItemDel:
                                TblSO_ItemDel.DefaultView.RowFilter = "[identifier] = '" & rowSO_Item.identifier & "'"
                                For Each dvDel As DataRowView In TblSO_ItemDel.DefaultView
                                    rowSO_ItemDel = dvDel.Row
                                    TblSO_ItemDel.Rows.Remove(rowSO_ItemDel)
                                Next
                                If TransSOItem = False Then
                                    TblSO_Item.Rows.Remove(rowSO_Item)
                                End If
                            Else
                                AddSO_ = True
                                If rowSO_Item.Qty_Multi = 0 Then
                                    GoTo DeleteItemDel
                                End If
                            End If
                        Next

                        If AddSO_ = False Then
                            TblSO.Rows.Remove(rowSO)
                        End If
                    Else
                        TblSO.Rows.Remove(rowSO)
                    End If
                Next
            End If

            Me.RowUpdate("Transfer Sales Order Balance Data", Me.StatusOK)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
            Return False
        Finally
            DtaDO.Dispose()
            DtaDoItem.Dispose()
            DtaSRItem.Dispose()
            DtaSISRItem.Dispose()
            DtaSOCItem.Dispose()
            TblDo3.Dispose()
            TblDoItem3.Dispose()
            TblDoItem_OutSumQty.Dispose()
            TblSISRItem_OutSumQty.Dispose()
            TblSRItem_OutSumQty.Dispose()
            TblSOCItem_OutSumQty.Dispose()
            TblDoItem_OutSumQty_Filter.Dispose()
            TblSRItem_OutSumQty_filter.Dispose()
            TblSOCItem_OutSumQty_Filter.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function GetTblDo_FromStockOutSales(ByVal TblDo3 As DseTrans_CutOffDB.tbltrans_doDataTable, _
    ByVal TblDoItem3 As DseTrans_CutOffDB.tbltrans_do_itemDataTable, _
    ByRef TblDoItem_OutSumQty As DseTrans_CutOffDB.tbltrans_do_itemDataTable) As Boolean
        Dim TblDo3_filter As New DseTrans_CutOffDB.tbltrans_doDataTable
        Dim TblDoItem3_filter As New DseTrans_CutOffDB.tbltrans_do_itemDataTable
        Dim DtaSOS As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TblSOS_Out, TblSISR_Out As New DseTrans_CutOffDB.tbltransaksi_stockinoutDataTable
        Dim rowDo, rowDo3 As DseTrans_CutOffDB.tbltrans_doRow
        Dim row, rowDOItem_SumQty As DseTrans_CutOffDB.tbltrans_do_itemRow
        Dim rowSos, rowSisr As DseTrans_CutOffDB.tbltransaksi_stockinoutRow

        Dim index As Integer = -1
        Dim ListPrimarySO As New List(Of Double)
        Dim PrimaryDO_Max As Double = 0
        Dim PrimaryBatch As Integer = 1000
        Dim Count As Integer = 0
        Dim count2 As Integer = 0
        Dim BatchProgBar As Integer = 1

        Try
            DtaSOS.Connection.ConnectionString = Me.SumberConstring
            DtaSOS.FillBy_OutCutOff(TblSOS_Out, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, Me.EndDateCutOff)
            DtaSOS.FillBy_OutCutOff(TblSISR_Out, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn, Me.EndDateCutOff)

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TblSOS_Out.Rows.Count
            TblSOS_Out.DefaultView.Sort = "[Primary DO Sales] ASC"
            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSOS_Out.DefaultView.Count)
            For i As Integer = 0 To TblSOS_Out.DefaultView.Count - 1
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Sales Order Balance Data (1 of 2)", BatchProgBar)

                rowSos = TblSOS_Out.DefaultView(i).Row
                If PrimaryDO_Max < rowSos.Primary_DO_Sales Then
                    PrimaryDO_Max = rowSos.Primary_DO_Sales + PrimaryBatch
                    Me.ObjCutOff.GetTblDObySOS(TblDo3, TblDo3_filter, TblDoItem3, TblDoItem3_filter, rowSos.Primary_DO_Sales, PrimaryDO_Max)
                End If

                index = i
                TblDo3_filter.DefaultView.RowFilter = "[Primary]=" & rowSos.Primary_DO_Sales
                If TblDo3_filter.DefaultView.Count > 0 Then
                    Count += 1
                    rowDo3 = TblDo3_filter.DefaultView(0).Row
                    rowDo = Me.TblDo_Out.Newtbltrans_doRow
                    rowDo.Primary = rowSos.Primary_DO_Sales
                    rowDo.Tanggal_Transaksi = Now
                    rowDo.No_Transaksi = rowDo3.No_Transaksi
                    rowDo.Primary_Stock_Out_Sales = rowSos.Primary
                    rowDo.Primary_Invoice = -1
                    rowDo.Sudah_Buat_Invoice = 0
                    Me.TblDo_Out.Addtbltrans_doRow(rowDo)

                    TblDoItem3_filter.DefaultView.RowFilter = "[Primary DO]=" & rowSos.Primary_DO_Sales
                    If TblDoItem3_filter.DefaultView.Count > 0 Then
                        For ij As Integer = 0 To TblDoItem3_filter.DefaultView.Count - 1
                            row = TblDoItem3_filter.DefaultView(ij).Row
                            TblDoItem_OutSumQty.DefaultView.RowFilter = "[Primary Product]=" & row.Primary_Product & " AND [Primary SO] = " & row.Primary_SO
                            If TblDoItem_OutSumQty.DefaultView.Count = 0 Then
                                rowDOItem_SumQty = TblDoItem_OutSumQty.Newtbltrans_do_itemRow
                                With rowDOItem_SumQty
                                    .Primary = row.Primary
                                    .Primary_DO = row.Primary_DO
                                    .Primary_Product = row.Primary_Product
                                    .Primary_SO = row.Primary_SO
                                    .Primary_Units = row.Primary_Units
                                    .Primary_Units_Multi = row.Primary_Units_Multi
                                    .Total_Qty = row.Total_Qty
                                    .Total_Qty_Multi = row.Total_Qty_Multi
                                End With
                                TblDoItem_OutSumQty.Addtbltrans_do_itemRow(rowDOItem_SumQty)
                            Else
                                rowDOItem_SumQty = TblDoItem_OutSumQty.DefaultView(0).Row
                                rowDOItem_SumQty.Total_Qty = row.Total_Qty
                                rowDOItem_SumQty.Total_Qty_Multi = row.Total_Qty_Multi
                            End If
                        Next
                    End If
                End If
            Next

            PrimaryDO_Max = 0
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TblSISR_Out.Rows.Count
            TblSISR_Out.DefaultView.Sort = "[Primary DO Sales] ASC"
            For i As Integer = 0 To TblSISR_Out.DefaultView.Count - 1
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Sales Order Balance Data (2 of 2)", BatchProgBar)

                rowSisr = TblSISR_Out.DefaultView(i).Row
                rowSisr.Primary_DO_Sales = rowSisr.Primary_DO_Sales
                If PrimaryDO_Max < rowSisr.Primary_DO_Sales Then
                    PrimaryDO_Max = rowSisr.Primary_DO_Sales + PrimaryBatch
                    Me.ObjCutOff.GetTblDObySOS(TblDo3, TblDo3_filter, TblDoItem3, TblDoItem3_filter, rowSisr.Primary_DO_Sales, PrimaryDO_Max)
                End If

                index = i
                TblDo3_filter.DefaultView.RowFilter = "[Primary]=" & rowSisr.Primary_DO_Sales
                If TblDo3_filter.DefaultView.Count > 0 Then
                    count2 += 1
                    rowDo3 = TblDo3_filter.DefaultView(0).Row
                    Me.TblDo_Out.DefaultView.RowFilter = "[Primary]=" & rowSisr.Primary_DO_Sales
                    If Me.TblDo_Out.DefaultView.Count = 0 Then
                        rowDo = Me.TblDo_Out.Newtbltrans_doRow
                        rowDo.Primary = rowSisr.Primary_DO_Sales
                        rowDo.Tanggal_Transaksi = Now
                        rowDo.No_Transaksi = rowDo3.No_Transaksi
                        rowDo.Primary_Stock_Out_Sales = "-3"
                        rowDo.Primary_Invoice = -1
                        rowDo.Sudah_Buat_Invoice = 0
                        Me.TblDo_Out.Addtbltrans_doRow(rowDo)
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message & "- row ::" & index, MsgBoxStyle.Critical, "Error: GetTblDo_FromStockOutSales")
            Return False
        Finally
            DtaSOS.Dispose()
            TblSOS_Out.Dispose()
            TblSISR_Out.Dispose()
            TblDo3_filter.Dispose()
            TblDoItem3_filter.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferDO() As Boolean
        Dim rowDo As DseTrans_CutOffDB.tbltrans_doRow
        Dim FrmDo As New Sales_Module.FrmTransaksiDO3
        Dim PrimaryUpdate As Double = -1
        Dim TransItemDO As Boolean = True
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Delivery Order Data", Me.StatusOnProses)
            'buang do yang dibuat dari quick inv
            Me.ObjCutOff.DeleteDOFromQuickInv(Me.TblDo_Out, Me.EndDateCutOff, Me.SumberConstring)

            Me.TblDo_Out.DefaultView.RowFilter = Nothing
            Me.TblDo_Out.DefaultView.Sort = "[Primary] ASC"
            If Me.TblDo_Out.DefaultView.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, Me.TblDo_Out.DefaultView.Count)
                For Each dv As DataRowView In Me.TblDo_Out.DefaultView
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    rowDo = dv.Row
                    Me.SetProgressBarNLabel("Transfer Delivery Order Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    'saat trans sob dicek bila do yang ada dalam period cutt off di trans karena ada stock in sales retur yang merujuk ke do tsb
                    'maka item nya tidak usah di transfer
                    If rowDo.Primary_Stock_Out_Sales = -3 Then
                        TransItemDO = False
                    Else
                        TransItemDO = True
                    End If
                    FrmDo.FromCutOff_DeliveryOrder_Create(rowDo.Primary, Me.ObjCutOff_Core, Me.TblPrimary_SO, TransItemDO)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmDo.FromCutOff_DeliveryOrder_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowDo.Primary, PrimaryUpdate, Me.TblPrimary_DO)
                Next
            End If
            Me.RowUpdate("Transfer Delivery Order Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferDO")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.TblDo_Out.Dispose()
            FrmDo.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Public Function TransferStockOutSales() As Boolean
        Dim DtaSOS As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TblSOS_Out As New DseTrans_CutOffDB.tbltransaksi_stockinoutDataTable
        Dim rowSoS As DseTrans_CutOffDB.tbltransaksi_stockinoutRow
        Dim FrmSOS As New Inventory_Module.FrmTransaksiStockOutSales
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock Out Sales Data", Me.StatusOnProses)
            DtaSOS.Connection.ConnectionString = Me.SumberConstring
            DtaSOS.FillBy_OutCutOff(TblSOS_Out, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, Me.EndDateCutOff)

            TblSOS_Out.DefaultView.RowFilter = "[Primary Quick Acuan] = -1"
            TblSOS_Out.DefaultView.Sort = "[Primary] ASC"
            If TblSOS_Out.DefaultView.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblSOS_Out.DefaultView.Count)
                For Each dv As DataRowView In TblSOS_Out.DefaultView
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    rowSoS = dv.Row
                    Me.SetProgressBarNLabel("Transfer Stock Out Sales Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    FrmSOS.FromCutOff_StockOutSales_Create(rowSoS.Primary, Me.ObjCutOff_Core, TblPrimary_SO, TblPrimary_DO)
                    Me.SetTargetConnectionString()
                    FrmSOS.FromCutOff_StockOutSales_Save()
                Next
            End If

            Me.RowUpdate("Transfer Stock Out Sales Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockOutSales")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaSOS.Dispose()
            TblSOS_Out.Dispose()
            FrmSOS.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Public Function TransferStockInSalesReturn() As Boolean
        Dim DtaSISR As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TblSISR_Out As New DseTrans_CutOffDB.tbltransaksi_stockinoutDataTable
        Dim rowSISR As DseTrans_CutOffDB.tbltransaksi_stockinoutRow
        Dim FrmSISR As New Inventory_Module.FrmTransaksiStockInSalesReturn2
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock In Sales Return Data", Me.StatusOnProses)
            DtaSISR.Connection.ConnectionString = Me.SumberConstring
            DtaSISR.FillBy_OutCutOff(TblSISR_Out, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn, Me.EndDateCutOff)

            'karena stock in sales retur yang dibuat dari quick sr primary quick acuannya tetap -1
            'jd untuk membedakannya di lihat dari primary do bila = -1 berarti di buat dari quick sr
            TblSISR_Out.DefaultView.RowFilter = "[Primary DO Sales] <> -1"
            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSISR_Out.DefaultView.Count)
            For Each dv As DataRowView In TblSISR_Out.DefaultView
                rowSISR = dv.Row
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Stock In Sales Return Data", BatchProgBar)

                Me.SetDefaultConnectionString()
                FrmSISR.FromCutOff_StockInSalesReturn_Create(rowSISR.Primary, Me.ObjCutOff_Core, TblPrimary_SO, TblPrimary_DO)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmSISR.FromCutOff_StockInSalesReturn_Save()
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowSISR.Primary, PrimaryUpdate, Me.TblPrimary_SISR)
            Next

            Me.RowUpdate("Transfer Stock In Sales Return Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockInSalesReturn")
            Return False
        Finally
            DtaSISR.Dispose()
            TblSISR_Out.Dispose()
            FrmSISR.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Public Function TransferSalesReturn() As Boolean
        Dim DtaSR As New DseTrans_CutOffDBTableAdapters.tbltrans_salesreturn1TableAdapter
        Dim TblSR_Out As New DseTrans_CutOffDB.tbltrans_salesreturn1DataTable
        Dim rowSR As DseTrans_CutOffDB.tbltrans_salesreturn1Row
        Dim FrmSR As New Sales_Module.FrmTransaksiSalesReturn
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Sales Return Data", Me.StatusOnProses)
            DtaSR.Connection.ConnectionString = Me.SumberConstring
            DtaSR.FillBy_OutCutOff(TblSR_Out, Me.EndDateCutOff, 0)

            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSR_Out.Rows.Count)
            For Each rowSR In TblSR_Out
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Sales Return Data", BatchProgBar)

                Me.SetDefaultConnectionString()
                FrmSR.FromCutOff_SalesReturn_Create(rowSR.Primary, Me.ObjCutOff_Core, TblPrimary_SISR, TblPrimary_DO, TblPrimary_SO)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmSR.FromCutOff_SalesReturn_Save()
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowSR.Primary, PrimaryUpdate, Me.TblPrimary_SR)
            Next

            Me.RowUpdate("Transfer Sales Return Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesReturn")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaSR.Dispose()
            TblSR_Out.Dispose()
            FrmSR.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Public Function TransferQuickSalesReturn() As Boolean
        Dim DtaSR As New DseTrans_CutOffDBTableAdapters.tbltrans_salesreturn1TableAdapter
        Dim TblSR_Out As New DseTrans_CutOffDB.tbltrans_salesreturn1DataTable
        Dim rowQSR As DseTrans_CutOffDB.tbltrans_salesreturn1Row
        Dim FrmQSR As New Sales_Module.FrmTransaksiSalesReturn_Quick
        Dim Dta_SisrNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_stockinout_noTableAdapter
        Dim Tbl_SisrNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noDataTable
        Dim NoStockInSR As String = ""
        Dim PrimaryUpdate As Double = -1
        Dim PrimaryUpdate_SISR As Double = -1

        Try
            Me.RowUpdate("Transfer Quick Sales Return Data", Me.StatusOnProses)
            DtaSR.Connection.ConnectionString = Me.SumberConstring
            DtaSR.FillBy_OutCutOff(TblSR_Out, Me.EndDateCutOff, 1)

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TblSR_Out.Count

            Dta_SisrNo.Connection.ConnectionString = Me.SumberConstring
            Dta_SisrNo.Fill(Tbl_SisrNo, Me.EndDateCutOff, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn)

            For Each rowQSR In TblSR_Out
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.ProgressBar1.Value += 1
                Me.Label1.Text = "Transfer Quick Sales Return Data " & ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100, 2, _
                MidpointRounding.AwayFromZero), 2) + "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)

                NoStockInSR = Me.ObjCutOff_Core.GetNoInOut(Tbl_SisrNo, rowQSR.Primary)

                Me.SetDefaultConnectionString()
                FrmQSR.FromCutOff_SReturn_Quick_Create(rowQSR.Primary)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmQSR.FromCutOff_SReturn_Quick_Save(NoStockInSR, PrimaryUpdate_SISR)
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowQSR.Primary, PrimaryUpdate, Me.TblPrimary_SR)
                'tambahkan juga primary update stock in sales return 
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowQSR.Primary_Stock_In_Sales_Return, PrimaryUpdate_SISR, Me.TblPrimary_SISR)
            Next

            Me.RowUpdate("Transfer Quick Sales Return Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferQuickSalesReturn")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaSR.Dispose()
            TblSR_Out.Dispose()
            Dta_SisrNo.Dispose()
            Tbl_SisrNo.Dispose()
            FrmQSR.Dispose()
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Sub UpdatePrimaryInv_Salesreturn()
        Dim DtaSR As New DseTrans_CutOffDBTableAdapters.tbltrans_salesreturn1TableAdapter
        Dim TblSR_Out As New DseTrans_CutOffDB.tbltrans_salesreturn1DataTable
        Dim rowSR As DseTrans_CutOffDB.tbltrans_salesreturn1Row
        Dim ObjDO As New Sales_Module.ClsTransaksiDO
        Dim PrimaryInvoice As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.SetTargetConnectionString()
            DtaSR.Connection.ConnectionString = Me.TargetConstring
            DtaSR.FillBy_OutCutOff(TblSR_Out, Me.EndDateCutOff, 0)

            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSR_Out.Rows.Count)
            For Each rowSR In TblSR_Out
                If Me.Bgw1.CancellationPending Then
                    Exit Sub
                End If

                Me.SetProgressBarNLabel("Update Invoice In Sales Return Data", BatchProgBar)

                If rowSR.Tipe_SR = Sales_Module.ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak And rowSR.Quick_Sales_Return = 0 Then
                    PrimaryInvoice = ObjDO.GetPrimaryInvoice(rowSR.Primary_DO)
                    DtaSR.UpdateQuery_PrimaryInvoiceByPrimarySR(PrimaryInvoice, rowSR.Primary)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesReturn")
            Me.SetDefaultConnectionString()
        Finally
            DtaSR.Dispose()
            TblSR_Out.Dispose()
            ObjDO = Nothing
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Sub

    Public Function TransferDownPaymentInvoice(ByVal TblSalesInv As DseTrans_CutOffDB.tbltrans_salesinvoiceDataTable) As Boolean
        Dim DtaSalesInvDP As New DseTrans_CutOffDBTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TblSalesInvDP As New DseTrans_CutOffDB.tbltrans_salesinvoiceDataTable
        Dim DtaSalesInvItemDp As New DseTrans_CutOffDBTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        Dim TblSalesInvItemDp As New DseTrans_CutOffDB.tbltrans_salesinvoice_itemdpDataTable
        Dim rowDPInv As DseTrans_CutOffDB.tbltrans_salesinvoiceRow
        Dim FrmDPInv As New Sales_Module.FrmTransaksiSalesInvoiceDP
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Downpayment Invoice Data", Me.StatusOnProses)
            DtaSalesInvDP.Connection.ConnectionString = Me.SumberConstring
            DtaSalesInvDP.Fill_All(TblSalesInvDP, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP)
            DtaSalesInvItemDp.Connection.ConnectionString = Me.SumberConstring
            DtaSalesInvItemDp.Fill_All(TblSalesInvItemDp)

            TblSalesInvDP.DefaultView.RowFilter = "[Tanggal Transaksi] <= #" & Me.EndDateCutOff & "#"
            For Each dv As DataRowView In TblSalesInvDP.DefaultView
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                rowDPInv = dv.Row
                TblSalesInvItemDp.DefaultView.RowFilter = Nothing
                TblSalesInvItemDp.DefaultView.RowFilter = "[Primary Invoice DP] = " & rowDPInv.Primary
                If TblSalesInvItemDp.DefaultView.Count = 0 Then
                    TblSalesInvDP.Rows.Remove(rowDPInv)
                Else
                    TblSalesInv.DefaultView.RowFilter = Nothing
                    TblSalesInv.DefaultView.RowFilter = "[Primary] = " & rowDPInv.Primary & " AND [Tanggal Transaksi] <= #" & Me.EndDateCutOff & "#"
                    If TblSalesInv.DefaultView.Count = 0 Then
                        TblSalesInvDP.Rows.Remove(rowDPInv)
                    End If
                End If
            Next

            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSalesInvDP.Rows.Count)
            For Each rowDPInv In TblSalesInvDP
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Downpayment Invoice Data", BatchProgBar, True, TblSalesInvDP.Rows.Count)

                Dim ApakahSimpanJurnal As Boolean = True
                If rowDPInv.Tanggal_Transaksi.Date <= Me.EndDateCutOff Then
                    ApakahSimpanJurnal = False
                End If
                Me.SetDefaultConnectionString()
                FrmDPInv.FromCutOff_SalesInvoiceDP_Create(rowDPInv.Primary, Me.ObjCutOff_Core, Me.TblPrimary_SO)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmDPInv.FromCutOff_SalesInvoiceDP_Save(ApakahSimpanJurnal)
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowDPInv.Primary, PrimaryUpdate, Me.TblPrimary_Inv)
            Next

            Me.RowUpdate("Transfer Downpayment Invoice Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferDownPaymentInvoice")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaSalesInvDP.Dispose()
            TblSalesInvDP.Dispose()
            DtaSalesInvItemDp.Dispose()
            TblSalesInvItemDp.Dispose()
            FrmDPInv.Dispose()
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferSalesInvoice() As Boolean
        Dim DtaSalesInv As New DseTrans_CutOffDBTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TblSalesInv As New DseTrans_CutOffDB.tbltrans_salesinvoiceDataTable
        Dim rowSalesInv As DseTrans_CutOffDB.tbltrans_salesinvoiceRow
        Dim FrmSalesInv As New Sales_Module.FrmTransaksiSalesInvoice
        Dim PrimaryUpdate As Double = -1
        Dim tblPrimaryInvDP As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim BatchProgBar As Integer = 1

        Try
            DtaSalesInv.Connection.ConnectionString = Me.SumberConstring
            DtaSalesInv.FillBy_OutCutOff(TblSalesInv, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice, Me.EndDateCutOff, 0)

            Me.TransferDownPaymentInvoice(TblSalesInv)

            Me.RowUpdate("Transfer Sales Invoice Data", Me.StatusOnProses)
            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSalesInv.Count)
            tblPrimaryInvDP = Me.TblPrimary_Inv.Copy
            For Each rowSalesInv In TblSalesInv
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Sales Invoice Data", BatchProgBar)

                Me.SetDefaultConnectionString()
                FrmSalesInv.FromCutOff_SalesInvoice_Create(rowSalesInv.Primary, Me.ObjCutOff_Core, _
                Me.TblPrimary_DO, Me.TblPrimary_SO, tblPrimaryInvDP, Me.TblPrimary_SR)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmSalesInv.FromCutOff_SalesInvoice_Save()
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowSalesInv.Primary, PrimaryUpdate, Me.TblPrimary_Inv)
            Next
            Me.RowUpdate("Transfer Sales Invoice Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesInvoice")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaSalesInv.Dispose()
            TblSalesInv.Dispose()
            FrmSalesInv.Dispose()
            tblPrimaryInvDP.Dispose()
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferQuickSalesInvoice() As Boolean
        Dim DtaSalesInv As New DseTrans_CutOffDBTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TblSalesInv As New DseTrans_CutOffDB.tbltrans_salesinvoiceDataTable
        Dim rowQuickSalesInv As DseTrans_CutOffDB.tbltrans_salesinvoiceRow
        Dim FrmSalesInv As New Sales_Module.FrmTransaksiSalesInvoiceQuick
        Dim DtaDOno As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_do_noTableAdapter
        Dim TblDOno As New Core_Module.DseTrans_CutOffDB_Core.tbl_do_noDataTable

        Dim vNoDO As String = ""
        Dim PrimaryInvMin As Double = 1
        Dim PrimaryInvMax As Double = 999999
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Quick Invoice Data", Me.StatusOnProses)
            DtaSalesInv.Connection.ConnectionString = Me.SumberConstring
            DtaSalesInv.FillBy_OutCutOff(TblSalesInv, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice, Me.EndDateCutOff, 1)

            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSalesInv.Count)
            If TblSalesInv.Count > 0 Then
                TblSalesInv.DefaultView.Sort = "[Primary] ASC"
                rowQuickSalesInv = TblSalesInv.DefaultView(0).Row
                PrimaryInvMin = rowQuickSalesInv.Primary
                rowQuickSalesInv = TblSalesInv.DefaultView(TblSalesInv.DefaultView.Count - 1).Row
                PrimaryInvMax = rowQuickSalesInv.Primary
            End If

            Me.SetDefaultConnectionString()
            DtaDOno.Fill(TblDOno, Me.EndDateCutOff, PrimaryInvMax, PrimaryInvMin)

            For Each rowQuickSalesInv In TblSalesInv
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer Quick Invoice Data", BatchProgBar, True, TblSalesInv.Count)

                vNoDO = Me.ObjCutOff_Core.GetNo_DOByPrimaryInv(TblDOno, rowQuickSalesInv.Primary)

                Me.SetDefaultConnectionString()
                FrmSalesInv.FromCutOff_SalesInvoiceQuick_Create(rowQuickSalesInv.Primary, rowQuickSalesInv.No_Transaksi)
                Me.SetTargetConnectionString()
                PrimaryUpdate = FrmSalesInv.FromCutOff_SalesInvoiceQuick_Save(vNoDO)
                Me.ObjCutOff_Core.AddPrimaryUpdate(rowQuickSalesInv.Primary, PrimaryUpdate, Me.TblPrimary_Inv)
            Next
            Me.RowUpdate("Transfer Quick Invoice Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferQuickSalesInvoice")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaSalesInv.Dispose()
            TblSalesInv.Dispose()
            FrmSalesInv.Dispose()
            DtaDOno.Dispose()
            TblDOno.Dispose()
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferSOC() As Boolean
        Dim DtaSOC As New DseTrans_CutOffDBTableAdapters.tbltrans_socTableAdapter
        Dim TblSOC_Out As New DseTrans_CutOffDB.tbltrans_socDataTable
        Dim rowSOC As DseTrans_CutOffDB.tbltrans_socRow
        Dim FrmSOC As New Sales_Module.FrmTransaksiSOC
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer SO Cancellation Data", Me.StatusOnProses)
            DtaSOC.Connection.ConnectionString = Me.SumberConstring
            DtaSOC.FillBy_OutCutOff(TblSOC_Out, Me.EndDateCutOff)

            Me.SetProgressBarNLabel("", BatchProgBar, True, TblSOC_Out.DefaultView.Count)
            For Each rowSOC In TblSOC_Out
                If Me.Bgw1.CancellationPending Then
                    Return False
                End If

                Me.SetProgressBarNLabel("Transfer SO Cancellation Data", BatchProgBar)

                Me.SetDefaultConnectionString()
                FrmSOC.FromCutOff_SOCancellation_Create(rowSOC.Primary, Me.ObjCutOff_Core, Me.TblPrimary_SO)
                Me.SetTargetConnectionString()
                FrmSOC.FromCutOff_SOCancellation_Save()
            Next

            Me.RowUpdate("Transfer SO Cancellation Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesCancellation")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaSOC.Dispose()
            TblSOC_Out.Dispose()
            FrmSOC.Dispose()
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Private Sub SetOrderBalanceSales()
        Me.SetTargetConnectionString()
        Dim ObjAcc As New AccModule.ClsTransaksiClosing
        Dim Date_Conv As Date = ObjAcc.GetConvertionDate
        Dim ObjSo As New Sales_Module.ClsTransaksiSalesOrderItem
        Dim DtaSO As New DseTrans_CutOffDBTableAdapters.tbltrans_soTableAdapter
        Dim TblSO As New DseTrans_CutOffDB.tbltrans_soDataTable
        Dim rowSO As DseTrans_CutOffDB.tbltrans_soRow

        Try
            DtaSO.Connection.ConnectionString = Me.TargetConstring
            DtaSO.Fill_All(TblSO)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TblSO.Count
            For Each rowSO In TblSO.Rows
                With rowSO
                    Me.ProgressBar1.Value += 1
                    Label1.Text = "Set Order Balance Sales " & ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100, 2, _
                    MidpointRounding.AwayFromZero), 2) + "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)

                    ObjSo.Set_Order_Balance(rowSO.Primary)
                End With
            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaSO.Dispose()
            TblSO.Dispose()
            ObjSo = Nothing
            ObjAcc = Nothing
            Me.SetDefaultConnectionString()
        End Try
    End Sub

    'TRANSFER INVENTORY MODULE
    Public Function TransferStockInAdjust() As Boolean
        Dim DtaSIA As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TblSIA As New DseTrans_CutOffDB.tbltransaksi_stockinoutDataTable
        Dim rowSIA As DseTrans_CutOffDB.tbltransaksi_stockinoutRow
        Dim FrmSIA As New Inventory_Module.FrmTransaksiStockInAdj
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock In Adjust Data", Me.StatusOnProses)
            DtaSIA.Connection.ConnectionString = Me.SumberConstring
            DtaSIA.FillBy_OutCutOff(TblSIA, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust, Me.EndDateCutOff)

            ''Transfer semua stock in adjust di luar period cutoff
            If TblSIA.Rows.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblSIA.Rows.Count)
                For Each rowSIA In TblSIA.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Stock In Adjust Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    FrmSIA.FromCutOff_InAdjust_Create(rowSIA.Primary)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmSIA.FromCutOff_InAdjust_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowSIA.Primary, PrimaryUpdate, Me.TblPrimary_InAdj)
                Next
            End If

            Me.RowUpdate("Transfer Stock In Adjust Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockInAdjust")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmSIA.Dispose()
            DtaSIA.Dispose()
            TblSIA.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferStockOutAdjust() As Boolean
        Dim DtaSOA As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TblSOA As New DseTrans_CutOffDB.tbltransaksi_stockinoutDataTable
        Dim rowSOA As DseTrans_CutOffDB.tbltransaksi_stockinoutRow
        Dim FrmSOA As New Inventory_Module.FrmTransaksiStockOutAdj
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock Out Adjust Data", Me.StatusOnProses)
            DtaSOA.Connection.ConnectionString = Me.SumberConstring
            DtaSOA.FillBy_OutCutOff(TblSOA, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Me.EndDateCutOff)

            ''Transfer semua stock out adjust di luar period cutoff
            If TblSOA.Rows.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblSOA.Rows.Count)
                For Each rowSOA In TblSOA.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Stock Out Adjust Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    FrmSOA.FromCutOff_OutAdjust_Create(rowSOA.Primary)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmSOA.FromCutOff_OutAdjust_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowSOA.Primary, PrimaryUpdate, Me.TblPrimary_OutAdj)
                Next
            End If

            Me.RowUpdate("Transfer Stock Out Adjust Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockOutAdjust")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmSOA.Dispose()
            DtaSOA.Dispose()
            TblSOA.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferStockBundling() As Boolean
        Dim DtaBund As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TblBund As New DseTrans_CutOffDB.tbltransaksi_stockbundlingDataTable
        Dim rowBund As DseTrans_CutOffDB.tbltransaksi_stockbundlingRow
        Dim FrmBund As New Inventory_Module.FrmTransaksiStockBundling
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock Bundling Data", Me.StatusOnProses)
            DtaBund.Connection.ConnectionString = Me.SumberConstring
            DtaBund.FillBy_OutCutOff(TblBund, Me.EndDateCutOff)

            ''Transfer semua stock bundling di luar period cutoff
            If TblBund.Rows.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblBund.Rows.Count)
                For Each rowBund In TblBund.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Stock Bundling Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    FrmBund.FromCutOff_StockBundling_Create(rowBund.Primary)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmBund.FromCutOff_StockBundling_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowBund.Primary, PrimaryUpdate, Me.TblPrimary_Bund)
                Next
            End If

            Me.RowUpdate("Transfer Stock Bundling Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockBundling")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmBund.Dispose()
            DtaBund.Dispose()
            TblBund.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferStockTaking() As Boolean
        Dim DtaTaking As New DseTrans_CutOffDBTableAdapters.tbltrans_stock_takingTableAdapter
        Dim TblTaking, TblTakingNew As New DseTrans_CutOffDB.tbltrans_stock_takingDataTable
        Dim DtaTag As New DseTrans_CutOffDBTableAdapters.tbltrans_stock_taking_tagTableAdapter
        Dim TblTag As New DseTrans_CutOffDB.tbltrans_stock_taking_tagDataTable
        Dim rowTaking As DseTrans_CutOffDB.tbltrans_stock_takingRow
        Dim rowTag As DseTrans_CutOffDB.tbltrans_stock_taking_tagRow

        Dim DtaPenyesuaian As New DseTrans_CutOffDBTableAdapters.tbltrans_stock_taking_itempenyesuaianTableAdapter
        Dim TblPenyesuaian As New DseTrans_CutOffDB.tbltrans_stock_taking_itempenyesuaianDataTable
        Dim RowPenyesuaian As DseTrans_CutOffDB.tbltrans_stock_taking_itempenyesuaianRow

        Dim DtaLapCost As New DseTrans_CutOffDBTableAdapters.tbltrans_stock_taking_laporancostTableAdapter
        Dim TblLapCost As New DseTrans_CutOffDB.tbltrans_stock_taking_laporancostDataTable
        'Dim rowLapCost As DseTrans_CutOffDB.tbltrans_stock_taking_laporancostRow

        Dim DTASC As New Inventory_Module.dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim DTAProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim DTAProdQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLProdQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable
        'Dim RowProdQty As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Dim FrmTaking As New Inventory_Module.frmTransaksiStockTaking
        Dim FrmTag As New Inventory_Module.frmTransaksiStockTag

        Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Dim PrimaryUpdate As Double = -1
        Dim count, QtyIn, QtyOut As Double
        Dim SaldoAkhir As Decimal = 0
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Stock Taking Data", Me.StatusOnProses)
            DtaTaking.Connection.ConnectionString = Me.SumberConstring
            DtaTag.Connection.ConnectionString = Me.SumberConstring
            DtaTaking.FillBy_OutCutOff(TblTaking, Me.EndDateCutOff)
            DtaTag.Fill(TblTag)

            ''Transfer semua stock taking di luar period cutoff
            If TblTaking.Rows.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblTaking.Rows.Count)
                For Each rowTaking In TblTaking.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Stock Taking Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    FrmTaking.FromCutOff_StockTaking_Create(rowTaking.Primary)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmTaking.FromCutOff_StockTaking_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowTaking.Primary, PrimaryUpdate, Me.TblPrimary_Taking)
                Next

                count = 0
                For Each rowTaking In TblTaking.Rows
                    count += 1
                    TblTag.DefaultView.RowFilter = Nothing
                    TblTag.DefaultView.RowFilter = "[Primary Stock Taking] = " & rowTaking.Primary
                    If TblTag.DefaultView.Count > 0 Then
                        Me.SetProgressBarNLabel("", BatchProgBar, True, TblTag.DefaultView.Count)
                        For Each dv As DataRowView In TblTag.DefaultView
                            If Me.Bgw1.CancellationPending Then
                                Return False
                            End If

                            rowTag = dv.Row
                            Me.SetProgressBarNLabel("Transfer Inventory Check Tag Data (" & count & " of " & TblTaking.Rows.Count & ")", BatchProgBar)

                            PrimaryUpdate = Me.ObjCutOff_Core.GetPrimaryUpdate(rowTag.Primary_Stock_Taking, Me.TblPrimary_Taking)

                            Me.SetDefaultConnectionString()
                            FrmTag.FromCutOff_StockTag_Create(rowTag.Primary)
                            Me.SetTargetConnectionString()
                            FrmTag.FromCutOff_StockTag_Save(PrimaryUpdate)
                        Next
                    End If
                Next

                Me.SetDefaultConnectionString()

                'update data product qty dan simpan data ke stock card
                DtaPenyesuaian.Connection.ConnectionString = Me.SumberConstring
                DtaPenyesuaian.FillBy_OutCutOff(TblPenyesuaian, Me.EndDateCutOff)
                count = 0
                Me.SetTargetConnectionString()
                DtaPenyesuaian.Connection.ConnectionString = Me.TargetConstring
                For Each rowTaking In TblTaking.Rows
                    count += 1
                    TblPenyesuaian.DefaultView.RowFilter = "[Primary Stock Taking] = " & rowTaking.Primary
                    If TblPenyesuaian.DefaultView.Count > 0 Then
                        Me.SetProgressBarNLabel("", BatchProgBar, True, TblPenyesuaian.DefaultView.Count)
                        For Each dv As DataRowView In TblPenyesuaian.DefaultView
                            If Me.Bgw1.CancellationPending Then
                                Return False
                            End If

                            RowPenyesuaian = dv.Row
                            Me.SetProgressBarNLabel("Transfer End Stock Taking Procces (" & count & " of " & TblTaking.Rows.Count & ")", BatchProgBar)

                            With RowPenyesuaian
                                If .Qty <> 0 Then
                                    SaldoAkhir = 0

                                    If .Qty < 0 Then
                                        QtyIn = 0
                                        QtyOut = Math.Abs(.Qty)
                                    Else
                                        QtyIn = .Qty
                                        QtyOut = 0
                                    End If

                                    PrimaryUpdate = Me.ObjCutOff_Core.GetPrimaryUpdate(.Primary_Stock_Taking, Me.TblPrimary_Taking)
                                    DtaPenyesuaian.Insert(PrimaryUpdate, .Primary_Product, .Batch_No, .Expiry_Date, .Primary_Location, .Qty, .Qty_Secondary)

                                    'Update Stock Card
                                    ObjSC.Simpan(rowTaking.Primary, Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking, _
                                    rowTaking.No_Transaksi, .Primary_Product, QtyIn, QtyOut, "Stock Taking Adjustment", _
                                    Inventory_Module.ClsTransaksiStockCard.EnumJenisStockCard.Stock_Taking, rowTaking.Start_Date, DTASC, _
                                    RowPenyesuaian.Batch_No, RowPenyesuaian.Expiry_Date, RowPenyesuaian.Primary_Location, "")
                                End If
                            End With
                        Next
                    End If
                Next

                'simpan data stock taking laporan cost
                DtaLapCost.Connection.ConnectionString = Me.SumberConstring
                DtaLapCost.FillBy_OutCutOff(TblLapCost, Me.EndDateCutOff)
                count = 0
                Me.SetTargetConnectionString()
                For Each rowTaking In TblTaking.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    count += 1
                    'simpan data stock taking laporan cost
                    PrimaryUpdate = Me.ObjCutOff_Core.GetPrimaryUpdate(rowTaking.Primary, Me.TblPrimary_Taking)
                    TransferStockTaking_UsingScript(rowTaking.Primary, TblLapCost, count, TblTaking.Rows.Count, PrimaryUpdate)
                Next
            End If
            Me.RowUpdate("Transfer Stock Taking Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockTaking")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmTaking.Dispose()
            FrmTag.Dispose()
            DtaPenyesuaian.Dispose()
            TblPenyesuaian.Dispose()
            DtaLapCost.Dispose()
            TblLapCost.Dispose()
            DTASC.Dispose()
            DTAProdQty.Dispose()
            TBLProdQty.Dispose()
            DTAProductQty.Dispose()
            ObjSC = Nothing
            ObjProduct = Nothing
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferStockTaking_UsingScript(ByVal vPrimaryStockTaking As Double, _
    ByVal vTblLapCost As DseTrans_CutOffDB.tbltrans_stock_taking_laporancostDataTable, _
    ByVal vCount As Integer, ByVal vTotalCount As Integer, ByVal PrimaryTaking_new As Double) As Boolean
        Dim rowLapCost As DseTrans_CutOffDB.tbltrans_stock_taking_laporancostRow
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0
        Dim QueryInsert As String = ""
        Dim Value_ As String = ""
        Dim Delimiter As String = ","
        Dim BatchProgBar As Integer = 1

        Try
            'open connection
            Conn.Open()

            vTblLapCost.DefaultView.RowFilter = Nothing
            vTblLapCost.DefaultView.RowFilter = "[Primary Stock Taking] = " & vPrimaryStockTaking
            If vTblLapCost.DefaultView.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, vTblLapCost.DefaultView.Count)
                QueryInsert = "INSERT INTO `tbltrans_stock_taking_laporancost` (`primary closing`, `primary stock taking`, `primary product`, `qty on hand`, `qty taking`, `difference`, `primary unit`, `sub total`, `analisis`, `counter measures`) VALUES "
                While StartRowIndex < vTblLapCost.DefaultView.Count
                    Value_ = ""
                    Delimiter = ","
                    For i As Double = 0 To BatchQty - 1 Step 1
                        If Me.Bgw1.CancellationPending Then
                            Return False
                        End If

                        If vTblLapCost.DefaultView.Count > StartRowIndex + i Then
                            rowLapCost = vTblLapCost.DefaultView(i + StartRowIndex).Row
                            Me.SetProgressBarNLabel("Transfer Stock Taking Cost Report (" & vCount & " of " & vTotalCount & ")", BatchProgBar, True, vTblLapCost.DefaultView.Count)

                            'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                            If vTblLapCost.DefaultView.Count - 1 = StartRowIndex + i Or i = BatchQty - 1 Then
                                Delimiter = ";"
                            End If

                            With rowLapCost
                                Value_ = Value_ + vbCrLf + "(" + .primary_closing.ToString + "," _
                                + PrimaryTaking_new.ToString + "," + _
                                CDbl(.primary_product).ToString + "," + CDec(.qty_on_hand).ToString + "," _
                                + CDec(.qty_taking).ToString + "," _
                                + CDec(.difference).ToString + "," _
                                + CDbl(.primary_unit).ToString + "," _
                                + CDec(.sub_total).ToString + ",'" _
                                + .analisis.ToString + "','" _
                                + .counter_measures + "')" + Delimiter
                            End With

                        Else
                            Exit For
                        End If
                    Next

                    sql = QueryInsert & Value_ & vbCrLf & vbCrLf & "COMMIT;"
                    script.Query = sql
                    'execute script
                    script.Execute()
                    StartRowIndex += BatchQty
                End While
            End If

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockTaking_UsingScript")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Public Function TransferStockTransfer() As Boolean
        Dim DtaLocTrans As New DseTrans_CutOffDBTableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim TblLocTrans As New DseTrans_CutOffDB.tbltransaksi_stocktransferDataTable
        Dim rowLocTrans As DseTrans_CutOffDB.tbltransaksi_stocktransferRow
        Dim FrmLocTrans As New Inventory_Module.FrmTransaksiStockTransfer
        Dim BatchProgBar As Integer = 1

        Try
            DtaLocTrans.Connection.ConnectionString = Me.SumberConstring
            DtaLocTrans.FillBy_OutCutOff(TblLocTrans, Me.EndDateCutOff)

            Me.RowUpdate("Transfer Location Transfer Data", Me.StatusOnProses)
            ''Transfer semua stock transfer di luar period cutoff
            If TblLocTrans.Rows.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblLocTrans.Rows.Count)
                For Each rowLocTrans In TblLocTrans.Rows
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Location Transfer Data", BatchProgBar)

                    Me.SetDefaultConnectionString()
                    FrmLocTrans.FromCutOff_LocationTrans_Create(rowLocTrans.Primary)
                    Me.SetTargetConnectionString()
                    FrmLocTrans.FromCutOff_LocationTrans_Save()
                Next
            End If

            Me.RowUpdate("Transfer Location Transfer Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockTransfer")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmLocTrans.Dispose()
            DtaLocTrans.Dispose()
            TblLocTrans.Dispose()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferMP_SF() As Boolean
        Dim ListPrimary As New List(Of Double)
        Dim MP_SF As String = ""
        Dim comm As New MySql.Data.MySqlClient.MySqlCommand
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection

        Try
            For i As Integer = 1 To 2
                If i = 1 Then
                    MP_SF = "Transfer Material Planning Data"
                Else
                    MP_SF = "Transfer Shipping Forecast Data"
                End If

                Me.RowUpdate(MP_SF, Me.StatusOnProses)
                ListPrimary.Clear()
                Me.TblListTransfer.DefaultView.RowFilter = Nothing
                Me.TblListTransfer.DefaultView.RowFilter = "[Description] = '" & MP_SF.ToString & "'"
                For Each dv As DataRowView In Me.TblListTransfer.DefaultView
                    Me.RowListTransfer = dv.Row
                    Me.ProgressBar1.Value = 0
                    If Me.TransferMaster_UsingScript(Me.RowListTransfer.Tabel_Name, ListPrimary) = False Then
                        Return False
                    End If
                Next
                Me.RowUpdate(MP_SF, Me.StatusOK)
            Next

            Me.Label1.Text = "Update Material Planning Data. Please Wait ..."
            Conn.ConnectionString = Me.TargetConstring
            Conn.Open()
            comm.Connection = Conn
            comm.CommandText = "update tbltrans_material_planning_group_item_adjustment set `primary adj` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_group_item_bundling set `primary bundling` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_group_item_invoice set `primary invoice` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_group_item_sr set `primary sr` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_itemdo set `primary do rec` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_itemgit set `primary packing list rec` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_itempob set `primary po rec` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_itemsob set `primary so rec` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_material_planning_itempurch_req set `primary pr rec` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_shipping_forecast_item_pob set `primary po` = -1"
            comm.ExecuteNonQuery()
            comm.CommandText = "update tbltrans_shipping_forecast_itemdelivery set `primary so` = -1"
            comm.ExecuteNonQuery()
            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockTransfer")
            Return False
        Finally
            comm = Nothing
            Conn = Nothing
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    Public Function TransferMonthlyClosing() As Boolean
        Dim ListPrimary As New List(Of Double)
        Dim MP_SF As String = ""
        Dim Dta As New MySqlDataAdapter
        Dim tbl As New DataTable
        Dim row As DataRow
        Dim comm As New MySql.Data.MySqlClient.MySqlCommand
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(Me.TargetConstring)
        Dim ObjClosing As New ClsTransaksiMonthlyClosing
        Dim BatchProgBar As Integer = 1

        Try
            Me.RowUpdate("Transfer Monthly Closing Data", Me.StatusOnProses)
            ListPrimary.Clear()
            Me.TblListTransfer.DefaultView.RowFilter = Nothing
            Me.TblListTransfer.DefaultView.RowFilter = "[Description] = 'Transfer Monthly Closing Data'"
            For Each dv As DataRowView In Me.TblListTransfer.DefaultView
                Me.RowListTransfer = dv.Row
                Me.ProgressBar1.Value = 0
                If Me.TransferMaster_UsingScript(Me.RowListTransfer.Tabel_Name, ListPrimary) = False Then
                    Return False
                End If
            Next

            Dim ObjCL As New ClsMonthlyClosing01
            ObjCL.StartCutOffClosingPertama(Me.ObjCD.GetFirstDateofTheMonth(Me.EndDateCutOff), Me.TargetConstring, Me.SumberConstring, Me.ProgressBar1, Me.Label1)

            'Update Cost Invoice
            Me.SetTargetConnectionString()
            comm.Connection = Conn
            comm.CommandText = "select * from tbltrans_purchase_closing"
            Dta.SelectCommand = comm
            Dta.Fill(tbl)
            Me.SetProgressBarNLabel("", BatchProgBar, True, tbl.Rows.Count)
            For i As Integer = 0 To tbl.Rows.Count - 1
                row = tbl.Rows(i)
                Me.SetProgressBarNLabel("Purchase Closing Update Cost Invoice", BatchProgBar)

                ObjClosing.SimpanDataCost(row.Item("Primary"), row.Item("Closing Period"))
            Next

            Me.RowUpdate("Transfer Monthly Closing Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferStockTransfer")
            Return False
        Finally
            comm = Nothing
            Conn = Nothing
            Dta.Dispose()
            tbl.Dispose()
            ObjClosing = Nothing
            Me.SetDefaultConnectionString()
            Me.ProgressBar1.Value = 0
            Me.Label1.Text = ""
        End Try
    End Function

    'TRANSFER ACCOUNTING MODULE
    Public Function TransferGeneralJournal_CashBank() As Boolean
        Dim Dta As New MySqlDataAdapter
        Dim DtaJurnal As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnalTableAdapter
        Dim TblJurnal, Tbl_Item, Tbl_ItemPayDet, TblLoan, TblLoanPay, TblPay As New DataTable
        Dim rowJurnal As DataRow
        Dim cmd As New MySqlCommand
        Dim ConnSumber As New MySql.Data.MySqlClient.MySqlConnection(Me.SumberConstring)
        Dim NoJurnal As String = ""
        Dim PrimaryJurnalUpd As Double = -1
        Dim ListPrimaryJurnal As New List(Of Double)

        Try
            Me.RowUpdate("Transfer General Journal & Cash Bank Data", Me.StatusOnProses)
            cmd.Connection = ConnSumber
            cmd.CommandText = "select * from tbltrans_jurnal_item"
            Dta.SelectCommand = cmd
            Dta.Fill(Tbl_Item)
            Tbl_Item.Columns.RemoveAt(0)

            cmd.CommandText = "select * from tbltrans_jurnal_item_payment_detail"
            Dta.SelectCommand = cmd
            Dta.Fill(Tbl_ItemPayDet)
            Tbl_ItemPayDet.Columns.RemoveAt(0)

            If Me.TransferMaster_UsingScript("tbltrans_jurnal_quick_add") = False Then
                Return False
            End If
            If Me.TransferMaster_UsingScript("tbltrans_jurnal_quick_add_item") = False Then
                Return False
            End If
            If Me.TransferMaster_UsingScript("tbltrans_jurnal", ListPrimaryJurnal, "tbltrans_jurnal_item", Tbl_Item, True, True, "tbltrans_jurnal_item_payment_detail", Tbl_ItemPayDet) = False Then
                Return False
            End If

            'setelah transfer jurnal selesai maka update primary jurnal pada transaksi payment, loan dan loan payment
            ConnSumber = New MySql.Data.MySqlClient.MySqlConnection(Me.TargetConstring)
            cmd.Connection = ConnSumber
            'JURNAL
            cmd.CommandText = "select * from tbltrans_jurnal"
            Dta.SelectCommand = cmd
            Dta.Fill(TblJurnal)
            DtaJurnal.Connection.ConnectionString = Me.TargetConstring
            TblJurnal.DefaultView.RowFilter = "[Tipe Journal] = 0 AND ([Tipe Jurnal 2] = 8 OR [Tipe Jurnal 2] = 9 OR [Tipe Jurnal 2] = 10)"
            If TblJurnal.DefaultView.Count > 0 Then
                For Each dv As DataRowView In TblJurnal.DefaultView
                    rowJurnal = dv.Row
                    If rowJurnal.Item("Tipe Journal") = 0 And rowJurnal.Item("Tipe Jurnal 2") = 8 Then 'payment
                        DtaJurnal.UpdateQuery_PrimaryJurnalInPayment(rowJurnal.Item("Primary"), rowJurnal.Item("Primary Document Acuan"))
                    ElseIf rowJurnal.Item("Tipe Journal") = 0 And rowJurnal.Item("Tipe Jurnal 2") = 9 Then 'loan
                        DtaJurnal.UpdateQuery_PrimaryJurnalInLoan(rowJurnal.Item("Primary"), rowJurnal.Item("Primary Document Acuan"))
                    ElseIf rowJurnal.Item("Tipe Journal") = 0 And rowJurnal.Item("Tipe Jurnal 2") = 10 Then 'loan payment
                        DtaJurnal.UpdateQuery_PrimaryJurnalInLoanPay(rowJurnal.Item("Primary"), rowJurnal.Item("Primary Document Acuan"))
                    End If
                Next
            End If
            Me.RowUpdate("Transfer General Journal & Cash Bank Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferGeneralJournal")
            Return False
        Finally
            Dta.Dispose()
            TblJurnal.Dispose()
            Tbl_Item.Dispose()
            TblLoan.Dispose()
            TblLoanPay.Dispose()
            TblPay.Dispose()
            TblJurnal.Dispose()
            DtaJurnal.Dispose()
        End Try
    End Function

    Public Function TransferCreditDebitNote() As Boolean
        Dim DtaCD As New DseTrans_CutOffDBTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TblCD As New DseTrans_CutOffDB.tbltrans_credit_noteDataTable
        Dim rowCD As DseTrans_CutOffDB.tbltrans_credit_noteRow
        Dim FrmCD As New AccModule.FrmTransaksiCreditNote
        Dim PrimaryUpdate As Double = -1
        Dim BatchProgBar As Integer = 1

        Try
            DtaCD.Connection.ConnectionString = Me.SumberConstring
            DtaCD.FillBy_OutCutOff(TblCD, Me.EndDateCutOff)

            Me.RowUpdate("Transfer Credit & Debit Note Data", Me.StatusOnProses)
            If TblCD.DefaultView.Count > 0 Then
                Me.SetProgressBarNLabel("", BatchProgBar, True, TblCD.DefaultView.Count)
                For Each dv As DataRowView In TblCD.DefaultView
                    rowCD = dv.Row
                    If Me.Bgw1.CancellationPending Then
                        Return False
                    End If

                    Me.SetProgressBarNLabel("Transfer Credit Note Data", BatchProgBar)

                    'Get Primary AOP & Invoice New
                    If rowCD.primary_invoice <> -1 Then
                        If rowCD.tipe_kartu = AccModule.ClsTransaksiCreditNote.EnumTipeKartu.Customer Then
                            If rowCD.primary_invoice <> -1 Then
                                rowCD.primary_invoice = Me.ObjCutOff_Core.GetPrimaryUpdate(rowCD.primary_invoice, Me.TblPrimary_Inv)
                            End If
                        Else
                            If rowCD.primary_invoice <> -1 Then
                                rowCD.primary_invoice = Me.ObjCutOff_Core.GetPrimaryUpdate(rowCD.primary_invoice, Me.TblPrimary_AOP)
                            End If
                        End If
                    End If

                    Me.SetDefaultConnectionString()
                    FrmCD.FromCutOff_CreditDebitNote_Create(rowCD.primary, rowCD.primary_invoice)
                    Me.SetTargetConnectionString()
                    PrimaryUpdate = FrmCD.FromCutOff_CreditDebitNote_Save()
                    Me.ObjCutOff_Core.AddPrimaryUpdate(rowCD.primary, PrimaryUpdate, Me.TblPrimary_CD)
                Next
            End If
            Me.RowUpdate("Transfer Credit & Debit Note Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferCreditNote")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            FrmCD.Dispose()
            DtaCD.Dispose()
            TblCD.Dispose()
        End Try
    End Function

    Public Function TransferPurchasePayment_Old() As Boolean
        'Dim DtaJournal As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnalTableAdapter
        'Dim TblJournal_Purch As New DseTrans_CutOffDB.tbltrans_jurnalDataTable
        'Dim rowJournal_Purch As DseTrans_CutOffDB.tbltrans_jurnalRow
        'Dim DtaBalancePurch, DtaBalancePurch_New As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbltrans_balance_purchaseTableAdapter
        'Dim TblBalancePurch, TblBalancePurch_New As New Core_Module.DseTrans_CutOffDB_Core.tbltrans_balance_purchaseDataTable
        'Dim rowbalancePurch, rowbalancePurch_New As Core_Module.DseTrans_CutOffDB_Core.tbltrans_balance_purchaseRow

        'Dim FrmPPJ As New AccModule.FrmTransaksiJournal_Purchase
        'Dim BatchProgBar As Integer = 1

        'Try
        '    Me.RowUpdate("Transfer Purchase Payment Data", Me.StatusOnProses)
        '    Me.SetDefaultConnectionString()
        '    DtaBalancePurch.Connection.ConnectionString = Me.SumberConstring
        '    DtaBalancePurch.Fill(TblBalancePurch)
        '    DtaBalancePurch.Connection.ConnectionString = Me.TargetConstring
        '    DtaBalancePurch.Fill(TblBalancePurch_New)
        '    DtaJournal.Connection.ConnectionString = Me.SumberConstring
        '    DtaJournal.FillBy_OutCutOff(TblJournal_Purch, AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal, _
        '    AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping, Me.EndDateCutOff)

        '    Me.SetProgressBarNLabel("Transfer Purchase Payment Data", BatchProgBar, True, TblJournal_Purch.Count)
        '    For Each rowJournal_Purch In TblJournal_Purch.Rows
        '        If Me.Bgw1.CancellationPending Then
        '            Return False
        '        End If
        '        Me.SetProgressBarNLabel("Transfer Purchase Payment Data", BatchProgBar)

        '        Me.SetDefaultConnectionString()
        '        If FrmPPJ.FromCutOff_JurnalPurchase_Create(rowJournal_Purch.Primary) = True Then
        '            Me.SetTargetConnectionString()
        '            FrmPPJ.FromCutOff_JurnalPurchase_Save()
        '        End If
        '    Next

        '    Me.RowUpdate("Transfer Purchase Payment Data", Me.StatusOK)

        '    Return True
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        '    Me.SetDefaultConnectionString()
        '    Return False
        'Finally
        '    Me.SetDefaultConnectionString()
        '    FrmPPJ.Dispose()
        '    DtaBalancePurch.Dispose()
        '    TblBalancePurch.Dispose()
        '    TblBalancePurch_New.Dispose()
        '    DtaJournal.Dispose()
        '    TblJournal_Purch.Dispose()
        'End Try
    End Function

    Public Function TransferPurchasePayment() As Boolean
        Dim DtaJurnal As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal1TableAdapter
        Dim DtaJurnal_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal_Item1TableAdapter
        Dim DtaJurnal_ItemOther As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal_ItemOther1TableAdapter
        Dim DtaJurnal_ItemPay As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal_ItemPayment1TableAdapter
        
        Dim DtaCDNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_CDNoTableAdapter
        Dim TblCDNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoDataTable

        Dim TblJurnal As New DseTrans_CutOffDB.tbltrans_jurnal1DataTable
        Dim TblJurnal_Item As New DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable
        Dim TblJurnal_ItemOther As New DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable
        Dim TblJurnal_ItemPay As New DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable
        
        Dim DtaAPAR As New AccModule.DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter

        Try
            Me.SetDefaultConnectionString()
            Me.Label1.Text = "Get Purchase Payment Data, Please Wait..."
            Me.RowUpdate("Transfer Purchase Payment Journal Data", Me.StatusOnProses)
            DtaJurnal.Connection.ConnectionString = Me.SumberConstring
            DtaJurnal_Item.Connection.ConnectionString = Me.SumberConstring
            DtaJurnal_ItemOther.Connection.ConnectionString = Me.SumberConstring
            DtaJurnal_ItemPay.Connection.ConnectionString = Me.SumberConstring

            DtaJurnal.Fill(TblJurnal, AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping, Me.EndDateCutOff)
            DtaJurnal_Item.FillBy_OutCutOff(TblJurnal_Item, AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping, Me.EndDateCutOff)
            DtaJurnal_ItemOther.FillBy_OutCutOff(TblJurnal_ItemOther, Me.EndDateCutOff, AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping)
            DtaJurnal_ItemPay.FillBy_OutCutOff(TblJurnal_ItemPay, Me.EndDateCutOff, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping, AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal)
            DtaCDNo.Fill(TblCDNo, CDate("1/1/1950"))

            Me.TransferSalesReceivement_UsingScript(TblJurnal, TblJurnal_Item, TblJurnal_ItemOther, TblJurnal_ItemPay, TblCDNo, DtaAPAR, AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal)

            Me.RowUpdate("Transfer Purchase Payment Journal Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesReceivement")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaJurnal.Dispose()
            DtaJurnal_Item.Dispose()
            DtaJurnal_ItemOther.Dispose()
            DtaJurnal_ItemPay.Dispose()
            TblJurnal.Dispose()
            TblJurnal_Item.Dispose()
            TblJurnal_ItemOther.Dispose()
            TblJurnal_ItemPay.Dispose()
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Public Function TransferSalesReceivement() As Boolean
        Dim DtaJurnal As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal1TableAdapter
        Dim DtaJurnal_Item As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal_Item1TableAdapter
        Dim DtaJurnal_ItemOther As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal_ItemOther1TableAdapter
        Dim DtaJurnal_ItemPay As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnal_ItemPayment1TableAdapter
       
        Dim DtaCDNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_CDNoTableAdapter
        Dim TblCDNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoDataTable

        Dim TblJurnal As New DseTrans_CutOffDB.tbltrans_jurnal1DataTable
        Dim TblJurnal_Item As New DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable
        Dim TblJurnal_ItemOther As New DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable
        Dim TblJurnal_ItemPay As New DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable
        
        Dim DtaAPAR As New AccModule.DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter

        Try
            Me.SetDefaultConnectionString()
            Me.Label1.Text = "Get Sales Receivement Data, Please Wait..."
            Me.RowUpdate("Transfer Sales Receivement Journal Data", Me.StatusOnProses)
            DtaJurnal.Connection.ConnectionString = Me.SumberConstring
            DtaJurnal_Item.Connection.ConnectionString = Me.SumberConstring
            DtaJurnal_ItemOther.Connection.ConnectionString = Me.SumberConstring
            DtaJurnal_ItemPay.Connection.ConnectionString = Me.SumberConstring
            
            DtaJurnal.Fill(TblJurnal, AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping, Me.EndDateCutOff)
            DtaJurnal_Item.FillBy_OutCutOff(TblJurnal_Item, AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping, Me.EndDateCutOff)
            DtaJurnal_ItemOther.FillBy_OutCutOff(TblJurnal_ItemOther, Me.EndDateCutOff, AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping)
            DtaJurnal_ItemPay.FillBy_OutCutOff(TblJurnal_ItemPay, Me.EndDateCutOff, AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping, AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal)
            DtaCDNo.Fill(TblCDNo, CDate("1/1/1950"))
            
            Me.TransferSalesReceivement_UsingScript(TblJurnal, TblJurnal_Item, TblJurnal_ItemOther, TblJurnal_ItemPay, TblCDNo, DtaAPAR, AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal)

            Me.RowUpdate("Transfer Sales Receivement Journal Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesReceivement")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            DtaJurnal.Dispose()
            DtaJurnal_Item.Dispose()
            DtaJurnal_ItemOther.Dispose()
            DtaJurnal_ItemPay.Dispose()
            TblJurnal.Dispose()
            TblJurnal_Item.Dispose()
            TblJurnal_ItemOther.Dispose()
            TblJurnal_ItemPay.Dispose()
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Private Function TransferSalesReceivement_UsingScript(ByVal TblJurnal As DseTrans_CutOffDB.tbltrans_jurnal1DataTable, _
    ByVal TblJurnal_Item As DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable, _
    ByVal TblJurnal_ItemOther As DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable, _
    ByVal TblJurnal_ItemPay As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable, _
    ByVal TblCDNo As Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoDataTable, _
    ByVal DtaAPAR As AccModule.DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter, _
    ByVal vTipe As AccModule.ClsTransaksiJournal.EnumTipeJournal) As Boolean
        'purchase payment
        Dim DtaBalancePurch, DtaBalancePurch_New As New DseTrans_CutOffDBTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TblBalancePurch, TblBalancePurch_New As New DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable
        Dim DTABalancePurchUpdt As New AccModule.DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim DtaHistoricalPurch As New DseTrans_CutOffDBTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim TblHistoricalPurch As New DseTrans_CutOffDB.tbltrans_historicalpurchasebalancemanualDataTable
        Dim DtaHistoricalPR As New DseTrans_CutOffDBTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim TblHistoricalPR As New DseTrans_CutOffDB.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim DtaAopNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbltrans_aopTableAdapter
        Dim TblAopNo As New Core_Module.DseTrans_CutOffDB_Core.tbltrans_aopDataTable
        Dim DtaPRNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TblPRNo As New Core_Module.DseTrans_CutOffDB_Core.tbltrans_purchasereturnDataTable

        'sales receivement
        Dim DtaBalanceSales, DtaBalanceSales_New As New DseTrans_CutOffDBTableAdapters.tbltrans_balance_sales1TableAdapter
        Dim TblBalanceSales, TblBalanceSales_New As New DseTrans_CutOffDB.tbltrans_balance_sales1DataTable
        Dim DTABalanceSalesUpdt As New AccModule.DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DtaHistoricalSales As New DseTrans_CutOffDBTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim TblHistoricalSales As New DseTrans_CutOffDB.tbltrans_historicalsalesDataTable
        Dim DtaHistoricalSR As New DseTrans_CutOffDBTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TblHistoricalSR As New DseTrans_CutOffDB.tbltrans_historicalsalesreturn_manualDataTable
        Dim DtaInvNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_InvNoTableAdapter
        Dim TblInvNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoDataTable
        Dim DtaSRNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_SRNoTableAdapter
        Dim TblSRNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_SRNoDataTable

        Dim TblJurnal_Item_Filter As New DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable
        Dim TblJurnal_ItemOther_Filter As New DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable
        Dim TblJurnal_ItemPay_Filter As New DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable
        Dim TblBalancePurch_Filter, TblBalancePurch_New_Filter As New DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable
        Dim TblBalanceSales_Filter, TblBalanceSales_New_Filter As New DseTrans_CutOffDB.tbltrans_balance_sales1DataTable
        Dim TblAOPNo_Filter As New Core_Module.DseTrans_CutOffDB_Core.tbltrans_aopDataTable
        Dim TblInvNo_Filter As New Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoDataTable
        Dim RowJurnal As DseTrans_CutOffDB.tbltrans_jurnal1Row

        Dim cmd As New MySqlCommand
        Dim ConnTarget As New MySql.Data.MySqlClient.MySqlConnection(Me.TargetConstring)
        Dim tblDummy, tblDummy2, tblDummy3, tblDummy4 As New Core_Module.dsFormat.tblFormatDataTable
        Dim rowsCommand As MySqlCommand = ConnTarget.CreateCommand()
        Dim rowsReader As MySqlDataReader

        Dim PrimaryPihakAcuan As Double = -1
        Dim PrimaryJurnalMax As Double = -1
        Dim NamaKolom_Jurnal As String = ""
        Dim NamaKolom_JurnalItem As String = ""
        Dim NamaKolom_JurnalItemOther As String = ""
        Dim NamaKolom_JurnalItemPay As String = ""
        Dim Value_, ValueInsert_, ValueInsertItem_, ValueInsertItemOther_, ValueInsertItemPay_ As String
        Dim QueryInsert, QueryInsert_Item, QueryInsert_ItemOther, QueryInsert_ItemPay As String
        Dim Sql As String = ""
        Dim Script As New MySql.Data.MySqlClient.MySqlScript(ConnTarget, Sql)
        Dim StartRowIndex As Double = 0
        Dim BatchQty As Double = 100
        Dim ListPrimaryBalanceSales As New List(Of Double)
        Dim ObjJurnalItem As New AccModule.ClsTransaksiJournal_Item

        Try
            If vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                DtaBalancePurch.Connection.ConnectionString = Me.SumberConstring
                DtaBalancePurch_New.Connection.ConnectionString = Me.TargetConstring
                DtaHistoricalPurch.Connection.ConnectionString = Me.SumberConstring
                DtaHistoricalPR.Connection.ConnectionString = Me.SumberConstring
                DtaBalancePurch.Fill(TblBalancePurch)
                DtaBalancePurch_New.Fill(TblBalancePurch_New)
                DtaAopNo.ClearBeforeFill = False
                DtaAopNo.Fill(TblAopNo, CDate("1/1/1950"))
                DtaPRNo.Fill(TblPRNo, CDate("1/1/1950"))
                DtaHistoricalPurch.Fill(TblHistoricalPurch)
                DtaHistoricalPR.Fill(TblHistoricalPR)
            ElseIf vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                DtaBalanceSales.Connection.ConnectionString = Me.SumberConstring
                DtaBalanceSales_New.Connection.ConnectionString = Me.TargetConstring
                DtaHistoricalSales.Connection.ConnectionString = Me.SumberConstring
                DtaHistoricalSR.Connection.ConnectionString = Me.SumberConstring
                DtaInvNo.ClearBeforeFill = False
                DtaInvNo.FillBy_OutCutOff(TblInvNo, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP, CDate("1/1/1950"))
                DtaInvNo.FillBy_OutCutOff(TblInvNo, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice, CDate("1/1/1950"))
                DtaSRNo.Fill(TblSRNo, CDate("1/1/1950"))
                DtaHistoricalSales.Fill_All(TblHistoricalSales)
                DtaHistoricalSR.Fill_All(TblHistoricalSR)
                DtaBalanceSales.Fill(TblBalanceSales)
                DtaBalanceSales_New.Fill(TblBalanceSales_New)
            End If

            PrimaryJurnalMax = Me.ObjCutOff.GetPrimaryMaxJurnal(Me.TargetConstring)

            ConnTarget.Open()
            'dapatkan nama kolom jurnal
            tblDummy.Rows.Clear()
            rowsCommand.CommandText = "DESCRIBE tbltrans_jurnal"
            rowsReader = rowsCommand.ExecuteReader()
            For i As Integer = 1 To TblJurnal.Columns.Count - 1
                rowsReader.Read()
                If rowsReader(0).ToString <> TblJurnal.Columns(i).ColumnName Then
                    rowsReader.Read()
                End If
                tblDummy.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, i)
                If i <> TblJurnal.Columns.Count - 1 Then
                    NamaKolom_Jurnal &= "`" & TblJurnal.Columns(i).ColumnName & "`" & ","
                Else
                    NamaKolom_Jurnal &= "`" & TblJurnal.Columns(i).ColumnName & "`"
                End If
            Next
            ConnTarget.Close()

            ConnTarget.Open()
            'dapatkan nama kolom jurnal item
            tblDummy2.Rows.Clear()
            rowsCommand.CommandText = "DESCRIBE tbltrans_jurnal_item"
            rowsReader = rowsCommand.ExecuteReader()
            For i As Integer = 1 To TblJurnal_Item.Columns.Count - 2
                rowsReader.Read()
                If rowsReader(0).ToString <> TblJurnal_Item.Columns(i).ColumnName Then
                    rowsReader.Read()
                End If
                tblDummy2.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, i)
                If i <> TblJurnal_Item.Columns.Count - 2 Then
                    NamaKolom_JurnalItem &= "`" & TblJurnal_Item.Columns(i).ColumnName & "`" & ","
                Else
                    NamaKolom_JurnalItem &= "`" & TblJurnal_Item.Columns(i).ColumnName & "`"
                End If
            Next
            ConnTarget.Close()

            ConnTarget.Open()
            'dapatkan nama kolom jurnal item other
            tblDummy3.Rows.Clear()
            rowsCommand.CommandText = "DESCRIBE tbltrans_jurnal_item_other"
            rowsReader = rowsCommand.ExecuteReader()
            For i As Integer = 1 To TblJurnal_ItemOther.Columns.Count - 2
                rowsReader.Read()
                If rowsReader(0).ToString <> TblJurnal_ItemOther.Columns(i).ColumnName Then
                    rowsReader.Read()
                End If
                tblDummy3.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, i)
                If i <> TblJurnal_ItemOther.Columns.Count - 2 Then
                    NamaKolom_JurnalItemOther &= "`" & TblJurnal_ItemOther.Columns(i).ColumnName & "`" & ","
                Else
                    NamaKolom_JurnalItemOther &= "`" & TblJurnal_ItemOther.Columns(i).ColumnName & "`"
                End If
            Next
            ConnTarget.Close()

            ConnTarget.Open()
            'dapatkan nama kolom jurnal item payment
            tblDummy4.Rows.Clear()
            rowsCommand.CommandText = "DESCRIBE tbltrans_jurnal_item_payment"
            rowsReader = rowsCommand.ExecuteReader()
            For i As Integer = 1 To TblJurnal_ItemPay.Columns.Count - 2
                rowsReader.Read()
                If rowsReader(0).ToString <> TblJurnal_ItemPay.Columns(i).ColumnName Then
                    rowsReader.Read()
                End If
                tblDummy4.Rows.Add(rowsReader.GetString("Type").ToString.ToUpper, i)
                If i <> TblJurnal_ItemPay.Columns.Count - 2 Then
                    NamaKolom_JurnalItemPay &= "`" & TblJurnal_ItemPay.Columns(i).ColumnName & "`" & ","
                Else
                    NamaKolom_JurnalItemPay &= "`" & TblJurnal_ItemPay.Columns(i).ColumnName & "`"
                End If
            Next
            ConnTarget.Close()

            ConnTarget.Open()
            Me.SetProgressBarNLabel("", BatchQty, True, TblJurnal.Rows.Count)
            'dapatkan value dari tiap kolom
            Me.SetTargetConnectionString()
            While StartRowIndex < TblJurnal.Rows.Count
                ValueInsert_ = ""
                ValueInsertItem_ = ""
                ValueInsertItemOther_ = ""
                ValueInsertItemPay_ = ""

                For i As Double = 0 To BatchQty - 1 Step 1
                    If TblJurnal.Rows.Count > StartRowIndex + i Then
                        If Me.Bgw1.CancellationPending Then
                            Return False
                        End If

                        Value_ = ""
                        RowJurnal = TblJurnal.Rows(StartRowIndex + i)
                        For i2 As Integer = 1 To TblJurnal.Columns.Count - 1
                            If i2 <> 1 Then
                                Value_ &= ","
                            End If
                            Value_ &= Me.FrmBackup.FormatValue(tblDummy, (i2 - 1), TblJurnal.Rows(TblJurnal.Rows.IndexOf(RowJurnal))(i2))
                        Next
                        ValueInsert_ &= "(" & Value_ & "), "

                        PrimaryJurnalMax += 1

                        'filter data berdasarkan primary pihak acuan/primary customer
                        If PrimaryPihakAcuan <> RowJurnal.Primary_Pihak_Acuan Then
                            PrimaryPihakAcuan = RowJurnal.Primary_Pihak_Acuan
                            If vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                                Me.ObjCutOff.PurchasePayment_FilterData(TblJurnal_Item, TblJurnal_Item_Filter, TblJurnal_ItemOther, _
                                TblJurnal_ItemOther_Filter, TblJurnal_ItemPay, TblJurnal_ItemPay_Filter, TblBalancePurch, _
                                TblBalancePurch_Filter, TblBalancePurch_New, TblBalancePurch_New_Filter, TblAopNo, TblAOPNo_Filter, _
                                PrimaryPihakAcuan)
                            ElseIf vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                                Me.ObjCutOff.SalesReceivement_FilterData(TblJurnal_Item, TblJurnal_Item_Filter, TblJurnal_ItemOther, _
                                TblJurnal_ItemOther_Filter, TblJurnal_ItemPay, TblJurnal_ItemPay_Filter, TblBalanceSales, _
                                TblBalanceSales_Filter, TblBalanceSales_New, TblBalanceSales_New_Filter, TblInvNo, TblInvNo_Filter, _
                                PrimaryPihakAcuan)
                            End If
                        End If

                        'Jurnal Item 
                        Me.TransferSalesReceivement_UsingScriptItem(TblJurnal_Item_Filter, RowJurnal.Primary, PrimaryJurnalMax, ValueInsertItem_, tblDummy2)
                        'Jurnal Item Other
                        Me.TransferSalesReceivement_UsingScriptItemOther(TblJurnal_ItemOther_Filter, RowJurnal.Primary, PrimaryJurnalMax, ValueInsertItemOther_, tblDummy3)
                        If vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                            'Jurnal Item Payment (purchase payment)
                            Me.TransferPurchasePayment_UsingScriptItemPayment(TblJurnal_ItemPay_Filter, TblBalancePurch_Filter, TblBalancePurch_New_Filter, _
                            TblAOPNo_Filter, TblPRNo, TblCDNo, TblJurnal, DTABalancePurchUpdt, DtaAPAR, TblHistoricalPurch, TblHistoricalPR, RowJurnal, PrimaryJurnalMax, ValueInsertItemPay_, _
                            ListPrimaryBalanceSales, ObjJurnalItem, tblDummy4)
                        ElseIf vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                            'Jurnal Item Payment (sales receivement)
                            Me.TransferSalesReceivement_UsingScriptItemPayment(TblJurnal_ItemPay_Filter, TblBalanceSales_Filter, TblBalanceSales_New_Filter, _
                            TblInvNo_Filter, TblSRNo, TblCDNo, TblJurnal, DTABalanceSalesUpdt, DtaAPAR, TblHistoricalSales, TblHistoricalSR, RowJurnal, PrimaryJurnalMax, ValueInsertItemPay_, _
                            ListPrimaryBalanceSales, ObjJurnalItem, tblDummy4)
                        End If
                    Else
                        Exit For
                    End If
                Next

                'Execute Jurnal
                ValueInsert_ = ValueInsert_.Remove(ValueInsert_.Length - 2, 2) & ";"
                QueryInsert = "INSERT INTO tbltrans_jurnal (" + NamaKolom_Jurnal + ") VALUES " & vbCrLf & ValueInsert_ & vbCrLf & vbCrLf & "COMMIT;"
                Sql = QueryInsert
                Script.Query = Sql
                Script.Execute()

                'Execute Jurnal Item
                If ValueInsertItem_ <> "" Then
                    ValueInsertItem_ = ValueInsertItem_.Remove(ValueInsertItem_.Length - 2, 2) & ";"
                    QueryInsert_Item = "INSERT INTO tbltrans_jurnal_item (" + NamaKolom_JurnalItem + ") VALUES " & vbCrLf & ValueInsertItem_ & vbCrLf & vbCrLf & "COMMIT;"
                    Sql = QueryInsert_Item
                    Script.Query = Sql
                    Script.Execute()
                End If

                'Execute Jurnal Item Other
                If ValueInsertItemOther_ <> "" Then
                    ValueInsertItemOther_ = ValueInsertItemOther_.Remove(ValueInsertItemOther_.Length - 2, 2) & ";"
                    QueryInsert_ItemOther = "INSERT INTO tbltrans_jurnal_item_other (" + NamaKolom_JurnalItemOther + ") VALUES " & vbCrLf & ValueInsertItemOther_ & vbCrLf & vbCrLf & "COMMIT;"
                    Sql = QueryInsert_ItemOther
                    Script.Query = Sql
                    Script.Execute()
                End If

                'Execute Jurnal Item Payment
                If ValueInsertItemPay_ <> "" Then
                    ValueInsertItemPay_ = ValueInsertItemPay_.Remove(ValueInsertItemPay_.Length - 2, 2) & ";"
                    QueryInsert_ItemPay = "INSERT INTO tbltrans_jurnal_item_payment (" + NamaKolom_JurnalItemPay + ") VALUES " & vbCrLf & ValueInsertItemPay_ & vbCrLf & vbCrLf & "COMMIT;"
                    Sql = QueryInsert_ItemPay
                    Script.Query = Sql
                    Script.Execute()
                End If

                If vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                    Me.SetProgressBarNLabel("Transfer Purchase Payment Journal Data", BatchQty)
                ElseIf vTipe = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                    Me.SetProgressBarNLabel("Transfer Sales Receivement Journal Data", BatchQty)
                End If

                StartRowIndex += BatchQty
            End While

            ConnTarget.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesReceivement_UsingScript")
            Return False
        Finally
            ConnTarget = Nothing
            Script = Nothing
            ObjJurnalItem = Nothing
            DtaBalanceSales.Dispose()
            DtaBalanceSales_New.Dispose()
            TblBalanceSales.Dispose()
            TblBalanceSales_New.Dispose()
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Private Function TransferSalesReceivement_UsingScriptItem(ByVal TblJurnal_Item As DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable, _
    ByVal vPrimaryJurnal As Double, ByVal vPrimaryJurnalMax As Double, ByRef vValueItem As String, _
    ByVal tblDummy2 As Core_Module.dsFormat.tblFormatDataTable) As Boolean
        Dim Row2 As DataRow
        Dim Value_ As String = ""

        Try
            TblJurnal_Item.DefaultView.RowFilter = "[Primary Jurnal] = " & vPrimaryJurnal
            If TblJurnal_Item.DefaultView.Count > 0 Then
                For ii As Integer = 0 To TblJurnal_Item.DefaultView.Count - 1
                    Value_ = ""
                    Row2 = TblJurnal_Item.DefaultView(ii).Row
                    For i As Integer = 1 To TblJurnal_Item.Columns.Count - 2
                        If i <> 1 Then
                            Value_ &= ","
                        End If

                        If TblJurnal_Item.Columns(i).ColumnName = "Primary Jurnal" Then
                            Value_ &= vPrimaryJurnalMax
                        Else
                            Value_ &= Me.FrmBackup.FormatValue(tblDummy2, (i - 1), TblJurnal_Item.Rows(TblJurnal_Item.Rows.IndexOf(Row2))(i))
                        End If
                    Next
                    vValueItem &= "(" & Value_ & "), "
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesReceivement_UsingScriptItem")
            Return False
        End Try
    End Function

    Private Function TransferSalesReceivement_UsingScriptItemOther(ByVal TblJurnal_ItemOther As DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable, _
    ByVal vPrimaryJurnal As Double, ByVal vPrimaryJurnalMax As Double, ByRef vValueItem_Other As String, _
    ByVal tblDummy3 As Core_Module.dsFormat.tblFormatDataTable) As Boolean
        Dim Row2 As DataRow
        Dim Value_ As String = ""

        Try
            TblJurnal_ItemOther.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimaryJurnal
            If TblJurnal_ItemOther.DefaultView.Count > 0 Then
                For ii As Integer = 0 To TblJurnal_ItemOther.DefaultView.Count - 1
                    Value_ = ""
                    Row2 = TblJurnal_ItemOther.DefaultView(ii).Row
                    For i As Integer = 1 To TblJurnal_ItemOther.Columns.Count - 2
                        If i <> 1 Then
                            Value_ &= ","
                        End If

                        If TblJurnal_ItemOther.Columns(i).ColumnName = "primary transaksi" Then
                            Value_ &= vPrimaryJurnalMax
                        Else
                            Value_ &= Me.FrmBackup.FormatValue(tblDummy3, (i - 1), TblJurnal_ItemOther.Rows(TblJurnal_ItemOther.Rows.IndexOf(Row2))(i))
                        End If
                    Next

                    vValueItem_Other &= "(" & Value_ & "), "
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster_UsingScript")
            Return False
        End Try
    End Function

    Private Function TransferPurchasePayment_UsingScriptItemPayment(ByVal TblJurnal_ItemPay As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable, _
    ByVal TblBalancePurch As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByVal TblBalancePurch_New As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByVal TblAopNo As Core_Module.DseTrans_CutOffDB_Core.tbltrans_aopDataTable, _
    ByVal TblPRNo As Core_Module.DseTrans_CutOffDB_Core.tbltrans_purchasereturnDataTable, _
    ByVal TblCDNo As Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoDataTable, _
    ByVal TblJurnal As DseTrans_CutOffDB.tbltrans_jurnal1DataTable, _
    ByVal DTABalanceUpdt As AccModule.DseTransaksiPurchaseBalanceTableAdapters.tbltrans_balance_purchaseTableAdapter, _
    ByVal DtaAPAR As AccModule.DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter, _
    ByVal TblHistoricalPurch As DseTrans_CutOffDB.tbltrans_historicalpurchasebalancemanualDataTable, _
    ByVal TblHistoricalPR As DseTrans_CutOffDB.tbltrans_historicalpurchasereturnbalancemanualDataTable, _
    ByVal RowJurnal As DseTrans_CutOffDB.tbltrans_jurnal1Row, ByVal vPrimaryJurnalMax As Double, _
    ByRef vValueItem_Pay As String, ByRef vListPrimaryBalancePurch As List(Of Double), _
    ByVal ObjJurnalItem As AccModule.ClsTransaksiJournal_Item, _
    ByVal tblDummy4 As Core_Module.dsFormat.tblFormatDataTable) As Boolean
        Dim rowJurnal_ItemPay As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1Row
        Dim rowBalancePurch As DseTrans_CutOffDB.tbltrans_balance_purchaseRow

        Dim PrimaryDocAcuanNew_ As Double = -1
        Dim Balance_ As Decimal = 0
        Dim BalanceOriginal_ As Decimal = 0
        Dim Received_ As Decimal = 0
        Dim ReceivedOriginal_ As Decimal = 0
        Dim PrimaryBalancePurchNew_ As Double = -1
        Dim PrimaryCurr_ As Double = -1
        Dim NoTransAcuan As String = ""
        Dim TglAcuan As Date = "1/1/1950"
        Dim PaymentTermAcuan As Integer = 0
        Dim Value_ As String = ""

        Try
            TblJurnal_ItemPay.DefaultView.RowFilter = "[Primary Journal] = " & RowJurnal.Primary
            If TblJurnal_ItemPay.DefaultView.Count > 0 Then
                For ii As Integer = 0 To TblJurnal_ItemPay.DefaultView.Count - 1
                    Value_ = ""
                    rowJurnal_ItemPay = TblJurnal_ItemPay.DefaultView(ii).Row

                    TblBalancePurch.DefaultView.RowFilter = "[Primary Supplier] = " & rowJurnal_ItemPay.Primary_Pihak_Acuan & _
                    " AND [Primary AOP] = " & rowJurnal_ItemPay.Primary_Document_Acuan & " AND [Tipe] = " & _
                    rowJurnal_ItemPay.Tipe_Document_Acuan
                    If TblBalancePurch.DefaultView.Count > 0 Then
                        rowBalancePurch = TblBalancePurch.DefaultView(0).Row
                        Balance_ = rowBalancePurch.Balance
                        BalanceOriginal_ = rowBalancePurch.balance_original
                        Received_ = rowBalancePurch.Paid
                        ReceivedOriginal_ = rowBalancePurch.paid_original
                        PrimaryCurr_ = rowBalancePurch.primary_currency
                    Else
                        Balance_ = 0
                        BalanceOriginal_ = 0
                        Received_ = 0
                        ReceivedOriginal_ = 0
                        PrimaryCurr_ = -1
                    End If

                    NoTransAcuan = ""
                    TglAcuan = "1/1/1950"
                    PaymentTermAcuan = 0
                    PrimaryDocAcuanNew_ = -1
                    PrimaryBalancePurchNew_ = -1
                    'Update Primary Document Acuan, Tipe Document Acuan, Primary Balance Sales
                    If rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                        PrimaryDocAcuanNew_ = ObjCutOff_Core.GetPrimaryUpdate(rowJurnal_ItemPay.Primary_Document_Acuan, Me.TblPrimary_AOP)
                        If PrimaryDocAcuanNew_ < 0 Then
                            If Balance_ < 0 Then
                                GoTo PurchBalanceReturnManual
                            End If
                            GoTo PurchBalanceManual
                        Else
                            NoTransAcuan = ObjCutOff_Core.GetAopNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblAopNo, TglAcuan, PaymentTermAcuan)
                            PrimaryBalancePurchNew_ = Me.ObjCutOff.GetPrimaryBalanceOrAcuanPurch_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, PrimaryDocAcuanNew_, _
                            AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic, TblBalancePurch_New)
                            rowJurnal_ItemPay.Primary_Document_Acuan = PrimaryDocAcuanNew_
                        End If
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
PurchBalanceManual:
                        If rowJurnal_ItemPay.Tipe_Document_Acuan <> AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                            NoTransAcuan = Me.ObjCutOff_Core.GetAopNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblAopNo, TglAcuan, PaymentTermAcuan)
                        Else
                            NoTransAcuan = Me.ObjCutOff.GetNoHistoricalPurch(TblHistoricalPurch, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan)
                        End If
                        rowJurnal_ItemPay.Primary_Document_Acuan = Me.ObjCutOff.GetPrimaryBalanceOrAcuanPurch_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                        AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual, TblBalancePurch_New, NoTransAcuan, PrimaryBalancePurchNew_)
                        rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                        If PrimaryBalancePurchNew_ < 0 Then
                            GoTo PurchBalanceReturnManual2
                        End If
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                        PrimaryDocAcuanNew_ = ObjCutOff_Core.GetPrimaryUpdate(rowJurnal_ItemPay.Primary_Document_Acuan, TblPrimary_PR)
                        If PrimaryDocAcuanNew_ < 0 Then
                            GoTo PurchBalanceReturnManual
                        Else
                            NoTransAcuan = ObjCutOff_Core.GetPurchReturnNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblPRNo, TglAcuan, PaymentTermAcuan)
                            PrimaryBalancePurchNew_ = Me.ObjCutOff.GetPrimaryBalanceOrAcuanPurch_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, PrimaryDocAcuanNew_, _
                            AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR, TblBalancePurch_New)
                            rowJurnal_ItemPay.Primary_Document_Acuan = PrimaryDocAcuanNew_
                        End If
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
PurchBalanceReturnManual:
                        If rowJurnal_ItemPay.Tipe_Document_Acuan <> AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                            NoTransAcuan = ObjCutOff_Core.GetPurchReturnNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblPRNo, TglAcuan, PaymentTermAcuan)
                        Else
                            NoTransAcuan = ObjCutOff.GetNoHistoricalPurchReturn(TblHistoricalPR, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan)
                        End If
PurchBalanceReturnManual2:
                        rowJurnal_ItemPay.Primary_Document_Acuan = Me.ObjCutOff.GetPrimaryBalanceOrAcuanPurch_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                        AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR, TblBalancePurch_New, NoTransAcuan, PrimaryBalancePurchNew_)
                        rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Or _
                    rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                        NoTransAcuan = ObjCutOff_Core.GetCDNoteNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblCDNo, TglAcuan)
                        PrimaryDocAcuanNew_ = ObjCutOff_Core.GetPrimaryUpdate(rowJurnal_ItemPay.Primary_Document_Acuan, TblPrimary_CD)
                        If PrimaryDocAcuanNew_ < 0 Then
                            NoTransAcuan = Me.ObjCutOff_Core.GetCDNoteNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblCDNo, TglAcuan, PaymentTermAcuan)
                            PrimaryDocAcuanNew_ = Me.ObjCutOff.GetPrimaryBalanceOrAcuanPurch_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                            rowJurnal_ItemPay.Tipe_Document_Acuan, TblBalancePurch_New, NoTransAcuan, PrimaryBalancePurchNew_)
                            If rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                                rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                            Else
                                rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                            End If
                            rowJurnal_ItemPay.Primary_Document_Acuan = PrimaryDocAcuanNew_
                        Else
                            rowJurnal_ItemPay.Primary_Document_Acuan = Me.ObjCutOff.GetPrimaryBalanceOrAcuanPurch_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                            rowJurnal_ItemPay.Tipe_Document_Acuan, TblBalancePurch_New, NoTransAcuan, PrimaryBalancePurchNew_)
                        End If
                    End If

                    'Masukan item payment ke dalam list insert query
                    For i As Integer = 1 To TblJurnal_ItemPay.Columns.Count - 2
                        If i <> 1 Then
                            Value_ &= ","
                        End If

                        If TblJurnal_ItemPay.Columns(i).ColumnName = "Primary Journal" Then
                            Value_ &= vPrimaryJurnalMax
                        Else
                            Value_ &= Me.FrmBackup.FormatValue(tblDummy4, (i - 1), TblJurnal_ItemPay.Rows(TblJurnal_ItemPay.Rows.IndexOf(rowJurnal_ItemPay))(i))
                        End If
                    Next

                    If NoTransAcuan = "" Or TglAcuan = "1/1/1950" Then
                        MsgBox("stop")
                    End If

                    'Update Data Balance Sales
                    If vListPrimaryBalancePurch.Contains(PrimaryBalancePurchNew_) = False Then
                        vListPrimaryBalancePurch.Add(PrimaryBalancePurchNew_)
                        DTABalanceUpdt.UpdateQuery(Balance_, Received_, ReceivedOriginal_, BalanceOriginal_, PrimaryBalancePurchNew_, _
                        PrimaryCurr_, rowJurnal_ItemPay.Tipe_Document_Acuan)
                    End If

                    'Simpan Kartu AP & AR
                    If rowJurnal_ItemPay.Payment > 0 Then
                        'untuk pembayaran sales invoice
                        DtaAPAR.Insert(RowJurnal.Primary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment, _
                        rowJurnal_ItemPay.Primary_Pihak_Acuan, RowJurnal.Ref_Number, rowJurnal_ItemPay.Payment, 0, NoTransAcuan, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        RowJurnal._Date, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan, ObjJurnalItem.GetTipeTransaksiAsalPurchasePaymentJurnal(rowJurnal_ItemPay.Tipe_Document_Acuan), _
                        rowJurnal_ItemPay.primary_currency, rowJurnal_ItemPay.exchange_rate, rowJurnal_ItemPay.payment_original, 0, rowJurnal_ItemPay.primary_akun_ap_ar)
                    Else
                        'untuk pembayaran sales return
                        DtaAPAR.Insert(RowJurnal.Primary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Purchase_Payment, _
                        rowJurnal_ItemPay.Primary_Pihak_Acuan, RowJurnal.Ref_Number, 0, Math.Abs(rowJurnal_ItemPay.Payment), NoTransAcuan, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        RowJurnal._Date, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan, ObjJurnalItem.GetTipeTransaksiAsalPurchasePaymentJurnal(rowJurnal_ItemPay.Tipe_Document_Acuan), _
                        rowJurnal_ItemPay.primary_currency, rowJurnal_ItemPay.exchange_rate, 0, Math.Abs(rowJurnal_ItemPay.payment_original), rowJurnal_ItemPay.primary_akun_ap_ar)
                    End If

                    vValueItem_Pay &= "(" & Value_ & "), "
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster_UsingScript")
            Return False
        End Try
    End Function

    Private Function TransferSalesReceivement_UsingScriptItemPayment(ByVal TblJurnal_ItemPay As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable, _
    ByVal TblBalanceSales As DseTrans_CutOffDB.tbltrans_balance_sales1DataTable, _
    ByVal TblBalanceSales_New As DseTrans_CutOffDB.tbltrans_balance_sales1DataTable, _
    ByVal TblInvNo As Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoDataTable, _
    ByVal TblSRNo As Core_Module.DseTrans_CutOffDB_Core.tbl_SRNoDataTable, _
    ByVal TblCDNo As Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoDataTable, _
    ByVal TblJurnal As DseTrans_CutOffDB.tbltrans_jurnal1DataTable, _
    ByVal DTABalanceUpdt As AccModule.DseTransaksiSalesBalanceTableAdapters.tbltrans_balance_salesTableAdapter, _
    ByVal DtaAPAR As AccModule.DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter, _
    ByVal TblHistoricalSales As DseTrans_CutOffDB.tbltrans_historicalsalesDataTable, _
    ByVal TblHistoricalSR As DseTrans_CutOffDB.tbltrans_historicalsalesreturn_manualDataTable, _
    ByVal RowJurnal As DseTrans_CutOffDB.tbltrans_jurnal1Row, ByVal vPrimaryJurnalMax As Double, _
    ByRef vValueItem_Pay As String, ByRef vListPrimaryBalanceSales As List(Of Double), _
    ByVal ObjJurnalItem As AccModule.ClsTransaksiJournal_Item, _
    ByVal tblDummy4 As Core_Module.dsFormat.tblFormatDataTable) As Boolean
        Dim rowJurnal_ItemPay As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1Row
        Dim rowBalanceSales As DseTrans_CutOffDB.tbltrans_balance_sales1Row

        Dim PrimaryDocAcuanNew_ As Double = -1
        Dim Balance_ As Decimal = 0
        Dim BalanceOriginal_ As Decimal = 0
        Dim Received_ As Decimal = 0
        Dim ReceivedOriginal_ As Decimal = 0
        Dim PrimaryBalanceSalesNew_ As Double = -1
        Dim PrimaryCurr_ As Double = -1
        Dim NoTransAcuan As String = ""
        Dim TglAcuan As Date = "1/1/1950"
        Dim PaymentTermAcuan As Integer = 0
        Dim Value_ As String = ""

        Try
            TblJurnal_ItemPay.DefaultView.RowFilter = "[Primary Journal] = " & RowJurnal.Primary
            If TblJurnal_ItemPay.DefaultView.Count > 0 Then
                For ii As Integer = 0 To TblJurnal_ItemPay.DefaultView.Count - 1
                    Value_ = ""
                    rowJurnal_ItemPay = TblJurnal_ItemPay.DefaultView(ii).Row

                    TblBalanceSales.DefaultView.RowFilter = "[Primary Customer] = " & rowJurnal_ItemPay.Primary_Pihak_Acuan & _
                    " AND [Primary Invoice] = " & rowJurnal_ItemPay.Primary_Document_Acuan & " AND [Tipe] = " & _
                    rowJurnal_ItemPay.Tipe_Document_Acuan
                    If TblBalanceSales.DefaultView.Count > 0 Then
                        rowBalanceSales = TblBalanceSales.DefaultView(0).Row
                        Balance_ = rowBalanceSales.Balance
                        BalanceOriginal_ = rowBalanceSales.balance_original
                        Received_ = rowBalanceSales.Received
                        ReceivedOriginal_ = rowBalanceSales.received_original
                        PrimaryCurr_ = rowBalanceSales.primary_currency
                    Else
                        Balance_ = 0
                        BalanceOriginal_ = 0
                        Received_ = 0
                        ReceivedOriginal_ = 0
                        PrimaryCurr_ = -1
                    End If

                    NoTransAcuan = ""
                    TglAcuan = "1/1/1950"
                    PaymentTermAcuan = 0
                    PrimaryDocAcuanNew_ = -1
                    PrimaryBalanceSalesNew_ = -1
                    'Update Primary Document Acuan, Tipe Document Acuan, Primary Balance Sales
                    If rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                        PrimaryDocAcuanNew_ = ObjCutOff_Core.GetPrimaryUpdate(rowJurnal_ItemPay.Primary_Document_Acuan, Me.TblPrimary_Inv)
                        If PrimaryDocAcuanNew_ < 0 Then
                            If Balance_ < 0 Then
                                GoTo SalesBalanceReturnManual
                            End If
                            GoTo SalesBalanceManual
                        Else
                            NoTransAcuan = ObjCutOff_Core.GetInvNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblInvNo, TglAcuan, PaymentTermAcuan)
                            PrimaryBalanceSalesNew_ = Me.ObjCutOff.GetPrimaryBalanceOrAcuanSales_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, PrimaryDocAcuanNew_, _
                            AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic, TblBalanceSales_New)
                            rowJurnal_ItemPay.Primary_Document_Acuan = PrimaryDocAcuanNew_
                        End If
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
SalesBalanceManual:
                        If rowJurnal_ItemPay.Tipe_Document_Acuan <> AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                            NoTransAcuan = Me.ObjCutOff_Core.GetInvNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblInvNo, TglAcuan, PaymentTermAcuan)
                        Else
                            NoTransAcuan = Me.ObjCutOff.GetNoHistoricalSales(TblHistoricalSales, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan)
                        End If
                        rowJurnal_ItemPay.Primary_Document_Acuan = Me.ObjCutOff.GetPrimaryBalanceOrAcuanSales_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                        AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual, TblBalanceSales_New, NoTransAcuan, PrimaryBalanceSalesNew_)
                        rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                        If PrimaryBalanceSalesNew_ < 0 Then
                            GoTo SalesBalanceReturnManual2
                        End If
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                        PrimaryDocAcuanNew_ = ObjCutOff_Core.GetPrimaryUpdate(rowJurnal_ItemPay.Primary_Document_Acuan, TblPrimary_SR)
                        If PrimaryDocAcuanNew_ < 0 Then
                            GoTo SalesBalanceReturnManual
                        Else
                            NoTransAcuan = ObjCutOff_Core.GetSalesReturnNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblSRNo, TglAcuan, PaymentTermAcuan)
                            PrimaryBalanceSalesNew_ = Me.ObjCutOff.GetPrimaryBalanceOrAcuanSales_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, PrimaryDocAcuanNew_, _
                            AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR, TblBalanceSales_New)
                            rowJurnal_ItemPay.Primary_Document_Acuan = PrimaryDocAcuanNew_
                        End If
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
SalesBalanceReturnManual:
                        If rowJurnal_ItemPay.Tipe_Document_Acuan <> AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                            NoTransAcuan = ObjCutOff_Core.GetSalesReturnNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblSRNo, TglAcuan, PaymentTermAcuan)
                        Else
                            NoTransAcuan = ObjCutOff.GetNoHistoricalSalesReturn(TblHistoricalSR, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan)
                        End If
SalesBalanceReturnManual2:
                        rowJurnal_ItemPay.Primary_Document_Acuan = Me.ObjCutOff.GetPrimaryBalanceOrAcuanSales_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                        AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR, TblBalanceSales_New, NoTransAcuan, PrimaryBalanceSalesNew_)
                        rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                    ElseIf rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Or _
                    rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                        NoTransAcuan = ObjCutOff_Core.GetCDNoteNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblCDNo, TglAcuan)
                        PrimaryDocAcuanNew_ = ObjCutOff_Core.GetPrimaryUpdate(rowJurnal_ItemPay.Primary_Document_Acuan, TblPrimary_CD)
                        If PrimaryDocAcuanNew_ < 0 Then
                            NoTransAcuan = ObjCutOff_Core.GetCDNoteNo(rowJurnal_ItemPay.Primary_Document_Acuan, TblCDNo, TglAcuan, PaymentTermAcuan)
                            PrimaryDocAcuanNew_ = Me.ObjCutOff.GetPrimaryBalanceOrAcuanSales_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                            rowJurnal_ItemPay.Tipe_Document_Acuan, TblBalanceSales_New, NoTransAcuan, PrimaryBalanceSalesNew_)
                            If rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Then
                                rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual
                            Else
                                rowJurnal_ItemPay.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR
                            End If
                            rowJurnal_ItemPay.Primary_Document_Acuan = PrimaryDocAcuanNew_
                        Else
                            rowJurnal_ItemPay.Primary_Document_Acuan = Me.ObjCutOff.GetPrimaryBalanceOrAcuanSales_New(rowJurnal_ItemPay.Primary_Pihak_Acuan, -1, _
                            rowJurnal_ItemPay.Tipe_Document_Acuan, TblBalanceSales_New, NoTransAcuan, PrimaryBalanceSalesNew_)
                        End If
                    End If

                    'Masukan item payment ke dalam list insert query
                    For i As Integer = 1 To TblJurnal_ItemPay.Columns.Count - 2
                        If i <> 1 Then
                            Value_ &= ","
                        End If

                        If TblJurnal_ItemPay.Columns(i).ColumnName = "Primary Journal" Then
                            Value_ &= vPrimaryJurnalMax
                        Else
                            Value_ &= Me.FrmBackup.FormatValue(tblDummy4, (i - 1), TblJurnal_ItemPay.Rows(TblJurnal_ItemPay.Rows.IndexOf(rowJurnal_ItemPay))(i))
                        End If
                    Next

                    If NoTransAcuan = "" Then
                        MsgBox("stop")
                    End If

                    'Update Data Balance Sales
                    If vListPrimaryBalanceSales.Contains(PrimaryBalanceSalesNew_) = False Then
                        vListPrimaryBalanceSales.Add(PrimaryBalanceSalesNew_)
                        DTABalanceUpdt.UpdateQuery(Balance_, Received_, ReceivedOriginal_, BalanceOriginal_, PrimaryBalanceSalesNew_, _
                        PrimaryCurr_, rowJurnal_ItemPay.Tipe_Document_Acuan)
                    End If

                    'Simpan Kartu AP & AR
                    If rowJurnal_ItemPay.Payment > 0 Then
                        'untuk pembayaran sales invoice
                        DtaAPAR.Insert(RowJurnal.Primary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
                        rowJurnal_ItemPay.Primary_Pihak_Acuan, RowJurnal.Ref_Number, 0, rowJurnal_ItemPay.Payment, NoTransAcuan, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                        RowJurnal._Date, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan, ObjJurnalItem.GetTipeTransaksiAsalSalesReceivementJurnal(rowJurnal_ItemPay.Tipe_Document_Acuan), _
                        rowJurnal_ItemPay.primary_currency, rowJurnal_ItemPay.exchange_rate, 0, rowJurnal_ItemPay.payment_original, rowJurnal_ItemPay.primary_akun_ap_ar)
                    Else
                        'untuk pembayaran sales return
                        DtaAPAR.Insert(RowJurnal.Primary, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
                        rowJurnal_ItemPay.Primary_Pihak_Acuan, RowJurnal.Ref_Number, Math.Abs(rowJurnal_ItemPay.Payment), 0, NoTransAcuan, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                        RowJurnal._Date, rowJurnal_ItemPay.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan, ObjJurnalItem.GetTipeTransaksiAsalSalesReceivementJurnal(rowJurnal_ItemPay.Tipe_Document_Acuan), _
                        rowJurnal_ItemPay.primary_currency, rowJurnal_ItemPay.exchange_rate, Math.Abs(rowJurnal_ItemPay.payment_original), 0, rowJurnal_ItemPay.primary_akun_ap_ar)
                    End If

                    vValueItem_Pay &= "(" & Value_ & "), "
                Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster_UsingScript")
            Return False
        End Try
    End Function

    Public Function TransferCollectionLetter() As Boolean
        Dim Dta_Item As New MySqlDataAdapter
        Dim Tbl_Item As New DataTable
        Dim cmd As New MySqlCommand
        Dim ConnSumber As New MySql.Data.MySqlClient.MySqlConnection(Me.SumberConstring)
        Dim ListPrimary As New List(Of Double)

        Try
            'Delete baris jika invoice dalam period cutoff 
            Me.RowUpdate("Transfer Collection Letter Data", Me.StatusOnProses)
            cmd.Connection = ConnSumber
            cmd.CommandText = "select * from tbltrans_surattagihan_item"
            Dta_Item.SelectCommand = cmd
            Dta_Item.Fill(Tbl_Item)
            Tbl_Item.Columns.RemoveAt(0)
            Me.TransferMaster_UsingScript("tbltrans_surattagihan", ListPrimary, "tbltrans_surattagihan_item", Tbl_Item, True, True)
            Me.RowUpdate("Transfer Collection Letter Data", Me.StatusOK)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.SetDefaultConnectionString()
            Return False
        Finally
            Me.SetDefaultConnectionString()
            Dta_Item.Dispose()
            Tbl_Item.Dispose()
        End Try
    End Function

    Private Function TransferSalesReceivement_UsingScriptItemPayment_Repair() As Boolean
        Dim DtaJurnal As New DseTrans_CutOffDBTableAdapters.DataTable1TableAdapter
        Dim TblJurnal As New DseTrans_CutOffDB.DataTable1DataTable
        Dim rowJurnal As DseTrans_CutOffDB.DataTable1Row

        Dim DtaHistoricalSales As New DseTrans_CutOffDBTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim TblHistoricalSales As New DseTrans_CutOffDB.tbltrans_historicalsalesDataTable

        Dim DtaHistoricalSR As New DseTrans_CutOffDBTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TblHistoricalSR As New DseTrans_CutOffDB.tbltrans_historicalsalesreturn_manualDataTable

        Dim DtaInvNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_InvNoTableAdapter
        Dim TblInvNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoDataTable

        Dim DtaSRNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_SRNoTableAdapter
        Dim TblSRNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_SRNoDataTable

        Dim DtaCDNo As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.tbl_CDNoTableAdapter
        Dim TblCDNo As New Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoDataTable

        Dim DtaAPAR As New AccModule.DseTrans_Kartu_APARTableAdapters.tbltransaksi_kartu_ar_apTableAdapter

        Dim NoTransAcuan As String = ""
        Dim TglAcuan As Date = "1/1/1950"
        Dim PaymentTermAcuan As Integer = 0
        Dim ObjJurnalItem As New AccModule.ClsTransaksiJournal_Item

        Try
            'untuk sementara perbaikan kartu ar ap
            Me.SetTargetConnectionString()
            DtaInvNo.ClearBeforeFill = False
            DtaInvNo.FillBy_OutCutOff(TblInvNo, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.InvoiceDP, CDate("1/1/1950"))
            DtaInvNo.FillBy_OutCutOff(TblInvNo, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice, CDate("1/1/1950"))
            DtaSRNo.Fill(TblSRNo, CDate("1/1/1950"))
            DtaCDNo.Fill(TblCDNo, CDate("1/1/1950"))

            DtaJurnal.Connection.ConnectionString = Me.TargetConstring
            DtaHistoricalSales.Connection.ConnectionString = Me.TargetConstring
            DtaHistoricalSR.Connection.ConnectionString = Me.TargetConstring
            DtaJurnal.Fillby_TipeJurnal(TblJurnal, AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal)
            DtaHistoricalSales.Fill_All(TblHistoricalSales)
            DtaHistoricalSR.Fill_All(TblHistoricalSR)

            For Each rowJurnal In TblJurnal.Rows
                NoTransAcuan = ""
                TglAcuan = "1/1/1950"
                PaymentTermAcuan = 0

                If rowJurnal.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic Then
                    rowJurnal.document1 = ObjCutOff_Core.GetInvNo(rowJurnal.Primary_Document_Acuan, TblInvNo, TglAcuan, PaymentTermAcuan)
                ElseIf rowJurnal.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual Then
                    rowJurnal.document1 = ObjCutOff.GetNoHistoricalSales(TblHistoricalSales, rowJurnal.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan)
                ElseIf rowJurnal.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic_SR Then
                    rowJurnal.document1 = ObjCutOff_Core.GetSalesReturnNo(rowJurnal.Primary_Document_Acuan, TblSRNo, TglAcuan, PaymentTermAcuan)
                ElseIf rowJurnal.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Manual_SR Then
                    rowJurnal.document1 = ObjCutOff.GetNoHistoricalSalesReturn(TblHistoricalSR, rowJurnal.Primary_Document_Acuan, TglAcuan)
                ElseIf rowJurnal.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note Or _
                rowJurnal.Tipe_Document_Acuan = AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note Then
                    rowJurnal.document1 = ObjCutOff_Core.GetCDNoteNo(rowJurnal.Primary_Document_Acuan, TblCDNo, TglAcuan)
                Else
                    MsgBox("stop")
                End If

                If rowJurnal.document1 = "" Then
                    MsgBox("stop")
                End If
                'Simpan Kartu AP & AR
                If rowJurnal.Payment > 0 Then
                    'untuk pembayaran sales invoice
                    DtaAPAR.Insert(rowJurnal.Primary_Journal, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
                    rowJurnal.Primary_Pihak_Acuan, rowJurnal.Ref_Number, 0, rowJurnal.Payment, rowJurnal.document1, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                    rowJurnal._Date, rowJurnal.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan, ObjJurnalItem.GetTipeTransaksiAsalSalesReceivementJurnal(rowJurnal.Tipe_Document_Acuan), _
                    rowJurnal.primary_currency, rowJurnal.exchange_rate, 0, rowJurnal.payment_original, rowJurnal.primary_akun_ap_ar)
                Else
                    'untuk pembayaran sales return
                    DtaAPAR.Insert(rowJurnal.Primary_Journal, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Sales_Receivement, _
                    rowJurnal.Primary_Pihak_Acuan, rowJurnal.Ref_Number, Math.Abs(rowJurnal.Payment), 0, rowJurnal.document1, ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                    rowJurnal._Date, rowJurnal.Primary_Document_Acuan, TglAcuan, PaymentTermAcuan, ObjJurnalItem.GetTipeTransaksiAsalSalesReceivementJurnal(rowJurnal.Tipe_Document_Acuan), _
                    rowJurnal.primary_currency, rowJurnal.exchange_rate, Math.Abs(rowJurnal.payment_original), 0, rowJurnal.primary_akun_ap_ar)
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferMaster_UsingScript")
            Return False
        End Try
    End Function

    '----------------------------------------
    Public Function StockCardVSProductQty(ByVal vEnumSetConstring As EnumProdQtyNStockCard) As Boolean
        Dim Dta As New Core_Module.DseTrans_CutOffDB_CoreTableAdapters.Tbl_PerbandinganQtyTableAdapter
        Dim Tbl, TblShow As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PerbandinganQtyDataTable
        Dim row, rowShow As Core_Module.DseTrans_CutOffDB_Core.Tbl_PerbandinganQtyRow
        Dim FrmExc As New Core_Module.FrmTrans_CutOffDB_ExceptionQty

        Try
            If vEnumSetConstring = EnumProdQtyNStockCard.SetSumber Then
                Me.SetDefaultConnectionString()
            Else
                Me.SetTargetConnectionString()
            End If

            Dta.Fill(Tbl)
            For Each row In Tbl.Rows
                If row.Total_Prod_Qty <> row.Total_SC_Qty Then
                    rowShow = TblShow.NewTbl_PerbandinganQtyRow
                    For i As Integer = 0 To TblShow.Columns.Count - 1
                        rowShow.Item(i) = row.Item(i)
                    Next
                    TblShow.Rows.Add(rowShow)
                End If
            Next

            If TblShow.Rows.Count > 0 Then
                FrmExc.TblShow = TblShow
                FrmExc.BackColorAll = Me.BackColorAll
                FrmExc.ShowDialog()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            TblShow.Dispose()
            FrmExc.Dispose()
            Me.SetDefaultConnectionString()
        End Try
    End Function

    Private Sub TruncateAllTable()
        Dim Dta As New MySqlDataAdapter
        Dim Tbl As New DataTable
        Dim Row As DataRow

        Dim cmdTarget As New MySqlCommand
        Dim ConnTarget As New MySqlConnection(Me.TargetConstring)
        Dim BatchProgBar As Integer = 1

        Try
            cmdTarget.Connection = ConnTarget
            cmdTarget.CommandText = "SHOW FULL TABLES WHERE Table_type <> 'VIEW'"
            Dta.SelectCommand = cmdTarget
            Dta.Fill(Tbl)

            Me.SetProgressBarNLabel("Clear Data In Database Cut Off", BatchProgBar, True, Tbl.Rows.Count)
            ConnTarget.Open()
            For Each Row In Tbl.Rows
                Me.SetProgressBarNLabel("Clear Data In Database Cut Off", BatchProgBar)
                cmdTarget.CommandText = "TRUNCATE TABLE " & Row.Item("tables_in_" & Me.Temporary_DB)
                cmdTarget.ExecuteNonQuery()
            Next
            ConnTarget.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            cmdTarget = Nothing
        End Try
    End Sub

    Private Sub TrancuteNTransfer(ByVal vTblName As String)
        Dim comm As New MySql.Data.MySqlClient.MySqlCommand
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection

        Try
            Conn.ConnectionString = Me.TargetConstring
            Conn.Open()
            comm.Connection = Conn
            comm.CommandText = "TRUNCATE TABLE " & vTblName
            comm.ExecuteNonQuery()
            Conn.Close()

            Me.TransferMaster_UsingScript(vTblName)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            comm = Nothing
        End Try
    End Sub

    Private Sub SetProgressBarNLabel(ByVal vDesc As String, ByVal vBatchProgBar As Integer, _
    Optional ByVal vSetDefaultProgBar As Boolean = False, Optional ByVal vCountMax As Integer = 0)

        Try
            If vSetDefaultProgBar = True Then
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = vCountMax
            Else
                If Me.ProgressBar1.Value + vBatchProgBar > Me.ProgressBar1.Maximum Then
                    Me.ProgressBar1.Value = Me.ProgressBar1.Maximum
                Else
                    Me.ProgressBar1.Value += vBatchProgBar
                End If
                Me.Label1.Text = vDesc & " " & ObjCD.DigitGrouping(Math.Round((Me.ProgressBar1.Value / Me.ProgressBar1.Maximum) * 100, 2, _
                MidpointRounding.AwayFromZero), 2) + "% - Data " + ObjCD.DigitGrouping(Me.ProgressBar1.Value) + " of " + ObjCD.DigitGrouping(Me.ProgressBar1.Maximum)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini di gunakan saat proses cutoff belum stabil dan sering terjadi out of memory
    Private Sub GetTblPrimaryUpdate()
        Dim Dta As New MySqlDataAdapter
        Dim vArrayTbl() As String = {"a_so", "a_do", "a_sr", "a_cd", "a_aop", "a_sisr", "a_pr", "a_inadj", "a_ouadj", "a_taking", "a_inv", "a_po", "a_sip", "a_grn"}

        Try
            Dim cmd As New MySqlCommand
            Dim ConnSumber As New MySql.Data.MySqlClient.MySqlConnection(Me.TargetConstring)
            cmd.Connection = ConnSumber

            For Each TblName As String In vArrayTbl
                Dim Tbl As New DataTable
                cmd.CommandText = "select * from " & TblName
                Dta.SelectCommand = cmd
                Dta.Fill(Tbl)
                If TblName = "a_inv" Then
                    Me.TblPrimary_Inv.Clear()
                    Me.TblPrimary_Inv.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_Inv.Columns.RemoveAt(2)
                    Me.TblPrimary_Inv.Columns.RemoveAt(2)
                ElseIf TblName = "a_aop" Then
                    Me.TblPrimary_AOP.Clear()
                    Me.TblPrimary_AOP.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_AOP.Columns.RemoveAt(2)
                    Me.TblPrimary_AOP.Columns.RemoveAt(2)
                ElseIf TblName = "a_po" Then
                    Me.TblPrimary_PO.Clear()
                    Me.TblPrimary_PO.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_PO.Columns.RemoveAt(2)
                    Me.TblPrimary_PO.Columns.RemoveAt(2)
                ElseIf TblName = "a_sisr" Then
                    Me.TblPrimary_SISR.Clear()
                    Me.TblPrimary_SISR.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_SISR.Columns.RemoveAt(2)
                    Me.TblPrimary_SISR.Columns.RemoveAt(2)
                ElseIf TblName = "a_pr" Then
                    Me.TblPrimary_PR.Clear()
                    Me.TblPrimary_PR.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_PR.Columns.RemoveAt(2)
                    Me.TblPrimary_PR.Columns.RemoveAt(2)
                ElseIf TblName = "a_inadj" Then
                    Me.TblPrimary_InAdj.Clear()
                    Me.TblPrimary_InAdj.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_InAdj.Columns.RemoveAt(2)
                    Me.TblPrimary_InAdj.Columns.RemoveAt(2)
                ElseIf TblName = "a_ouadj" Then
                    Me.TblPrimary_OutAdj.Clear()
                    Me.TblPrimary_OutAdj.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_OutAdj.Columns.RemoveAt(2)
                    Me.TblPrimary_OutAdj.Columns.RemoveAt(2)
                ElseIf TblName = "a_taking" Then
                    Me.TblPrimary_Taking.Clear()
                    Me.TblPrimary_Taking.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_Taking.Columns.RemoveAt(2)
                    Me.TblPrimary_Taking.Columns.RemoveAt(2)
                ElseIf TblName = "a_cd" Then
                    Me.TblPrimary_CD.Clear()
                    Me.TblPrimary_CD.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_CD.Columns.RemoveAt(2)
                    Me.TblPrimary_CD.Columns.RemoveAt(2)
                ElseIf TblName = "a_sr" Then
                    Me.TblPrimary_SR.Clear()
                    Me.TblPrimary_SR.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_SR.Columns.RemoveAt(2)
                    Me.TblPrimary_SR.Columns.RemoveAt(2)
                ElseIf TblName = "a_do" Then
                    Me.TblPrimary_DO.Clear()
                    Me.TblPrimary_DO.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_DO.Columns.RemoveAt(2)
                    Me.TblPrimary_DO.Columns.RemoveAt(2)
                ElseIf TblName = "a_so" Then
                    Me.TblPrimary_SO.Clear()
                    Me.TblPrimary_SO.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_SO.Columns.RemoveAt(2)
                    Me.TblPrimary_SO.Columns.RemoveAt(2)
                ElseIf TblName = "a_sip" Then
                    Me.TblPrimary_SIP.Clear()
                    Me.TblPrimary_SIP.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_SIP.Columns.RemoveAt(2)
                    Me.TblPrimary_SIP.Columns.RemoveAt(2)
                ElseIf TblName = "a_grn" Then
                    Me.TblPrimary_GRN.Clear()
                    Me.TblPrimary_GRN.Merge(Tbl.DefaultView.Table.Copy)
                    Me.TblPrimary_GRN.Columns.RemoveAt(2)
                    Me.TblPrimary_GRN.Columns.RemoveAt(2)
                End If
                Tbl.Dispose()
            Next

            cmd.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSalesCancellation")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    Dim status_status As String = "default"
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.Temporary_DB = Me.TxtDB.Text & "_temporary"

            Core_Module.SetConstring(Me.TargetConstring)
            If Me.status_status <> "default" Then
                Me.status_status = "default"
                Me.SumberConstring = "server=fitri-pc;user id=root;Password=123;persist security info=True;database=" & Me.TxtDB.Text & ";port=3307;default command timeout=19000000"
                Me.TargetConstring = "server=fitri-pc;user id=root;Password=123;persist security info=True;database=" & Me.Temporary_DB & ";port=3307;default command timeout=19000000"
                Me.SetDefaultConnectionString()
            Else
                Me.status_status = "target"
                Me.SumberConstring = "server=fitri-pc;user id=root;Password=123;persist security info=True;database=" & Me.TxtDB.Text & ";port=3307;default command timeout=19000000"
                Me.TargetConstring = "server=fitri-pc;user id=root;Password=123;persist security info=True;database=" & Me.Temporary_DB & ";port=3307;default command timeout=19000000"
                Me.SetTargetConnectionString()
            End If

            Me.LinkLabel1.Text = Me.status_status & " | " & Core_Module.GetDBName

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class