Option Explicit On
Option Strict Off

Imports Core_Module
Imports System.IO
Imports System.Environment

Public Class FrmMain
    Public EW As E_mail_Module.ClsEmailWatcher
    Dim ObjComapanyData As New Core_Module.ClsMasterCompanyData
    Dim CloseForm As Boolean = False
    Dim ObjLicense As New ClsLicense
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjHakAkses As New Core_Module.ClsHakAkses
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjClosing As New Purchase_Module.ClsTransaksiMonthlyClosing
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjMonitor As New Core_Module.ClsTransaksiMonitor
    Dim Identifier As Double = -1
    Dim ObjRandom As New Random

    ''' <summary>
    ''' melakukan daily activities yang dilakukan software sales ketika pertama kali
    ''' ada yang login
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DoFirstActivitySales()
        Dim DTA As New Sales_Module.dseTransActivitiesTableAdapters.tbltrans_activitiesTableAdapter
        Dim rnd As New Random
        Dim Str As String = CStr(rnd.Next(0, 50000))
        Dim JmlInsert As Double = 0
        Dim Primary, PrimaryMin As Double
        Dim lenght As Double = 0
        Dim Email_Dummy As String = ""

        Dim DTA1 As New Sales_Module.dseTransActivitiesTableAdapters.tbltrans_so_item2TableAdapter
        Dim TBLSOBalance As New Sales_Module.dseTransActivities.tbltrans_so_item2DataTable
        Dim rowSOBalance As Sales_Module.dseTransActivities.tbltrans_so_item2Row
        Dim BookingNo, SONumber, UserCode, EmailContent, Filter, DUmmyCustomer As String
        Dim DataView As DataRowView
        Dim Counter As Double = 0
        Dim ObjEmail As New E_mail_Module.ClsTransEmail
        Dim SendTo As New List(Of String)
        Dim Ketemu As Boolean = False
        DUmmyCustomer = ""

        Dim ObjBookingItem As New Sales_Module.ClsTransaksiBookingInvItem
        Dim DTA2 As New Sales_Module.dseTransActivitiesTableAdapters.tbltrans_booking_invitem_view0TableAdapter
        Dim TBL2 As New Sales_Module.dseTransActivities.tbltrans_booking_invitem_view0DataTable
        Dim rowBooking As Sales_Module.dseTransActivities.tbltrans_booking_invitem_view0Row

        Dim ijk As Integer = 0

        Try
            'cek jangan sampe dilakukan 2 kali
            DTA.Insert(1, Today, Str)
            If DTA.ScalarQuery_GetJml(1, Today).HasValue = True Then
                'tandai hari ini
                JmlInsert = DTA.ScalarQuery_GetJml(1, Today).Value
            Else
                JmlInsert = 0
            End If
            If JmlInsert > 0 Then
                'test apakah yang bersangkutan yang masuk duluan
                PrimaryMin = DTA.ScalarQuery_GetMinPrimary(1, Today).Value
                Primary = DTA.ScalarQuery_GetPrimary(1, Today, Str).Value
                If Primary <> PrimaryMin Then
                    DTA.Delete(CDec(Primary))
                    Exit Sub
                End If
            End If

            'yang dilakukan adalah:
            If MsgBox("Is today's date is : " + _
            Format(Today, "dd MMMMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            '1. kirimkan email untuk sales order yang jatuh tempo
            DTA1.Fill(TBLSOBalance, Today + TimeSpan.FromDays(1))
            While TBLSOBalance.Rows.Count > 0
                rowSOBalance = TryCast(TBLSOBalance.Rows(0), Sales_Module.dseTransActivities.tbltrans_so_item2Row)
                SONumber = rowSOBalance.SO_No
                UserCode = rowSOBalance.Created_By
                Counter = 0
                SendTo.Clear()

                'filter berdasarkan user
                Filter = "[Created By] = '" + UserCode + "'"
                TBLSOBalance.DefaultView.RowFilter = Filter

                'buat email
                EmailContent = "The following item are promised to be delivered tomorrow:" + vbCrLf
                For Each DataView In TBLSOBalance.DefaultView
                    Counter += 1
                    rowSOBalance = TryCast(DataView.Row, Sales_Module.dseTransActivities.tbltrans_so_item2Row)
                    If Counter = 1 Then
                        DUmmyCustomer = rowSOBalance.Customer_Code
                    Else
                        If DUmmyCustomer <> rowSOBalance.Customer_Code Then
                            EmailContent = EmailContent + vbCrLf
                            DUmmyCustomer = rowSOBalance.Customer_Code
                        End If
                    End If

                    EmailContent = EmailContent + Counter.ToString + ".  Customer:" + rowSOBalance.Customer_Name + "(" + rowSOBalance.Customer_Code + _
                    ") - SO Number: " + rowSOBalance.SO_No + " - Product Code: " + rowSOBalance.Product_Code + vbCrLf
                Next

                'kirimkan email
                ObjEmail.Content = EmailContent
                ObjEmail.Subject = "Tomorrow's delivery"
                ObjEmail.TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                ObjEmail.CreatedBy = "System"
                ObjEmail.Owner = UserCode
                ObjEmail.Email_Status = E_mail_Module.ClsTransEmail.EnumEmailStatus.NewEmail
                ObjEmail.Folder = E_mail_Module.ClsTransEmail.EnumFolder.Inbox
                SendTo.Add(UserCode)
                ObjEmail.SendEmail(SendTo)

                'hapuskan semua baris yang user codenya sama
                Counter = 0
                While Counter < TBLSOBalance.Rows.Count
                    Ketemu = False
                    rowSOBalance = TBLSOBalance.Rows(Counter)
                    If rowSOBalance.Created_By = UserCode Then
                        TBLSOBalance.Rows.Remove(rowSOBalance)
                        Ketemu = True
                    End If

                    If Ketemu = False Then
                        Counter += 1
                    End If
                End While

                ijk += 1
            End While

            '2. kirimkan email untuk booking yang jatuh tempo besok
            DTA2.FillBy_ValidUntil(TBL2, Today)
            While TBL2.Rows.Count > 0
                rowBooking = TBL2.Rows(0)
                BookingNo = rowBooking.Booking_Number
                UserCode = rowBooking.Created_By
                Counter = 0
                SendTo.Clear()

                'filter berdasarkan user
                Filter = "[Created By] = '" + UserCode + "'"
                TBL2.DefaultView.RowFilter = Filter

                'buat email
                EmailContent = "The following item are booked and promised to be delivered tomorrow:" + vbCrLf
                For Each DataView In TBL2.DefaultView
                    Counter += 1
                    rowBooking = DataView.Row
                    If Counter = 1 Then
                        DUmmyCustomer = rowBooking.Customer_Code
                    Else
                        If DUmmyCustomer <> rowBooking.Customer_Code Then
                            EmailContent = EmailContent + vbCrLf
                            DUmmyCustomer = rowBooking.Customer_Code
                        End If
                    End If

                    EmailContent = EmailContent + Counter.ToString + ".  Customer:" + rowBooking.Customer_Name + "(" + rowBooking.Customer_Code + _
                    ") - Booking Number: " + rowBooking.Booking_Number + " - SO Number: " + rowBooking.SO_Number + " - Product Code: " + rowBooking.Product_Code + vbCrLf
                Next

                'kirimkan email
                ObjEmail.Content = EmailContent
                ObjEmail.Subject = "Tomorrows's delivery (inventory booking)"
                ObjEmail.TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                ObjEmail.CreatedBy = "System"
                ObjEmail.Owner = UserCode
                ObjEmail.Email_Status = E_mail_Module.ClsTransEmail.EnumEmailStatus.NewEmail
                ObjEmail.Folder = E_mail_Module.ClsTransEmail.EnumFolder.Inbox
                SendTo.Add(UserCode)
                ObjEmail.SendEmail(SendTo)

                'hapuskan semua baris yang user codenya sama
                Counter = 0
                While Counter < TBL2.Rows.Count
                    Ketemu = False
                    rowBooking = TBL2.Rows(Counter)
                    If rowBooking.Created_By = UserCode Then
                        TBL2.Rows.Remove(rowBooking)
                        Ketemu = True
                    End If

                    If Ketemu = False Then
                        Counter += 1
                    End If
                End While
            End While

            '3. hapuskan booking yang jatuh kemarin hari ini + kirimkan email ke yang punya bahwa booking dihapuskan
            TBL2.Rows.Clear()
            DTA2.FillBy_ValidUntil2(TBL2, Today - TimeSpan.FromDays(1))

            While TBL2.Rows.Count > 0
                rowBooking = TBL2.Rows(0)
                BookingNo = rowBooking.Booking_Number
                UserCode = rowBooking.Created_By
                Counter = 0
                SendTo.Clear()

                'filter berdasarkan user
                Filter = "[Created By] = '" + UserCode + "'"
                TBL2.DefaultView.RowFilter = Filter

                'buat email
                EmailContent = "The following booking item are expired yesterday and automatically deleted by the system:" + vbCrLf
                For Each DataView In TBL2.DefaultView
                    Counter += 1
                    rowBooking = DataView.Row
                    If Counter = 1 Then
                        DUmmyCustomer = rowBooking.Customer_Code
                    Else
                        If DUmmyCustomer <> rowBooking.Customer_Code Then
                            EmailContent = EmailContent + vbCrLf
                            DUmmyCustomer = rowBooking.Customer_Code
                        End If
                    End If

                    Email_Dummy = EmailContent + Counter.ToString + ".  Customer:" + rowBooking.Customer_Name + "(" + rowBooking.Customer_Code + _
                    ") - Booking Number: " + rowBooking.Booking_Number + " - SO Number: " + rowBooking.SO_Number + " - Product Code: " + rowBooking.Product_Code

                    If Email_Dummy.Length > 990 Then
                        'kirimkan email
                        ObjEmail.Content = EmailContent
                        ObjEmail.Subject = "Deleted inventory booking"
                        ObjEmail.TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                        ObjEmail.CreatedBy = "System"
                        ObjEmail.Owner = UserCode
                        ObjEmail.Email_Status = E_mail_Module.ClsTransEmail.EnumEmailStatus.NewEmail
                        ObjEmail.Folder = E_mail_Module.ClsTransEmail.EnumFolder.Inbox
                        SendTo.Add(UserCode)
                        ObjEmail.SendEmail(SendTo)

                        'masukkan data email selanjutnya
                        Counter = 1
                        EmailContent = Counter.ToString + ".  Customer:" + rowBooking.Customer_Name + "(" + rowBooking.Customer_Code + _
                        ") - Booking Number: " + rowBooking.Booking_Number + " - SO Number: " + rowBooking.SO_Number + " - Product Code: " + rowBooking.Product_Code + vbCrLf
                        SendTo.Clear()
                    Else
                        EmailContent = EmailContent + Counter.ToString + ".  Customer:" + rowBooking.Customer_Name + "(" + rowBooking.Customer_Code + _
                        ") - Booking Number: " + rowBooking.Booking_Number + " - SO Number: " + rowBooking.SO_Number + " - Product Code: " + rowBooking.Product_Code + vbCrLf
                    End If
                Next

                'kirimkan email
                ObjEmail.Content = EmailContent
                ObjEmail.Subject = "Deleted inventory booking"
                ObjEmail.TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                ObjEmail.CreatedBy = "System"
                ObjEmail.Owner = UserCode
                ObjEmail.Email_Status = E_mail_Module.ClsTransEmail.EnumEmailStatus.NewEmail
                ObjEmail.Folder = E_mail_Module.ClsTransEmail.EnumFolder.Inbox
                SendTo.Add(UserCode)
                ObjEmail.SendEmail(SendTo)

                'hapuskan semua baris yang user codenya sama
                Counter = 0
                While Counter < TBL2.Rows.Count
                    Ketemu = False
                    rowBooking = TBL2.Rows(Counter)
                    If rowBooking.Created_By = UserCode Then
                        TBL2.Rows.Remove(rowBooking)
                        Ketemu = True
                    End If

                    If Ketemu = False Then
                        Counter += 1
                    End If
                End While
            End While
            ObjBookingItem.HapusBookingByValidDate(Today - TimeSpan.FromDays(1))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            rnd = Nothing
            ObjEmail = Nothing
            DTA1.Dispose()
            SendTo = Nothing
        End Try
    End Sub

    Private Sub FrmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            If IsNothing(Me.NotifyIcon1) = False Then
                Me.NotifyIcon1.Visible = False
                Me.NotifyIcon1.Dispose()
            End If
            If IsNothing(Me.EW) = False Then
                Me.EW.Enabled = False
                Me.EW.Dispose()
                Me.EW = Nothing
            End If
            FrmLogin.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            If IsNothing(EW) = False Then
                'activates email watcher
                Me.EW.User = FrmLogin.UsernameTextBox.Text
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim obj As New ClsUser
        Dim ObjLS As New clsMasterListedSoftware
        Dim CompanyName As String = ""
        Dim PerbaruiLicense As Boolean = False
        Dim DateDiff As Integer = 0
        Dim ObjReminder As New AccModule.ClsTransaksiReminder
        Dim ObjDB As New ClsDBVersion
        Dim ObjMasterPriorityLevel As New Core_Module.ClsMasterPriorityLevel
        Dim ObjUpdateDB As New ClsUpdateDatebase
        Dim ObjRecurringReminder As New AccModule.ClsTransaksiRecurring
        'Dim ObjUpdate As New Download_Module.ClsAutoUpdate
        'Dim ObjVersion As New Download_Module.ClsVersionControl

        Try
            'MsgBox("Ini sudah update.", MsgBoxStyle.Information, "Information")

            'pastikan versi, untuk installer paling baru
            'If ObjVersion.GetVersion() = "" Then
            '    ObjVersion.SetVersion("1.6.00")
            'End If
            ''check for update via LAN
            'Me.CheckForUpdateLAN(False)

            'start monitor
            Me.Identifier = Me.ObjRandom.Next(0, 10 ^ 7)
            Me.ObjMonitor.StartMonitor(FrmLogin.UsernameTextBox.Text, My.Computer.Name, _
            Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)
            Me.Timer1.Interval = Me.ObjMonitor.MonitorInterval
            Me.Timer1.Enabled = True
            If Me.ObjMonitor.ApakahLicenseSMESSCukup() = False Then
                MsgBox("Insufficient SMESS active license, please release some license.", _
                MsgBoxStyle.Information, "Information")
                Dim i As New Core_Module.FrmTransaksiMonitor
                i.ShowDialog()
                Me.Close()
                Exit Sub
            End If

            Me.InventoryTransactionToolStripMenuItem.Text = "Purchase"
            Me.SalesTransactionToolStripMenuItem.Text = "Sales"
            Me.InventoryTransactionToolStripMenuItem1.Text = "Inventory"
            Me.HRD_Trans_ToolStripMenuItem4.Text = "HR"
            Me.AccountingTransactionToolStripMenuItem.Text = "Accounting"

            Me.Icon = My.Resources.desktop_icon_3_fw
            'cek versi database
            If My.Settings.db_version <> ObjDB.GetDBVersion Then
                If ObjUpdateDB.UpdateDatabase(ObjDB.GetDBVersion) = False Then
                    Me.CloseForm = True
                    Me.Close()
                    Exit Sub
                End If
            End If

            'cek hak akses yang belum ada di dalam database
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
            If ObjMasterPriorityLevel.GetJml = 0 Then
                ObjMasterPriorityLevel.simpanBaru(1, "R", "Reguler", "USER")
                ObjMasterPriorityLevel.setDefaultByPrimary("USER", ClsMasterPriorityLevel.enumDefault.yes, ObjMasterPriorityLevel.getPrimary("R"))
            End If

            'cek visibilitas
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'load custom report form
            Dim ObjCR As New CustomSMESS.ClsTrans_CustomReport(FrmLogin.UsernameTextBox.Text)
            'masukkan hak akses untuk menu custom report
            ObjCR.UpdateMenu()
            ObjCR.IsikanMenuCustom(Me.CustomReportToolStripMenuItem, Me, FrmLogin)
            If Me.CustomReportToolStripMenuItem.DropDownItems.Count = 0 Then
                Me.CustomReportToolStripMenuItem.Dispose()
            End If
            ObjCR = Nothing

            Dim ObjCM As New CustomSMESS.ClsTrans_CustomTrans(FrmLogin.UsernameTextBox.Text)
            'masukkan hak akses untuk menu custom master/trans
            ObjCM.UpdateMenu()
            'If ObjCM.ApakahAdaCustomTrans = True Then
            ObjCM.IsikanMenuCustom(Me.CustomMasterDataToolStripMenuItem, Me, FrmLogin, CustomSMESS.ClsTrans_CustomTrans.EnumTipeMenu.Master, FrmLogin.UsernameTextBox.Text)
            If Me.CustomMasterDataToolStripMenuItem.DropDownItems.Count = 0 Then
                Me.CustomMasterDataToolStripMenuItem.Dispose()
            End If
            ObjCM.IsikanMenuCustom(Me.CustomTransactionToolStripMenuItem1, Me, FrmLogin, CustomSMESS.ClsTrans_CustomTrans.EnumTipeMenu.Transaction, FrmLogin.UsernameTextBox.Text)
            If Me.CustomTransactionToolStripMenuItem1.DropDownItems.Count = 0 Then
                Me.CustomTransactionToolStripMenuItem1.Dispose()
            End If
            'End If
            ObjCM = Nothing
            If Me.SMESSCustom_1.DropDownItems.Count = 0 Then
                Me.SMESSCustom_1.Dispose()
            End If

            'cek berkaitan dengan aktivitas sales
            Dim nt As New Threading.Thread(AddressOf DoFirstActivitySales)
            nt.Start()
            'Me.DoFirstActivitySales()

            'cek menu strip
            Me.ObjCD.CekVisibilitasMenu(Me.MenuStrip1, FrmLogin.UsernameTextBox.Text)

            'jika tidak ada stock taking aktif, maka stock taking process tidak perlu ditampilkan
            Dim ObjST As New Inventory_Module.clsTransaksiStockTaking
            If ObjST.CekApakahAdaYangMasihAktif(False) = False Then
                FrmTransaksiStockTakingProcess1.Visible = False
                frmTransaksiStockTaking_Print1.Visible = False
            End If
            ObjST = Nothing
            'default menu yg muncul dan yg tidak muncul
            Me.AturMenuDefault()

            'activates email watcher, untuk sementara email watcher dimatikan dulu, karen mengganggu proses debugging
            EW = New E_mail_Module.ClsEmailWatcher(FrmLogin.UsernameTextBox.Text, Me, Me.NotifyIcon1, _
            Me.AlertMeOnIncomingEmailToolStripMenuItem)
            If My.Settings.Email_Wathcer = "1" Then
                EW.Enabled = True
                AlertMeOnIncomingEmailToolStripMenuItem.Checked = True
            Else
                EW.Enabled = False
                AlertMeOnIncomingEmailToolStripMenuItem.Checked = False
            End If

            Dim ijk As New FrmMainSMESS
            ijk.MdiParent = Me
            ijk.BackColorAll = Core_Module.GetSMESSBackColorAll
            ijk.Show()

            CompanyName = Me.ObjComapanyData.GetCompanyName
            If CompanyName <> "" Then
                CompanyName = " - " + CompanyName
            End If
            Me.Text = "SMESS - Ver. " + ObjDB.GetDBVersion + _
            CompanyName + " - Connection: " + Core_Module.GetDBServer + "(" + Core_Module.GetDBName + ")"

            'coding berkaitan dengan licensing
            If My.Settings.a = "" Then
                'masuk trial period
                My.Settings.a = Me.ObjLicense.GetProductKey
                My.Settings.b = Me.ObjLicense.GetLicenseKey(Today, My.Settings.a, ClsLicense.EnumTipeLicense.Trial)
                My.Settings.c = 400
                My.Settings.Save()

                MsgBox("This is a trial aplication. Please enter license key to get full aplication function", MsgBoxStyle.Information, "Information")
                Me.Text = Me.Text + " (Trial)"
            Else
                'kurangkan login
                My.Settings.c -= 1
                My.Settings.Save()
                PerbaruiLicense = False

                '1. cek product key pada license dan hardware apakah sama
                If My.Settings.a <> Me.ObjLicense.GetProductKey Then
                    MsgBox("Your product key is invalid. This can happen if you change your hardware. Please contact your software vendor to obtain new license key.", _
                    MsgBoxStyle.Information, "Information")
                    Me.ShowLicenseToEnter()
                Else
                    '2. cek expiry date
                    If Today > Me.ObjLicense.GetExpiryDate(My.Settings.b) Then
                        MsgBox("Your license has expired. Please contact your software vendor to obtain new license key.", _
                        MsgBoxStyle.Information, "Information")
                        Me.ShowLicenseToEnter()
                    End If
                End If

                If Me.ObjLicense.GetTipeLicense(My.Settings.b) = ClsLicense.EnumTipeLicense.Trial Then
                    MsgBox("This is a trial aplication. Please enter license key to get full aplication function", MsgBoxStyle.Information, "Information")
                    Me.Text = Me.Text + " (Trial)"
                Else
                    'jika sudah 1 bulan sebelum license berakhir maka minta untuk update
                    DateDiff = (Me.ObjLicense.GetExpiryDate(My.Settings.b) - Today).TotalDays
                    If DateDiff <= 0 Then
                        DateDiff *= -1
                    End If

                    If DateDiff <= 30 Then
                        MsgBox("The license will expire in " + Format(Me.ObjLicense.GetExpiryDate(My.Settings.b), "dd MMM yyyy") + " please contact your software to renew your license.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If

            Core_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Inventory_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Purchase_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Sales_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            AccModule.ModGlobalVariable.SetLicense(My.Settings.b)
            E_mail_Module.ModGlobalVariable.SetLicense(My.Settings.b)

            'reminder
            If ObjReminder.GetJml(Today, FrmLogin.UsernameTextBox.Text) > 0 Then
                MsgBox("You have reminder for today", MsgBoxStyle.Information, "Information")
                Me.FrmTransaksiReminder1.PerformClick()
            End If

            'Dim ijklm As New Purchase_Module.FrmCekMovement
            'ijklm.ShowDialog()
            'Me.ManajemenUserToolStripMenuItem.Visible = True
            'Me.FrmManajemenHakAkses1.Visible = True
            'Me.FrmLaporanPembelian1.Visible = True
            'Me.FrmLaporanPenjualan1.Visible = True
            'My.Settings.a = "58DCC - S9MEX -   098 - XPMVQ - UACVP"
            'My.Settings.b = "6cT0O40ru2Sc9pq1pf1n90so0"
            'Me.FrmTransaksiAOP_Quick1.Visible = True
            'Me.CustomTransactionToolStripMenuItem.Visible = True
            'Me.TicketingInvoiceToolStripMenuItem.Visible = True
            'Me.FrmTransaksiPOSTransfer1.Visible = True

            'remove menu2 yang belum siap untuk dipublish
            Me.MenuStrip1.Items.Remove(POS_Menu)
            Me.AccountingTransactionToolStripMenuItem.DropDownItems.Remove(Me.FixedAssetToolStripMenuItem)
            Me.AccReportToolStripMenuItem.DropDownItems.Remove(FrmLapAsset_DepresiasiPerPeriod1)
            'Me.AccountingTransactionToolStripMenuItem.DropDownItems.Remove(PaToolStripMenuItem)
            'Me.SalesTransactionToolStripMenuItem.DropDownItems.Remove(FrmTransaksiSalesReturn_Quick1)
            'Me.DataMasterToolStripMenuItem.DropDownItems.Remove(FrmMasterDepartement1)
            Me.DataMasterToolStripMenuItem.DropDownItems.Remove(FrmMasterJob1)
            'Me.ToolsToolStripMenuItem.DropDownItems.Remove(CheckForUpdateToolStripMenuItem)
            Me.VersionControlToolStripMenuItem.Visible = True
            Me.FrmUpdateFromLAN1.Visible = True

            'HR Transaksi
            'Me.DataMasterToolStripMenuItem.DropDownItems.Remove(Me.HumanResourceToolStripMenuItem)
            'Me.MenuStrip1.Items.Remove(Me.HRD_Trans_ToolStripMenuItem4)
            'Me.ReportsToolStripMenuItem3.DropDownItems.Remove(Me.HRD_Report_ToolStripMenuItem1)
            'Me.ToolsToolStripMenuItem.DropDownItems.Remove(FrmTrans_TaxBeginningBalance11)
            'Me.ToolsToolStripMenuItem.DropDownItems.Remove(LoanBeginningBalanceToolStripMenuItem)

            My.Settings.Save()

            'start auto back up
            Dim ObjAutoDB As New Core_Module.ClsMasterBackUpDB
            ObjAutoDB.StartBackUpScheduler(FrmLogin.UsernameTextBox.Text)

            'update setting untuk inventory movement
            Purchase_Module.UpdateSettingsPO()

            'jika user ybs memiliki reminder jurnal, maka keluarkan remindernya.
            If ObjRecurringReminder.GetJmlReminderAktif(Today, Me.ObjUser.GetPrimary(FrmLogin.UsernameTextBox.Text)) > 0 Then
                FrmTransaksiRecuring_Reminder1.PerformClick()
            End If

            'update nama user pada core module
            Core_Module.SetUser(FrmLogin.UsernameTextBox.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            obj = Nothing
            ObjLS = Nothing
            ObjReminder = Nothing
            ObjDB = Nothing
            ObjMasterPriorityLevel = Nothing
            ObjUpdateDB = Nothing
            ObjRecurringReminder = Nothing
            'ObjVersion = Nothing
        End Try
    End Sub

    Public Sub AturMenuDefault()
        Try
            If Me.LicenseToolStripMenuItem.Visible = False Then
                Me.LicenseToolStripMenuItem.Visible = True
                Me.WindowsToolStripMenuItem.Visible = True
                Me.TileHorizontallyToolStripMenuItem.Visible = True
                Me.TileVericallyToolStripMenuItem.Visible = True
                Me.CascadeToolStripMenuItem.Visible = True
                Me.CloseAllWindowsToolStripMenuItem.Visible = True
                Me.InternalMessageToolStripMenuItem.Visible = True
                Me.AlertMeOnIncomingEmailToolStripMenuItem.Visible = True
                Me.Test1.Visible = True 'email pribadi
                Me.FileToolStripMenuItem.Visible = True
                Me.KeluarToolStripMenuItem.Visible = True
                Me.FrmTransaksiMonitor.Visible = True
                Windows.Forms.Cursor.Current = Cursors.Arrow
                Me.MainPanelToolStripMenuItem.Visible = True
                Me.FrmTransactionQueue1.Visible = False
            End If

            If Me.ObjClosing.GetJml > 0 Then
                Me.AccountingStartUpSettingToolStripMenuItem.Visible = False
                Me.FrmTransaksiInventoryBeginingBalance1.Visible = False
            End If

            'Me.FrmSettingUpdate1.Visible = True
            'UploadToolStripMenuItem.Visible = True
            'FrmDownload1.Visible = True
            'VersionControlToolStripMenuItem.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowLicenseToEnter()
        Dim i As New FrmTransaksiLicensing

        Try
            'cek kembali, jika masih bermasalah maka tutup aplikasi
            If My.Settings.a <> Me.ObjLicense.GetProductKey Then
                If i.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                    Me.CloseForm = True
                    Me.Close()
                End If

                If Today > Me.ObjLicense.GetExpiryDate(My.Settings.b) Then
                    MsgBox("Your license has expired. Please contact your software vendor to obtain new license key.", _
                    MsgBoxStyle.Information, "Information")
                    Me.Close()
                End If

                '3. jika tipe license trial, maka hanya bisa login sejumlah tertentu
                If My.Settings.c < 0 Then
                    MsgBox("Your trial has expired. Please contact your software vendor to obtain new license key.", _
                    MsgBoxStyle.Information, "Information")
                    Me.Close()
                End If
            Else
                i.ShowDialog()

                '2. cek expiry date
                If Today > Me.ObjLicense.GetExpiryDate(My.Settings.b) Then
                    'MsgBox("Your license has expired. Please contact your software vendor to obtain new license key.", _
                    'MsgBoxStyle.Information, "Information")
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FrmManajemenHakAkses1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmManajemenHakAkses1.Click
        Try
            Dim ij As FrmManajemenHakAkses2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmManajemenHakAkses2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmManajemenHakAkses2
            i.MdiParent = Me
            i.ShowIcon = False
            i.MainMenuStrip1 = Me.MenuStrip1
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        Dim ObjDB As New Core_Module.ClsMasterBackUpDB

        Try
            If MsgBox("Exit SMESS?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If ObjDB.ApakahRemindOnClose(FrmLogin.UsernameTextBox.Text.Trim) = True Then
                    If MsgBox("Back up before close?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        FrmDia.Title = "Please specify where to save back up file"
                        FrmDia.Filter = "*.smess|smess back up"
                        FrmDia.FilterIndex = 0
                        FrmDia.DefaultExt = ".smess"
                        FrmDia.OverwritePrompt = True
                        If FrmDia.ShowDialog <> DialogResult.OK Then
                            Exit Sub
                        End If
                        If FrmDia.FileName = "" Then
                            Exit Sub
                        End If

                        Dim ObjBackUp As New Core_Module.ClsMySQL
                        ObjBackUp.BackUpDataBase(Core_Module.GetConstring, FrmDia.FileName, Core_Module.GetDBServer, Core_Module.GetDBName, _
                        Core_Module.GetDBUserID, Core_Module.GetDBPassword, Core_Module.GetDBPort)
                        ObjBackUp = Nothing
                    End If
                End If

                Me.CloseForm = True
                FrmLogin.Close()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjDB = Nothing
        End Try
    End Sub

    Private Sub FrmMasterUser1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterUser1.Click
        Try
            Dim ij As FrmMasterUser
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterUser)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterUser
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmGantiPassword1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmGantiPassword1.Click
        Try
            Dim ij As FrmGantiPassword
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmGantiPassword)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmGantiPassword
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.KodeUser = FrmLogin.UsernameTextBox.Text
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmResetPassword1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmResetPassword1.Click
        Try
            Dim ij As FrmResetPassword
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmResetPassword)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmResetPassword
            ' Increment the caption counter.
            ' i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            ' i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterMember1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProductGroup1.Click
        Try
            Dim ij As FrmMasterProductGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterProductGroup)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterProductGroup
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BrandToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterBrand1.Click
        Try
            Dim ij As FrmMasterBrand
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterBrand)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterBrand
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SubBrandToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSubBrand1.Click
        Try
            Dim ij As FrmMasterSubBrand
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterSubBrand)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterSubBrand
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCategory1.Click
        Try
            Dim ij As frmMasterCategory
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, frmMasterCategory)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New frmMasterCategory
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ProductToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProduct1.Click
        Try
            Dim ij As frmMasterProduct
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, frmMasterProduct)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New frmMasterProduct
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CustomerGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerGroup1.Click
        Try
            Dim ij As FrmMasterCustomerGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterCustomerGroup)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterCustomerGroup
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomer1.Click
        Try
            Dim ij As FrmMasterCustomer
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterCustomer)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterCustomer
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CustomerTypeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerType1.Click
        Try
            Dim ij As FrmMasterCustomerType
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterCustomerType)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterCustomerType
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MarketSegmentationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerSegmentation1.Click
        Try
            Dim ij As FrmMasterCustomerSegmentation
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterCustomerSegmentation)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterCustomerSegmentation
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesArea1.Click
        Try
            Dim ij As FrmMasterSalesArea
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterSalesArea)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterSalesArea
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSupplier1.Click
        Try
            Dim ij As FrmMasterSupplier
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterSupplier)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterSupplier
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UnitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterUnit1.Click
        Try
            Dim ij As FrmMasterUnit
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterUnit)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterUnit
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CurrenryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCurrency1.Click
        Try
            Dim ij As FrmMasterCurrency2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterCurrency2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterCurrency2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterGudang1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterGudang1.Click
        Try
            Dim ij As FrmMasterGudang
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterGudang)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterGudang
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterStorageArea1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterStorageArea1.Click
        Try
            Dim ij As FrmMasterStorageArea
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterStorageArea)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterStorageArea
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPO1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPO1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPO
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPO)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPO
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPackingList1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPackingList1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPackingList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPackingList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPackingList
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOPLokal1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAOPLokal1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiAOP2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiAOP2)
                If IsNothing(ij) = False Then
                    If ij.Lokal = True Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiAOP2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Lokal = True
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOPRImport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAOPRImport1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiAOP2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiAOP2)
                If IsNothing(ij) = False Then
                    If ij.Lokal = False Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiAOP2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Lokal = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPurchaseReturn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPurchaseReturn1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPurchaseReturn
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPurchaseReturn)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPurchaseReturn
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiHistoricalPurchase1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiHistoricalPurchase1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiHistoricalPurchase
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiHistoricalPurchase)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiHistoricalPurchase
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCalculateUnitCost1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiMonthlyClosing1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiMonthlyClosing2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiMonthlyClosing2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiMonthlyClosing2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseBySupplier1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseBySupplier1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseBySupplier
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseBySupplier)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseBySupplier
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseOrderBalanceBySupplier1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseOrderBalanceBySupplier1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseOrderBalance
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseOrderBalance)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseOrderBalance
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseRequisition1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseRequisition
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseRequisition)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseRequisition
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseStatisticBySupplier1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseStatisticBySupplier1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseStatisticBySupplier
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseStatisticBySupplier)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseStatisticBySupplier
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockBundling1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockBundling1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockBundling
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockBundling)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockBundling
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.ObjStockBundling.TipeBundling = Inventory_Module.ClsTransaksiStockBundling.EnumTipeBundling.Purchase
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
            i = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanInventoryMovementReports1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanInventoryMovementReports1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanInventoryMovementReports
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanInventoryMovementReports)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanInventoryMovementReports
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseOrder1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseOrder1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseOrder
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseOrder)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseOrder
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            ' i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CascadeToolStripMenuItem.Click
        Try
            Me.LayoutMdi(MdiLayout.Cascade)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TileVericallyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileVericallyToolStripMenuItem.Click
        Try
            Me.LayoutMdi(MdiLayout.TileVertical)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CloseAllWindowsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseAllWindowsToolStripMenuItem.Click
        Try
            Dim i As Windows.Forms.Form
            Dim ij As FrmMainSMESS

            For Each i In Me.MdiChildren
                ij = TryCast(i, FrmMainSMESS)
                If IsNothing(ij) = True Then
                    i.BringToFront()
                    i.Close()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TileHorizontallyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileHorizontallyToolStripMenuItem.Click
        Try
            Me.LayoutMdi(MdiLayout.TileHorizontal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AlertMeOnIncomingEmailToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlertMeOnIncomingEmailToolStripMenuItem.Click
        Try
            If Me.AlertMeOnIncomingEmailToolStripMenuItem.Checked = True Then
                Me.EW.Start()
                My.Settings.Email_Wathcer = "1"
            Else
                Me.EW.Stop()
                My.Settings.Email_Wathcer = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Test1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Test1.Click
        Try
            Dim j As E_mail_Module.FrmMainEmail
            Dim k As Windows.Forms.Form
            For Each k In Me.MdiChildren
                j = TryCast(k, E_mail_Module.FrmMainEmail)
                If IsNothing(j) = False Then
                    'activates form
                    j.BringToFront()
                    j.Focus()
                    'keluar
                    Exit Sub
                End If
            Next
            Dim i As New E_mail_Module.FrmMainEmail
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSettingAutomatedEmail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSettingAutomatedEmail1.Click
        Try
            Dim ij As E_mail_Module.FrmTransaksiSettingAutomatedEmail
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, E_mail_Module.FrmTransaksiSettingAutomatedEmail)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New E_mail_Module.FrmTransaksiSettingAutomatedEmail
            ' Increment the caption counter.
            ' i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AuthorizePOPurchaseStaffToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv11.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOOtorisasi)
                If IsNothing(ij) = False Then
                    If ij.Text = "Authorize Purchase Order - Purchase Staff" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOOtorisasi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 0
            i.CancelAuthorizationLevel = 1
            ' Set the caption to be unique.
            i.Text = "Authorize Purchase Order - Purchase Staff"
            'back color
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AuthorizePOPurchaseManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv21.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOOtorisasi)
                If IsNothing(ij) = False Then
                    If ij.Text = "Authorize Purchase Order - Purchase Manager" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOOtorisasi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 1
            i.CancelAuthorizationLevel = 2
            ' Set the caption to be unique.
            i.Text = "Authorize Purchase Order - Purchase Manager"
            'back color
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AuthorizePODirectorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv31.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOOtorisasi)
                If IsNothing(ij) = False Then
                    If ij.Text = "Authorize Purchase Order - Director" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOOtorisasi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 2
            i.CancelAuthorizationLevel = 3
            ' Set the caption to be unique.
            i.Text = "Authorize Purchase Order - Director"
            'back color
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub StockCardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockCard1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockCard2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockCard2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockCard2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub StockInAdjustToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockInAdj1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInAdj
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockInAdj)
                If IsNothing(ij) = False Then
                    If ij.BeginingBalance = False Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockInAdj
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.BeginingBalance = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub StockOutAdjustToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockOutAdj1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockOutAdj
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockOutAdj)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockOutAdj
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmLaporanAdjustmentList1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmLaporanAdjustmentList1.Click
        Try
            Dim ij As Inventory_Module.frmLaporanAdjustmentList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.frmLaporanAdjustmentList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.frmLaporanAdjustmentList
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            ' i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesAnalysis1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesAnalysis1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesAnalysis
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesAnalysis)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesAnalysis
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            ' i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesReturnList1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesReturnList1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesReturnList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesReturnList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesReturnList
            ' Increment the caption counter.
            ' Set the caption to be unique.
            ' i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseReturnList1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseReturnList1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseReturnList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseReturnList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseReturnList
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanInventoryMovement1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanInventoryMovement1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanInventoryMovement
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanInventoryMovement)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanInventoryMovement
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSO1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSO1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSO
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSO)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSO
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiDO1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiDO1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiDO3
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiDO3)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiDO3
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesReturn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesReturn1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesReturn
            Dim j As Windows.Forms.Form '

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesReturn)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesReturn
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            'Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoice1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesInvoice1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesInvoice
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoice)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesInvoice
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInSalesReturn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockInSalesReturn1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInSalesReturn2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockInSalesReturn2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockInSalesReturn2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockOutSales1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockOutSales1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockOutSales
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockOutSales)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockOutSales
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesListByDO1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesListByDO1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesListByDO
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesListByDO)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesListByDO
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseAuditList1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseAuditList1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseAuditList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseAuditList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseAuditList
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInPurchase1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockInPurchase1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInPurchase2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockInPurchase2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockInPurchase2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockOutPurchaseReturn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockOutPurchaseReturn1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockOutPurchaseReturn2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesListByInvoice1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesListByInvoice1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesListByInvoice
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesListByInvoice)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesListByInvoice
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            ' i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesStatistic1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesStatistic1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesStatistic
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesStatistic)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesStatistic
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotation1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiQuotation1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiQuotation
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiQuotation)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiQuotation
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKondisiStock1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKondisiStock1.Click
        Try
            Dim ij As Inventory_Module.FrmMasterKondisiStock
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmMasterKondisiStock)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmMasterKondisiStock
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmTransaksiStockTaking1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmTransaksiStockTaking1.Click
        Try
            Dim ij As Inventory_Module.frmTransaksiStockTaking
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.frmTransaksiStockTaking)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.frmTransaksiStockTaking
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmTransaksiStockTag1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmTransaksiStockTag1.Click
        Try
            Dim ij As Inventory_Module.frmTransaksiStockTag
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.frmTransaksiStockTag)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.frmTransaksiStockTag
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockTakingProcess1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockTakingProcess1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockTakingProcess
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockTakingProcess)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockTakingProcess
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockPindah1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockPindah1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockTransfer
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockTransfer)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockTransfer
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmLaporanStockByBatchNumberNExpiryDate1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmLaporanStockByBatchNumberNExpiryDate1.Click
        Try
            Dim ij As Inventory_Module.frmLaporanStockByBatchNumberNExpiryDate
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.frmLaporanStockByBatchNumberNExpiryDate)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.frmLaporanStockByBatchNumberNExpiryDate
            ' Increment the caption counter.
            ' i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanProductRanking1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProductRanking1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProductRanking
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProductRanking)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProductRanking
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanStockDetail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanStockDetail1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanStockDetail2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanStockDetail2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanStockDetail2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiBookingInv1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiBookingInv1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiBookingInv
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiBookingInv)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiBookingInv
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasiSalesStaff1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiQuotationOtorisasiSalesStaff1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiQuotationOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiQuotationOtorisasi)
                If IsNothing(ij) = False Then
                    If ij.AuthorizationLevel = 0 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiQuotationOtorisasi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 0
            i.CancelAuthorizationLevel = 1
            ' Set the caption to be unique.
            i.Text = i.Text & " - Sales Staff Authorization"

            i.BackColorAll = Core_Module.GetSMESSBackColorAll

            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasiSalesManager1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiQuotationOtorisasiSalesManager1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiQuotationOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiQuotationOtorisasi)
                If IsNothing(ij) = False Then
                    If ij.AuthorizationLevel = 1 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiQuotationOtorisasi
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 1
            i.CancelAuthorizationLevel = 2
            ' Set the caption to be unique.
            i.Text = i.Text & " - Sales Manager Authorization"
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasiGeneralManager1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiQuotationOtorisasiGeneralManager1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiQuotationOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiQuotationOtorisasi)
                If IsNothing(ij) = False Then
                    If ij.AuthorizationLevel = 2 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiQuotationOtorisasi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 2
            i.CancelAuthorizationLevel = 3
            ' Set the caption to be unique.
            i.Text = i.Text & " - General Manager Authorization"
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseStatisticBySupplier1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseStatisticBySupplier1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseStatisticBySupplier
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseStatisticBySupplier)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseStatisticBySupplier
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSOC1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSOC1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSOC
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSOC)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSOC
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanHistoricalSales1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanHistoricalSales1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanHistoricalSales
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanHistoricalSales)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanHistoricalSales
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesOrderBalance1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesOrderBalance1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesOrderBalance
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesOrderBalance)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesOrderBalance
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSettingPO1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPO2.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiSettingPO
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiSettingPO)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiSettingPO
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesBalanceManual1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesBalanceManual1.Click
        Try
            Dim ij As AccModule.FrmTransaksiSalesBalanceManual
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiSalesBalanceManual)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiSalesBalanceManual
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesReturnBalanceManual1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesReturnBalanceManual1.Click
        Try
            Dim ij As AccModule.FrmTransaksiSalesReturnBalanceManual
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiSalesReturnBalanceManual)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiSalesReturnBalanceManual
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPurchaseBalanceManual1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPurchaseBalanceManual1.Click
        Try
            Dim ij As AccModule.FrmTransaksiPurchaseBalanceManual
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiPurchaseBalanceManual)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiPurchaseBalanceManual
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPurchaseReturnBalanceManual1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPurchaseReturnBalanceManual1.Click
        Try
            Dim ij As AccModule.FrmTransaksiPurchaseReturnBalanceManual
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiPurchaseReturnBalanceManual)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiPurchaseReturnBalanceManual
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiGeneralJournal1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiGeneralJournal1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesCycle12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesCycle12.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Sales
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Sales)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_Sales
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReceivementJournal
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Purchase11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Purchase11.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Purchase
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Purchase)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_Purchase
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Closing1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Closing1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.ClosingJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.ClosingJournal
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiReminder1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiReminder1.Click
        Try
            Dim ij As AccModule.FrmTransaksiReminder
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiReminder)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiReminder
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiClosing1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiClosing1.Click
        Try
            Dim ij As AccModule.FrmTransaksiClosing
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiClosing)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiClosing
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiTagihan1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiTagihan1.Click
        Try
            Dim ij As AccModule.FrmTransaksiTagihan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiTagihan)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiTagihan
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanLedger1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanLedger1.Click
        Try
            Dim ij As AccModule.FrmLaporanLedger
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanLedger)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanLedger
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanTrialBalance1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanTrialBalance1.Click
        Try
            Dim ij As AccModule.FrmLaporanTrialBalance
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanTrialBalance)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanTrialBalance
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporan_ProfitNLoss1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporan_ProfitNLoss1.Click
        Try
            Dim ij As AccModule.FrmLaporan_ProfitNLoss
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporan_ProfitNLoss)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporan_ProfitNLoss
            i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanBalanceSheet1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanBalanceSheet1.Click
        Try
            Dim ij As AccModule.FrmLaporanBalanceSheet
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanBalanceSheet)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanBalanceSheet
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanARSummary1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanARSummary1.Click
        Try
            Dim ij As AccModule.FrmLaporanARSummary
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanARSummary)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanARSummary
            i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanAPSummary1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanAPSummary1.Click
        Try
            Dim ij As AccModule.FrmLaporanAPSummary
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanAPSummary)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanAPSummary
            i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakStockUnderMinimumQuantity1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmCetakStockUnderMinimumQuantity1.Click
        Try
            Dim ij As Inventory_Module.FrmCetakStockUnderMinimumQuantity
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmCetakStockUnderMinimumQuantity)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmCetakStockUnderMinimumQuantity
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChartOfAccountToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterAkun1.Click
        Try
            Dim ij As AccModule.FrmMasterAkun
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmMasterAkun)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmMasterAkun
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiBudget1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiBudget1.Click
        Try
            Dim ij As AccModule.FrmMasterBudget
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmMasterBudget)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmMasterBudget
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCompanyData1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCompanyData1.Click
        Try
            Dim ij As FrmMasterCompanyData
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterCompanyData)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmMasterCompanyData
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim ObjDB As New Core_Module.ClsMasterBackUpDB

        Try
            If Me.CloseForm = False Then
                If MsgBox("Exit SMESS?", MsgBoxStyle.YesNo, "Confimation") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                Else
                    If ObjDB.ApakahRemindOnClose(FrmLogin.UsernameTextBox.Text.Trim) = True Then
                        If MsgBox("Back up before close?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            FrmDia.Title = "Please specify where to save back up file"
                            FrmDia.Filter = "*.smess|smess back up"
                            FrmDia.FilterIndex = 0
                            FrmDia.DefaultExt = ".smess"
                            FrmDia.OverwritePrompt = True
                            If FrmDia.ShowDialog <> DialogResult.OK Then
                                Exit Sub
                            End If
                            If FrmDia.FileName = "" Then
                                Exit Sub
                            End If

                            Dim ObjBackUp As New Core_Module.ClsMySQL
                            ObjBackUp.BackUpDataBase(Core_Module.GetConstring, FrmDia.FileName, Core_Module.GetDBServer, Core_Module.GetDBName, _
                            Core_Module.GetDBUserID, Core_Module.GetDBPassword, Core_Module.GetDBPort, "", "", True)
                            ObjBackUp = Nothing
                        End If
                    End If

                    'release smess license
                    Me.ObjMonitor.EndMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)
                End If
            Else
                'release smess license
                Me.ObjMonitor.EndMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjDB = Nothing
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_View1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_View1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_View
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_View)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_View
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Try
            Dim gr As Graphics = e.Graphics

            gr.DrawString("Master Data", Me.Font, Brushes.White, 100, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanAgingAR1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanAgingAR1.Click
        Try
            Dim ij As AccModule.FrmLaporanAgingAR
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanAgingAR)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanAgingAR
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanKartuAR1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanKartuAR1.Click
        Try
            Dim ij As AccModule.FrmLaporanKartuAR
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanKartuAR)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanKartuAR
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanAgingAP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanAgingAP1.Click
        Try
            Dim ij As AccModule.FrmLaporanAgingAP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanAgingAP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanAgingAP
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanKartuAP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanKartuAP1.Click
        Try
            Dim ij As AccModule.FrmLaporanKartuAP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanKartuAP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanKartuAP
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LicenseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LicenseToolStripMenuItem.Click
        Try
            Dim ij As FrmTransaksiLicensing
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmTransaksiLicensing)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmTransaksiLicensing
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceQuick2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesInvoiceQuick2.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesInvoiceQuick
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoiceQuick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesInvoiceQuick
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub QuickPurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAOP_Quick1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiAOP_Quick2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiAOP_Quick2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiAOP_Quick2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPembelian1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPembelian1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPembelian
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPembelian)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPembelian
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalesListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPenjualan1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanPenjualan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanPenjualan)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanPenjualan
            ' Increment the caption counter.
            ' Set the caption to be unique.
            ' i.BackCOlorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSettingAP_AR1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmSettingAP_AR1.Click
        Try
            Dim ij As FrmSettingAP_AR
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmSettingAP_AR)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmSettingAP_AR
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPurchaseReturn_Quick1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPurchaseReturn_Quick1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPurchaseReturn_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPurchaseReturn_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPurchaseReturn_Quick
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BatchNoExpiryUpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiExpiryDateUpdate1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiExpiryDateUpdate
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiExpiryDateUpdate)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiExpiryDateUpdate
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterNumbering1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterNumbering1.Click
        Try
            Dim ij As FrmMasterNumbering
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterNumbering)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterNumbering
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_MdiChildActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MdiChildActivate
        Try
            Me.AturMenuDefault()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceSettingPajak1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesInvoiceSettingPajak1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesInvoiceSettingPajak
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoiceSettingPajak)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesInvoiceSettingPajak
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterDefaultValues1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterDefaultValues1.Click
        Try
            Dim ij As FrmMasterDefaultValues
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterDefaultValues)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmMasterDefaultValues
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_EndOfMonth1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_EndOfMonth1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_EndOfMonth
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_EndOfMonth)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal And _
                    ij.TipeJurnal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_EndOfMonth
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
            i.TipeJurnal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub POSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_POS1.Click
        'Try
        '    Dim ij As SMESS_POS.FrmTrans_POS
        '    Dim j As Windows.Forms.Form

        '    For Each j In Me.MdiChildren
        '        ij = TryCast(j, SMESS_POS.FrmTrans_POS)
        '        If IsNothing(ij) = False Then
        '            ij.BringToFront()
        '            ij.Focus()
        '            Exit Sub
        '        End If
        '    Next

        '    Dim i As New SMESS_POS.FrmTrans_POS
        '    ' Increment the caption counter.
        '    i.UserCode = FrmLogin.UsernameTextBox.Text
        '    ' Set the caption to be unique.
        '    i.BackColorAll = Core_Module.GetSMESSBackColorAll 
        '    i.MdiParent = Me i.ShowIcon = False
        '    i.Show()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub SessionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Session1.Click
        'Try
        '    Dim ij As SMESS_POS.FrmTrans_Session
        '    Dim j As Windows.Forms.Form

        '    For Each j In Me.MdiChildren
        '        ij = TryCast(j, SMESS_POS.FrmTrans_Session)
        '        If IsNothing(ij) = False Then
        '            ij.BringToFront()
        '            ij.Focus()
        '            Exit Sub
        '        End If
        '    Next

        '    Dim i As New SMESS_POS.FrmTrans_Session
        '    ' Increment the caption counter.
        '    i.UserCode = FrmLogin.UsernameTextBox.Text
        '    ' Set the caption to be unique.
        '    i.BackColorAll = Core_Module.GetSMESSBackColorAll 
        '    i.MdiParent = Me i.ShowIcon = False
        '    i.Show()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub FrmTransaksiInventoryBeginingBalance1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiInventoryBeginingBalance1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInAdj
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockInAdj)
                If IsNothing(ij) = False Then
                    If ij.BeginingBalance = True Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockInAdj
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False
            i.BeginingBalance = True

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DepartementToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterDepartement1.Click
        Try
            Dim ij As FrmMasterDepartement
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterDepartement)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterDepartement
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterJob1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterJob1.Click
        Try
            Dim ij As FrmMasterJob
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterJob)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterJob
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOPDP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAOPDP1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiAOPDP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiAOPDP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiAOPDP
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Lokal = True
            i.Text = "Purchase Down Payment"
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceDP1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesInvoiceDP1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesInvoiceDP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoiceDP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesInvoiceDP
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_CashBank1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_CashBank1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_CashBank
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_CashBank)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()


                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_CashBank
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
            i.TipeJurnall2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Cash_n_Bank
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanChequeAging1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanChequeAging1.Click
        Try
            Dim ij As AccModule.FrmLaporanChequeAging
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanChequeAging)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()

                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanChequeAging
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanChequePerPeriod1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanChequePerPeriod1.Click
        Try
            Dim ij As AccModule.FrmLaporanChequePerPeriod
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanChequePerPeriod)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()

                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanChequePerPeriod
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AssetDepreciationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_AssetDepreciation1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_AssetDepreciation
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_AssetDepreciation)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_AssetDepreciation
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.TipeJournal_Asset = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
            i.TipeJournal2_Asset = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.AssetDepreciation
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAssetGroup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterAssetGroup1.Click
        Try
            Dim ij As FrmMasterAssetGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterAssetGroup)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterAssetGroup
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAsset1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterAsset1.Click
        Try
            Dim ij As FrmMasterAsset
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterAsset)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterAsset
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAssetAcquisition_MaterialUsageA1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAssetAcquisition_MaterialUsageA1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiAssetAcquisition_MaterialUsage
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiAssetAcquisition_MaterialUsage)
                If IsNothing(ij) = False Then
                    If ij.InOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiAssetAcquisition_MaterialUsage
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.InOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition

            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAssetAcquisition_MaterialUsageB1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAssetAcquisition_MaterialUsageB1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiAssetAcquisition_MaterialUsage
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiAssetAcquisition_MaterialUsage)
                If IsNothing(ij) = False Then
                    If ij.InOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiAssetAcquisition_MaterialUsage
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.InOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation

            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Acquisition_Revaluation_OtherA1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Acquisition_Revaluation_OtherA1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Acquisition_Revaluation_Other
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Acquisition_Revaluation_Other)
                If IsNothing(ij) = False Then
                    If ij.TipeJurnall2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.AssetAcquistion_Other Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New AccModule.FrmTransaksiJournal_Acquisition_Revaluation_Other
            ' Increment the caption counter.
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.TipeJurnall2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.AssetAcquistion_Other

            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Acquisition_Revaluation_OtherB1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Acquisition_Revaluation_OtherB1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Acquisition_Revaluation_Other
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Acquisition_Revaluation_Other)
                If IsNothing(ij) = False Then
                    If ij.TipeJurnall2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.AssetRevaluation_Other Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New AccModule.FrmTransaksiJournal_Acquisition_Revaluation_Other
            ' Increment the caption counter.
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.TipeJurnall2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.AssetRevaluation_Other

            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_AssetDisposal1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_AssetDisposal1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_AssetDisposal
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_AssetDisposal)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.AssetDisposal Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New AccModule.FrmTransaksiJournal_AssetDisposal
            ' Increment the caption counter.
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.TipeJournal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.AssetDisposal

            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanVAT1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanVAT1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanVAT
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanVAT)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Purchase_Module.FrmLaporanVAT
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProses1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProses1.Click
        Try
            Dim ij As Core_Module.FrmMasterProses
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterProses)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterProses
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AssetDepreciationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLapAsset_DepresiasiPerPeriod1.Click
        Try
            Dim ij As AccModule.FrmLapAsset_DepresiasiPerPeriod
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLapAsset_DepresiasiPerPeriod)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmLapAsset_DepresiasiPerPeriod
            ' Increment the caption counter.
            'i.UserCOde = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll

            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterBackUpDB1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterBackUpDB1.Click
        Try
            Dim ij As FrmMasterBackUpDB
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterBackUpDB)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterBackUpDB
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterBackUpDB_Manual1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterBackUpDB_Manual1.Click
        Try
            Dim ObjBackUp As New Core_Module.ClsMySQL

            FrmDia.Title = "Please specify where to save back up file"
            FrmDia.Filter = "*.smess|smess back up"
            FrmDia.FilterIndex = 0
            FrmDia.DefaultExt = ".smess"
            FrmDia.OverwritePrompt = False
            If FrmDia.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If FrmDia.FileName = "" Then
                Exit Sub
            End If

            ObjBackUp.BackUpDataBase(Core_Module.GetConstring, FrmDia.FileName, Core_Module.GetDBServer, Core_Module.GetDBName, _
            Core_Module.GetDBUserID, Core_Module.GetDBPassword, Core_Module.GetDBPort, "", _
            "SMESS Manual Back Up", True)

            FrmDia.FileName = ""
            ObjBackUp = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmTransaksiStockTaking_Print1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmTransaksiStockTaking_Print1.Click
        Try
            Dim ij As Inventory_Module.frmTransaksiStockTaking_Print
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.frmTransaksiStockTaking_Print)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.frmTransaksiStockTaking_Print
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub StockTakingReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanStockTaking_Cost1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanStockTaking_Cost1
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanStockTaking_Cost1)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Inventory_Module.FrmLaporanStockTaking_Cost1
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanProductRanking_Modif1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProductRanking_Modif1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProductRanking_Modif
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProductRanking_Modif)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Inventory_Module.FrmLaporanProductRanking_Modif
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesReturn_Quick1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesReturn_Quick1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesReturn_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesReturn_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesReturn_Quick
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmRestoreDatabase1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmRestoreDatabase1.Click
        Try
            'pastikan hanya ada 1 online user pada saat akan restore
            If Me.ObjMonitor.GetJmlMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License) > 1 Then
                MsgBox("Please make sure only 1 user is online to start database restore.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            'tutup form monitor
            Dim ij As Core_Module.FrmTransaksiMonitor
            Dim j As Windows.Forms.Form
            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmTransaksiMonitor)
                If IsNothing(ij) = False Then
                    ij.Close()
                End If
            Next

            Me.OpenFileDialog1.Title = "Please specify database file to restore"
            Me.OpenFileDialog1.Filter = "smess files (*.smess)|*.smess"
            Me.OpenFileDialog1.FilterIndex = 1
            Me.OpenFileDialog1.DefaultExt = "smess"
            Me.OpenFileDialog1.CheckFileExists = True
            Me.OpenFileDialog1.Multiselect = False
            Me.OpenFileDialog1.FileName = ""

            If Me.OpenFileDialog1.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            If Me.OpenFileDialog1.FileName = "" Then
                Exit Sub
            End If

            'cancel monitoring
            Me.Timer1.Enabled = False
            Me.ObjMonitor.EndMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)

            Dim FrmRestore As New Core_Module.FrmBackUpData2
            FrmRestore.vFileName = Me.OpenFileDialog1.FileName
            FrmRestore.vConStr = Core_Module.GetConstring
            FrmRestore.vCompress = True
            FrmRestore.ShowDialog()

            FrmDia.FileName = ""

            'start monitoring kembali
            Me.Identifier = Me.ObjRandom.Next(0, 10 ^ 7)
            Me.ObjMonitor.StartMonitor(FrmLogin.UsernameTextBox.Text, My.Computer.Name, _
            Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)
            Me.Timer1.Interval = Me.ObjMonitor.MonitorInterval
            Me.Timer1.Enabled = True
            If Me.ObjMonitor.ApakahLicenseSMESSCukup() = False Then
                MsgBox("Insufficient SMESS active license, please release some license.", _
                MsgBoxStyle.Information, "Information")
                Dim i As New Core_Module.FrmTransaksiMonitor
                i.ShowDialog()
                Me.Close()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_PriceList1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmMasterProduct_PriceList1.Click
        Try
            Dim ij As frmMasterProduct_PriceList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, frmMasterProduct_PriceList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New frmMasterProduct_PriceList
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckForUpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckForUpdateToolStripMenuItem.Click
        Dim i As New FrmAutoUpdate

        Try
            'yang aktif hanya Frm Main saja
            If Me.MdiChildren.Length = 1 Then
                If MsgBox("Check for update?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.CloseForm = True
                    i.ShowDialog()
                End If
            Else
                MsgBox("Please close all form before update.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_MaterialPlanningDashboard1.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_MaterialPlanningDashboard
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_MaterialPlanningDashboard)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_MaterialPlanningDashboard
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SalesDashboard1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_SalesDashboard1.Click
        Try
            Dim ij As Sales_Module.FrmTrans_SalesDashboard
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTrans_SalesDashboard)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTrans_SalesDashboard
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterUser_TempPass1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterUser_TempPass1.Click
        Try
            Dim ij As FrmMasterUser_TempPass
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterUser_TempPass)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterUser_TempPass
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateHakAkses()
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            Me.ObjHakAkses.SetConnectionString(DTA)

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPOC_QuickDel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPOC_QuickDel", "Hak akses untuk membuat transaksi po cancel sekaligus.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_Override_Otorisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_Override_Otorisasi", _
                "Hak akses untuk edit/hapus transaksi purchase request walaupun sudah diotorisasi.", _
                "FrmTransaksiPR", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanStockBundling", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanStockBundling", "Hak akses untuk melihat laporan stock bundling.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmUpdateFromLAN", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmUpdateFromLAN", "Hak akses untuk melakukan update via LAN.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmSettingUpdate", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmSettingUpdate", "Hak akses untuk mensetting lokasi folder yang digunakan untuk melakukan update via LAN.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmDownload", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmDownload", "Hak akses untuk mendownload update terakhir dari cloud.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmProductListReport", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmProductListReport", "Hak akses untuk laporan daftar product.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterMachineData", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterMachineData", "Hak akses untuk input, edit, dan hapus master mesin.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_PR_Balance", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_PR_Balance", "Hak akses untuk membuat laporan purchase request balance.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_Purchase_Request_PickUp_Order", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_Purchase_Request_PickUp_Order", "Hak akses untuk membuat laporan purchase request pick up order.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_Purchase_Request_Cancel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_Purchase_Request_Cancel", "Hak akses untuk membuat laporan purchase request cancellation.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_Purchase_Request", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_Purchase_Request", "Hak akses untuk membuat laporan purchase request.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterRefParty", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterRefParty", "Hak akses untuk input, edit, dan hapus master pihak yang melakukan purchase request.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_PickUpOrder", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_PickUpOrder", "Hak akses untuk input, edit, dan hapus transaksi purchase request pick up order.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_Cancel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_Cancel", "Hak akses untuk input, edit, dan hapus transaksi purchase request cancellation.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR", "Hak akses untuk input, edit, dan hapus transaksi purchase request.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterMachineGroup", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterMachineGroup", "Hak akses untuk input, edit, dan hapus master grup mesin.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterWorkUnit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterWorkUnit", "Hak akses untuk input, edit, dan hapus master work unit.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDownTime", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterDownTime", "Hak akses untuk input, edit, dan hapus master down time.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterReject", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterReject", "Hak akses untuk input, edit, dan hapus master reject.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterProductAccountGroup", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterProductAccountGroup", "Hak akses untuk input, edit, dan hapus master kelompok akun product.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterKompetensi_Product", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterKompetensi_Product", "Hak akses untuk input, edit, dan hapus master custom value untuk product.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiRecuring", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiRecuring", "Hak akses untuk input, edit, dan hapus transaksi recurring.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiRecuring_Reminder", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiRecuring_Reminder", "Hak akses untuk form reminder recurring.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDebetNote", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDebetNote", "Hak akses untuk input, edit, dan hapus transaksi debit note.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCreditNote1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCreditNote1", "Hak akses untuk input, edit, dan hapus transaksi credit note.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanPurchaseByInvoice", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanPurchaseByInvoice", "Hak akses untuk melihat laporan pembelian berdasarkan invoice supplier.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmSalesOrderReportByDateDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmSalesOrderReportByDateDetail", "Hak akses untuk melihat laporan detail sales order.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmPurchaseOrderReportByDateDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmPurchaseOrderReportByDateDetail", "Hak akses untuk melihat laporan detail purchase order.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanAkunProduct", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanAkunProduct", "Hak akses untuk melihat laporan akun produk.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("frmMasterProduct_Verifikasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("frmMasterProduct_Verifikasi", "Hak akses untuk memverifikasi akun produk.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Attendance_Actual_In_Report", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Attendance_Actual_In_Report", "Hak akses untuk melihat data clock in karyawan.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Attendance_Log", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Attendance_Log", "Hak akses untuk melakukan import data dari server kehadiran.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterAttendanceServer", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterAttendanceServer", "Hak akses untuk input, edit, dan hapus data server kehadiran.", "", 0)
            Else
                Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanAPPaid", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanAPPaid", "Hak akses untuk melihat laporan balance hutang.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanARPaid", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanARPaid", "Hak akses untuk melihat laporan balance piutang.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDebetNote", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDebetNote", "Hak akses untuk input, edit, dan delete transaksi debet note.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCreditNote", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCreditNote", "Hak akses untuk input, edit, dan delete transaksi credit note.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransactionQueue", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransactionQueue", "Hak akses untuk mereset data dari tabel antrian.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterUser_Otorisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterUser_Otorisasi", "Hak akses untuk laporan approval dokumen yang di otorisasi.", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanVAT2", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanVAT2", "Hak akses untuk laporan vat penambahan fitur e-Faktur 2015.", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterDepartement", "Hak akses untuk data master departement.", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Gaji_Override", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Gaji_Override", "Hak akses untuk transaksi salary ovveride amount.", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_New", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_New", "Hak akses untuk input transaksi baru material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Edit", "Hak akses untuk edit transaksi material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Delete", "Hak akses untuk hapus transaksi material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Print", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Print", "Hak akses untuk cetak transaksi material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Apply", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Apply", "Hak akses untuk mengaplikasikan data transaksi material planning group ke dalam parameter perhitungan.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_ShowDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_ShowDetail", "Hak akses untuk menampilkan detail perhitungan data transaksi material planning group ke dalam parameter perhitungan.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_OverrideEdit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_OverrideEdit", "Hak akses untuk edit data transaksi material planning group yang dibuat user lain.", "FrmTrans_MaterialPlanningDashboard_Data1", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_OverrideDelete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_OverrideDelete", "Hak akses untuk delete data transaksi material planning group yang dibuat user lain.", "FrmTrans_MaterialPlanningDashboard_Data1", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1", "Hak akses untuk membuka form pengelompokan perencanaan material")
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmEmailSettings") = False Then
                Me.ObjHakAkses.Simpan("FrmEmailSettings", "Hak akses untuk input, edit, dan hapus transaksi setting email.")
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanSalesPerformance") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanSalesPerformance", "Hak akses untuk melihat laporan performa tim sales")
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesBudget") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesBudget", "Hak akses untuk melihat transaksi budget penjualan")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSalesPersonGroup") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSalesPersonGroup", "Hak akses untuk melihat master sales person group")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmCetakAR", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmCetakAR", "Hak akses untuk mengakses laporan pembayaran piutang")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmCetakAP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmCetakAP", "Hak akses untuk mengakses laporan pembayaran hutang")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SalesDashboard", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SalesDashboard", "Hak akses untuk mengakses sales dashboard")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterPriorityLevel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterPriorityLevel", "Hak akses untuk membuka input, edit, dan delete level prioritas pengiriman")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard", "Hak akses untuk membuka form perencanaan material")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmRestoreDatabase", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmRestoreDatabase", "Hak akses untuk melakukan restore database dari file sql back up")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesReturn_Quick", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesReturn_Quick", "Hak akses untuk input, edit, dan delete transaksi sales return")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("frmTransaksiStockTaking_Print", DTA) = False Then
                Me.ObjHakAkses.Simpan("frmTransaksiStockTaking_Print", "Hak akses untuk cetak form kosong stock taking")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterNumbering", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterNumbering", "Hak akses untuk mengatur penomoran transaksi")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoiceSettingPajak", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoiceSettingPajak", "Hak akses untuk melakukan setting faktuk pajak")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanProductRanking_Modif", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanProductRanking_Modif", "Hak akses untuk melihat laporan non-moving stock")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanStockTaking_Cost", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanStockTaking_Cost", "Hak akses untuk melihat laporan stock taking")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDefaultValues", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterDefaultValues", "Hak akses untuk melakukan setting default nilai tukar mata uang asing")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_EndOfMonth", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_EndOfMonth", "Hak akses untuk input, edit, dan hapus proses jurnal akhir bulan")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Session", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Session", "Hak akses untuk input, edit, dan hapus session point of sale")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_POS", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_POS", "Hak akses untuk input, edit, dan hapus point of sale")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterBackUpDB", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterBackUpDB", "Hak akses untuk input, edit, dan hapus rencana back up data")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterBackUpDB_Manual", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterBackUpDB_Manual", "Hak akses untuk back up data secara manual")
            End If
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDepartement") = False Then
            'Me.ObjHakAkses.Simpan("FrmMasterDepartement", "Hak akses untuk input, edit, dan hapus departement")
            'End If
            'If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJob") = False Then
            'Me.ObjHakAkses.Simpan("FrmMasterJob", "Hak akses untuk input, edit, dan hapus job")
            'End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiInventoryBeginingBalance", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiInventoryBeginingBalance", "Hak akses untuk input, edit, dan hapus begining balance")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoiceDP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoiceDP", "Hak akses untuk input, edit, dan hapus down payment sales invoice ")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiAOPDP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiAOPDP", "Hak akses untuk input, edit, dan hapus down payment purchase")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_CashBank", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_CashBank", "Hak akses untuk input, edit, dan hapus jurnal kas dan bank")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanChequeAging", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanChequeAging", "Hak akses untuk membuat laporan aging giro/cek")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanChequePerPeriod", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanChequePerPeriod", "Hak akses untuk membuat laporan giro/cek per period")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterAssetGroup", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterAssetGroup", "Hak akses untuk input, edit, dan hapus data master group asset")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterAsset", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterAsset", "Hak akses untuk input, edit, dan hapus data master asset")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiAssetAcquisition_MaterialUsageA", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiAssetAcquisition_MaterialUsageA", "Hak akses untuk input, edit, dan hapus data penggunaan material pada saat pengadaan aset.")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_Acquisition_Revaluation_OtherA", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_Acquisition_Revaluation_OtherA", "Hak akses untuk input, edit, dan hapus data ongkos pada saat pengadaan aset.")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiAssetAcquisition_MaterialUsageB", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiAssetAcquisition_MaterialUsageB", "Hak akses untuk input, edit, dan hapus data penggunaan material pada saat revaluasi.")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_Acquisition_Revaluation_OtherB", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_Acquisition_Revaluation_OtherB", "Hak akses untuk input, edit, dan hapus data ongkos pada saat revaluasi.")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_AssetDepreciation", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_AssetDepreciation", "Hak akses untuk input, edit, dan hapus data depresiasi aset.")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_AssetDepreciation", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_AssetDepreciation", "Hak akses untuk input, edit, dan hapus data depresiasi aset.")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_AssetDisposal", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_AssetDisposal", "Hak akses untuk input, edit, dan hapus data disposal aset.")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanVAT", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanVAT", "Hak akses untuk melihat laporan PPN")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSalesPerson", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSalesPerson", "Hak akses untuk input, edit, dan delete master sales person")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterProses", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterProses", "Hak akses untuk input, edit, dan delete master sales proses")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLapAsset_DepresiasiPerPeriod", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLapAsset_DepresiasiPerPeriod", "Hak akses untuk melihat laporan depresiasi aset")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPurchaseReturn_Quick", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPurchaseReturn_Quick", "Hak akses untuk input, edit, dan delete transaksi quick purchase return")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("frmMasterProduct_PriceList", DTA) = False Then
                Me.ObjHakAkses.Simpan("frmMasterProduct_PriceList", "Hak akses untuk melakukan edit price list")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmManajemenHakAkses_Override_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmManajemenHakAkses_Override_Edit", "Hak akses untuk mengedit data transaksi hak akses yang dibuat oleh user lain", "FrmManajemenHakAkses", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmManajemenHakAkses_Override_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmManajemenHakAkses_Override_Delete", "Hak akses untuk menghapus data transaksi hak akses yang dibuat oleh user lain", "FrmManajemenHakAkses", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Baru", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Baru", "Hak akses untuk membuat data transaksi so", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Edit", "Hak akses untuk mengedit data transaksi so", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Override_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Override_Edit", "Hak akses untuk mengedit data transaksi so yang dibuat oleh user lain", "FrmTransaksiSO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Otorisasi_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Otorisasi_Edit", "Hak akses untuk mengotorisasi proses edit data transaksi so", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Delete", "Hak akses untuk menghapus data transaksi so.", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Override_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Override_Delete", "Hak akses untuk menghapus data transaksi so yang dibuat oleh user lain", "FrmTransaksiSO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Otorisasi_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Otorisasi_Delete", "Hak akses untuk mengotorisasi proses hapus data transaksi so.", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Print", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Print", "Hak akses untuk mencetak data transaksi so.", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Override_Date", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Override_Date", "Hak akses untuk mengedit tanggal so", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_Override_Promise_To_Deliver", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_Override_Promise_To_Deliver", "Hak akses untuk mengedit janji kirim pada item so", "FrmTransaksiSO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SalesDashboard_Override_User", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SalesDashboard_Override_User", "Hak akses untuk melihat transaksi yang dibuat user lain pada sales dashboard", "FrmTrans_SalesDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP", "Hak akses untuk input, edit, dan hapus transaksi material planning", "FrmTrans_MaterialPlanningDashboard", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_EditParameter", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_EditParameter", "Hak akses untuk input parameter perhitungan transaksi material planning", "FrmTrans_MaterialPlanningDashboard", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_OverrideEdit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_OverrideEdit", "Hak akses untuk mengedit transaksi material planning yang dibuat oleh user lain", "FrmTrans_MaterialPlanningDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_OverrideDelete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_OverrideDelete", "Hak akses untuk menghapus transaksi material planning yang dibuat oleh user lain", "FrmTrans_MaterialPlanningDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF", "Hak akses untuk input, edit, dan hapus transaksi shipping forecast", "FrmTrans_MaterialPlanningDashboard", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_OverrideEdit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_OverrideEdit", "Hak akses untuk mengedit transaksi shipping forecast yang dibuat oleh user lain", "FrmTrans_MaterialPlanningDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_OverrideDelete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_OverrideDelete", "Hak akses untuk menghapus transaksi shipping forecast yang dibuat oleh user lain", "FrmTrans_MaterialPlanningDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_CreateDO", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_CreateDO", "Hak akses untuk membuat do dari transaksi shipping forecast", "FrmTrans_MaterialPlanningDashboard", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_EditQtyDO", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_EditQtyDO", "Hak akses untuk mengedit data qty do", "FrmTrans_MaterialPlanningDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_DOAll", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_DOAll", "Hak akses untuk membuat do atau edit tanggal janji pengiriman atas so user lain", "FrmTrans_MaterialPlanningDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_OverrideDO", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_OverrideDO", "Hak akses untuk membuat do atas customer yang bermasalah pembayarannya atau melebih credit limit", "FrmTrans_MaterialPlanningDashboard", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_Override_PurchaseETD", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_Override_PurchaseETD", "Hak akses untuk mengedit tanggal janji pengiriman supplier pada transaksi shipping forecast", "FrmTrans_MaterialPlanningDashboard", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_Override_SalesETD", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_Override_SalesETD", "Hak akses untuk mengedit tanggal janji pengiriman ke customer pada transaksi shipping forecast", "FrmTrans_MaterialPlanningDashboard", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPO_Override_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPO_Override_Edit", "Hak akses untuk mengedit transaksi po yang dibuat oleh user lain", "FrmTransaksiPO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPO_Override_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPO_Override_Delete", "Hak akses untuk menghapus transaksi po yang dibuat oleh user lain", "FrmTransaksiPO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPO_Override_Promise_To_Deliver", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPO_Override_Promise_To_Deliver", "Hak akses untuk mengedit tanggal janji kirim pada transaksi po", "FrmTransaksiPO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPO_Override_Otorisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPO_Override_Otorisasi", "Hak akses untuk mengedit atau menghapus transaksi po yang sudah di otorisasi", "FrmTransaksiPO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_New", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_New", "Hak akses untuk membuat transaksi DO", "FrmTransaksiDO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Edit", "Hak akses untuk mengedit transaksi DO", "FrmTransaksiDO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Hapus", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Hapus", "Hak akses untuk menghapus transaksi DO", "FrmTransaksiDO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Cetak", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Cetak", "Hak akses untuk mencetak transaksi DO", "FrmTransaksiDO", 1)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Override_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Override_Edit", "Hak akses untuk mengedit transaksi DO yang dibuat oleh user lain", "FrmTransaksiDO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Override_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Override_Delete", "Hak akses untuk menghapus transaksi DO yang dibuat oleh user lain", "FrmTransaksiDO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Override_ARDue", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Override_ARDue", "Hak akses untuk menyimpan transaksi do atas customer yang memiliki piutang jatuh tempo", _
                "FrmTransaksiDO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Override_ARCreditLimit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Override_ARCreditLimit", _
                "Hak akses untuk menyimpan transaksi do atas customer yang memiliki piutang melebihi limit piutang", "FrmTransaksiDO", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiBookingInv_Override_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiBookingInv_Override_Edit", "Hak akses untuk mengedit transaksi booking inventory yang dibuat oleh user lain", "FrmTransaksiBookingInv", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiBookingInv_Override_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiBookingInv_Override_Delete", "Hak akses untuk menghapus transaksi booking inventory yang dibuat oleh user lain", "FrmTransaksiBookingInv", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterUser_TempPass", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterUser_TempPass", "Hak akses untuk input, edit, dan delete transaksi password sementara untuk user", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterUser_TempPass_Override_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterUser_TempPass_Override_Delete", "Hak akses untuk menghapus transaksi password sementara yang dibuat user lain", "FrmMasterUser_TempPass", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterUser_TempPass_OverrideEdit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterUser_TempPass_OverrideEdit", "Hak akses untuk mengedit transaksi password sementara yang dibuat user lain", "FrmMasterUser_TempPass", 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub UpdateHakAkses_HR()
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterWorkShift1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterWorkShift1", "Hak akses untuk mengakses form master workshift", "", 0)
            Else
                Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterKompetensi1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterKompetensi1", "Hak akses untuk mengakses form master kompetensi", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterKaryawan1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterKaryawan1", "Hak akses untuk mengakses form master karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterCuti1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterCuti1", "Hak akses untuk mengakses form master cuti", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCuti_Pengajuan_Setting1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCuti_Pengajuan_Setting1", "Hak akses untuk mengakses form setting akses pengajuan cuti", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterJobPosition1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterJobPosition1", "Hak akses untuk mengakses form master jabatan karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSalaryFormula1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSalaryFormula1", "Hak akses untuk mengakses form master formula gaji", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSalaryFormula_OtherDeduction1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSalaryFormula_OtherDeduction1", "Hak akses untuk mengakses form master formula pengurangan gaji lain-lain", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterPayrollAccess1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterPayrollAccess1", "Hak akses untuk mengakses form pengaturan akses perhitugan gaji", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSettingPajak1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSettingPajak1", "Hak akses untuk mengatur nilai persentase besaran pajak sesuai peraturan yang berlaku", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSlipDescription1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSlipDescription1", "Hak akses untuk mengakses form untuk mengatur komponen cetakan pada slip gaji", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterWorkShift1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterWorkShift1", "Hak akses untuk mengakses form master shift kerja", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterWorkShift_Allocate1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterWorkShift_Allocate1", "Hak akses untuk mengakses form alokasi shift kerja", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCuti_Pengajuan1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCuti_Pengajuan1", "Hak akses untuk mengakses form pengajuan cuti", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCuti_Approval1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCuti_Approval1", "Hak akses untuk mengakses form persetujuan/penolakan atas pengajuan cuti", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Attendance1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Attendance1", "Hak akses untuk mengakses form absensi kehadiran", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Attendance_DataImport1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Attendance_DataImport1", "Hak akses untuk mengakses form input absensi kehadiran sekaligus atau impor data kehadiran dari mesin absensi sidik jari", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiLoan1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiLoan1", "Hak akses untuk mengakses form pinjaman karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiLoan_Payment1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiLoan_Payment1", "Hak akses untuk mengakses form pembayaran pinjaman karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Gaji1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Gaji1", "Hak akses untuk mengakses form perhitungan gaji", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Gaji_Quick1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Gaji_Quick1", "Hak akses untuk mengakses form perhitungan gaji untuk beberapa karyawan sekaligus", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Payment1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Payment1", "Hak akses untuk mengakses form pembayaran gaji", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Export_eSPT1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Export_eSPT1", "Hak akses untuk mengakses form export perhitungan pajak penghasilkan ke e-spt", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanEmployeeData1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanEmployeeData1", "Hak akses untuk mengakses form laporan data karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanAtt1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanAtt1", "Hak akses untuk mengakses form laporan absensi karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanCuti1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanCuti1", "Hak akses untuk mengakses form laporan cuti karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanOvertimeAnalysis1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanOvertimeAnalysis1", "Hak akses untuk mengakses form laporan analisis lembur", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanStandarCompetence1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanStandarCompetence1", "Hak akses untuk mengakses form laporan data kompetensi karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanSalaryStaff1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanSalaryStaff1", "Hak akses untuk mengakses form laporan data gaji karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanIncomeTaxStaff1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanIncomeTaxStaff1", "Hak akses untuk mengakses form laporan data gaji karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_TaxBeginningBalance1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_TaxBeginningBalance1", "Hak akses untuk mengakses form setting data awal pajak", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiLoan_BB1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiLoan_BB1", "Hak akses untuk mengakses form setting data awal pinjaman karyawan", "", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiLoanPayment_BB1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiLoanPayment_BB1", "Hak akses untuk mengakses form setting data awal pembayaran pinjaman karyawan", "", 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.UpdateHakAkses()
            Me.UpdateHakAkses_HR()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    ''' <summary>
    ''' payroll 10 Maret 2015
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub FrmMasterJobPosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterJobPosition11.Click
        Try
            Dim ij As Core_Module.FrmMasterJobPosition
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterJobPosition)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterJobPosition
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKaryawan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKaryawan11.Click
        Try
            Dim ij As Core_Module.FrmMasterKaryawan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterKaryawan)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterKaryawan
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterWorkShift11.Click
        Try
            Dim ij As Core_Module.FrmMasterWorkShift
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterWorkShift)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterWorkShift
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_Allocate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterWorkShift_Allocate11.Click
        Try
            Dim ij As Core_Module.FrmMasterWorkShift_Allocate
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterWorkShift_Allocate)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterWorkShift_Allocate
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterPayrollAccess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterPayrollAccess11.Click
        Try
            Dim ij As Core_Module.FrmMasterPayrollAccess
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterPayrollAccess)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterPayrollAccess
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalaryFormula11.Click
        Try
            Dim ij As Core_Module.FrmMasterSalaryFormula
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterSalaryFormula)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterSalaryFormula
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Attendance11.Click
        Try
            Dim ij As HR.FrmTrans_Attendance
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Attendance)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()

                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Attendance
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Gaji11.Click
        Try
            Dim ij As HR.FrmTrans_Gaji
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Gaji)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Gaji
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiCuti_Pengajuan11.Click
        Dim ObjPengajuanCuti As New HR.ClsTrans_Cuti_Pengajuan
        Dim ObjUser As New Core_Module.ClsUser
        Try
            Dim ij As HR.FrmTransaksiCuti_Pengajuan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiCuti_Pengajuan)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTransaksiCuti_Pengajuan
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            If ObjPengajuanCuti.GetJmlUser(ObjUser.GetPrimary(FrmLogin.UsernameTextBox.Text)) <= 0 Then
                MsgBox("You don't have privilege to access this form.", MsgBoxStyle.Information, "Information")
                i.Dispose()
            Else
                i.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjPengajuanCuti = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Approval_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiCuti_Approval11.Click
        Try
            Dim ij As HR.FrmTransaksiCuti_Approval
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiCuti_Approval)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTransaksiCuti_Approval
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCuti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCuti11.Click
        Try
            Dim ij As Core_Module.FrmMasterCuti
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterCuti)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterCuti
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKompetensi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKompetensi11.Click
        Try
            Dim ij As Core_Module.FrmMasterKompetensi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterKompetensi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterKompetensi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanEmployeeData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanEmployeeData11.Click
        Try
            Dim ij As HR.FrmLaporanEmployeeData
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanEmployeeData)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanEmployeeData
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanCuti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanCuti11.Click
        Try
            Dim ij As HR.FrmLaporanCuti
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanCuti)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanCuti
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanStandarCompetence_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanStandarCompetence11.Click
        Try
            Dim ij As HR.FrmLaporanStandarCompetence
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanStandarCompetence)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanStandarCompetence
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan_Setting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiCuti_Pengajuan_Setting11.Click
        Try
            Dim ij As HR.FrmTransaksiCuti_Pengajuan_Setting
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiCuti_Pengajuan_Setting)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTransaksiCuti_Pengajuan_Setting
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EmployeeCompetenceDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKompetensi11.Click
        Try
            Dim ij As Core_Module.FrmMasterKompetensi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterKompetensi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterKompetensi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EmployeePositionDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterJobPosition11.Click
        Try
            Dim ij As Core_Module.FrmMasterJobPosition
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterJobPosition)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterJobPosition
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EmployeeDataToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKaryawan11.Click
        Try
            Dim ij As Core_Module.FrmMasterKaryawan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterKaryawan)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterKaryawan
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLoan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiLoan11.Click
        Try
            Dim ij As HR.FrmTransaksiLoan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiLoan)
                If IsNothing(ij) = False Then
                    If ij.EnumDipanggilDari = HR.ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New HR.FrmTransaksiLoan
            i.EnumDipanggilDari = HR.ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLoan_Payment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiLoan_Payment11.Click
        Try
            Dim ij As HR.FrmTransaksiLoan_Payment
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiLoan_Payment)
                If IsNothing(ij) = False Then
                    If ij.EnumDipanggilDari = HR.ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari.FrmTransLoanPayment Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New HR.FrmTransaksiLoan_Payment
            i.EnumDipanggilDari = HR.ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari.FrmTransLoanPayment
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub JurnalSalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Purchase
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Purchase)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmTransaksiJournal_Purchase
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MasterSalaryFormulaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalaryFormula11.Click
        Try
            Dim ij As FrmMasterSalaryFormula
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterSalaryFormula)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterSalaryFormula
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_DataImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Attendance_DataImport11.Click
        Try
            Dim ij As HR.FrmTrans_Attendance_DataImport
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Attendance_DataImport)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Attendance_DataImport
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Quick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Gaji_Quick11.Click
        Try
            Dim ij As HR.FrmTrans_Gaji_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Gaji_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Gaji_Quick
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_OtherDeduction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalaryFormula_OtherDeduction11.Click
        Try
            Dim ij As Core_Module.FrmMasterSalaryFormula_OtherDeduction
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterSalaryFormula_OtherDeduction)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterSalaryFormula_OtherDeduction
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_TaxBeginningBalance_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FrmTrans_TaxBeginningBalance11.Click
        Try
            Dim ij As HR.FrmTrans_BeginingBalanceTax
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_BeginingBalanceTax)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_BeginingBalanceTax
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Payment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FrmTrans_Payment11.Click
        Try
            Dim ij As HR.FrmTransaksiPayment
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiPayment)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTransaksiPayment
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Export_ESpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FrmTrans_Export_eSPT11.Click
        Try
            Dim ij As HR.FrmTrans_Export_eSPT
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Export_eSPT)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Export_eSPT
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanAtt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanAtt11.Click
        Try
            Dim ij As HR.FrmLaporanAtt
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanAtt)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New HR.FrmLaporanAtt
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanOvertimeAnalysis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanOvertimeAnalysis11.Click
        Try
            Dim ij As HR.FrmLaporanOvertimeAnalysis
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanOvertimeAnalysis)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New HR.FrmLaporanOvertimeAnalysis
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalaryStaff_IGI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalaryStaff11.Click
        Try
            Dim ij As HR.FrmLaporanSalaryStaff_SMESS
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanSalaryStaff_SMESS)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New HR.FrmLaporanSalaryStaff_SMESS
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporan_IncomeTax_IGI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanIncomeTaxStaff11.Click
        Try
            Dim ij As HR.FrmLaporanIncomeTaxStaff
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanIncomeTaxStaff)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New HR.FrmLaporanIncomeTaxStaff
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSlipDescription_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSlipDescription11.Click
        Try
            Dim ij As Core_Module.FrmMasterSalaryFormula_SlipDesc
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterSalaryFormula_SlipDesc)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New Core_Module.FrmMasterSalaryFormula_SlipDesc
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmWaitForm11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As HR.WaitForm
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.WaitForm)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New HR.WaitForm
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Payment11.Click
        Try
            Dim ij As HR.FrmTransaksiPayment
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiPayment)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New HR.FrmTransaksiPayment
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLoan_BB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiLoan_BB11.Click
        Try
            Dim ij As HR.FrmTransaksiLoan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiLoan)
                If IsNothing(ij) = False Then
                    If ij.EnumDipanggilDari = HR.ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New HR.FrmTransaksiLoan
            i.Text = "Loan Beginning Balance Data"
            i.EnumDipanggilDari = HR.ClsTransaksiLoan.EnumFrmTransLoanDari.FrmTransLoan_Beginning_Balance
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLoanPayment_BB11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiLoanPayment_BB11.Click
        Try
            Dim ij As HR.FrmTransaksiLoan_Payment
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTransaksiLoan_Payment)
                If IsNothing(ij) = False Then
                    If ij.EnumDipanggilDari = HR.ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari.FrmTransLoanPayment_Beginning_Balance Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New HR.FrmTransaksiLoan_Payment
            i.Text = "Loan Payment Beginning Balance Data"
            i.EnumDipanggilDari = HR.ClsTransaksiLoanPayment.EnumFrmTransLoanPaymentDari.FrmTransLoanPayment_Beginning_Balance
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakAR1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmCetakAR1.Click
        Try
            Dim ij As AccModule.FrmCetakAR
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmCetakAR)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmCetakAR
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmCetakAP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmCetakAP1.Click
        Try
            Dim ij As AccModule.FrmCetakAP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmCetakAP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmCetakAP
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalesPersonToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesPerson1.Click
        Try
            Dim ij As FrmMasterSalesPerson
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterSalesPerson)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterSalesPerson
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesBudget1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesBudget1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesBudget
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSalesBudget)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSalesBudget
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalesPersonGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesPersonGroup1.Click
        Try
            Dim ij As FrmMasterSalesPersonGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterSalesPersonGroup)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterSalesPersonGroup
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalesPerformanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesPerformance1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesPerformance
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesPerformance)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesPerformance
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmEmailSettings1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmEmailSettings1.Click
        Try
            Dim ij As FrmEmailSettings
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmEmailSettings)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmEmailSettings
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Data11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_MaterialPlanningDashboard_Data11.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_MaterialPlanningDashboard_Data1
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_MaterialPlanningDashboard_Data1)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_Data1
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Override1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Gaji_Override1.Click
        Try
            Dim ij As HR.FrmTrans_Gaji_Override
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Gaji_Override)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Gaji_Override
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.ShowIcon = False
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanVAT21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanVAT21.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanVAT2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanVAT2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Purchase_Module.FrmLaporanVAT2
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterUser_Otorisasi1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterUser_Otorisasi1.Click
        Try
            Dim ij As Core_Module.FrmLaporanOtorisasiTransaksi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmLaporanOtorisasiTransaksi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Core_Module.FrmLaporanOtorisasiTransaksi
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.ObjMonitor.StartMonitor(FrmLogin.UsernameTextBox.Text, My.Computer.Name, Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonitor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiMonitor.Click
        Try
            Dim ij As Core_Module.FrmTransaksiMonitor
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmTransaksiMonitor)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Core_Module.FrmTransaksiMonitor
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False
            i.FromMainMenu = True
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MainPanelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MainPanelToolStripMenuItem.Click

        Try
            Dim ij As FrmMainSMESS
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMainSMESS)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Sub

    Private Sub FrmTransactionQueue1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransactionQueue1.Click
        Try
            Dim ij As Core_Module.FrmTransaksiMonitor
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmTransaksiMonitor)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Core_Module.FrmTransaksiMonitor
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False
            i.FromMainMenu = True
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAttendanceServer1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterAttendanceServer1.Click
        Try
            Dim ij As HR.FrmMasterAttendanceServer
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmMasterAttendanceServer)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmMasterAttendanceServer
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_Log1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Attendance_Log1.Click
        Try
            Dim ij As HR.FrmTrans_Attendance_Log
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Attendance_Log)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Attendance_Log
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.backcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_Actual_In_Report1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Attendance_Actual_In_Report1.Click
        Try
            Dim ij As HR.FrmTrans_Attendance_Actual_In_Report
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Attendance_Actual_In_Report)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Attendance_Actual_In_Report
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_Verifikasi1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmMasterProduct_Verifikasi1.Click
        Try
            Dim ij As frmMasterProduct_Verifikasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, frmMasterProduct_Verifikasi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New frmMasterProduct_Verifikasi
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanAkunProduct1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanAkunProduct1.Click
        Try
            Dim ij As AccModule.FrmLaporanAkunProduct
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanAkunProduct)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmLaporanAkunProduct
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmPurchaseOrderReportByDateDetail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmPurchaseOrderReportByDateDetail1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseOrderReportByDateDetail
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseOrderReportByDateDetail)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseOrderReportByDateDetail
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSalesOrderReportByDateDetail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmSalesOrderReportByDateDetail1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesOrderReportByDateDetail
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesOrderReportByDateDetail)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesOrderReportByDateDetail
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseByInvoice1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseByInvoice1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseByInvoice
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseByInvoice)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseByInvoice
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanARPaid1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanARPaid1.Click
        Try
            Dim ij As AccModule.FrmLaporanARPaid
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanARPaid)
                If IsNothing(ij) = False Then
                    If ij.TipeKartu = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanARPaid
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.TipeKartu = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanAPPaid1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanAPPaid1.Click
        Try
            Dim ij As AccModule.FrmLaporanARPaid
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmLaporanARPaid)
                If IsNothing(ij) = False Then
                    If ij.TipeKartu = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmLaporanARPaid
            'i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.TipeKartu = AccModule.ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub FrmSettingUpdate1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmSettingUpdate1.Click
    '    Dim ij As SMESS_Upload.FrmSettingUpdate
    '    Dim j As Windows.Forms.Form

    '    For Each j In Me.MdiChildren
    '        ij = TryCast(j, SMESS_Upload.FrmSettingUpdate)
    '        If IsNothing(ij) = False Then
    '            ij.BringToFront()
    '            ij.Focus()
    '            Exit Sub
    '        End If
    '    Next

    '    Dim i As New SMESS_Upload.FrmSettingUpdate

    '    i.MdiParent = Me
    '    i.StartPosition = FormStartPosition.CenterParent
    '    i.Show()
    'End Sub

    'Private Sub FrmDownload1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmDownload1.Click
    '    Dim ObjAuto As New Download_Module.ClsUpdateProgram
    '    Dim Hasil As MsgBoxResult

    '    Try
    '        If ObjAuto.ConnectionInternet(Download_Module.ClsUpdateProgram.EnumFungsi._download, True) = True Then
    '            If ObjAuto.RunProsesDownloadUpdate(Download_Module.ClsUpdateProgram.EnumFungsi._download, True) = True Then
    '                Hasil = MsgBox("Download ?", MsgBoxStyle.YesNo, "Confirm")
    '                If Hasil = MsgBoxResult.Yes Then
    '                    Dim proc As New System.Diagnostics.Process()
    '                    proc = Process.Start(CurrentDirectory & "\Download Module.exe")
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjAuto = Nothing
    '    End Try
    'End Sub

    'Private Sub CheckForUpdateLAN(ByVal vShowMsg As Boolean)
    '    Dim ObjAuto As New Download_Module.ClsUpdateProgram
    '    Dim Obj1 As New Download_Module.ClsAutoUpdate
    '    Try
    '        If System.IO.Directory.Exists(Obj1.GetServerPath) = False Then
    '            If vShowMsg = True Then
    '                MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
    '            End If
    '        Else
    '            If ObjAuto.RunProsesDownloadUpdate(Download_Module.ClsUpdateProgram.EnumFungsi._UpdateViaServer, vShowMsg) = True Then
    '                MsgBox("Please wait, we are going to update your software.", MsgBoxStyle.Information, "Information")
    '                Dim proc As New System.Diagnostics.Process()
    '                proc = Process.Start(CurrentDirectory & "\Client Update Module.exe")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjAuto = Nothing
    '        Obj1 = Nothing
    '    End Try
    'End Sub

    'Private Sub UploadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmUpdateFromLAN1.Click
    '    Me.CheckForUpdateLAN(True)
    'End Sub

    Private Sub FrmTransaksiDebetNote1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiDebetNote1.Click
        Try
            Dim ij As AccModule.FrmTransaksiCreditNote
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiCreditNote)
                If IsNothing(ij) = False Then
                    If ij.TipeTransaksi = AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New AccModule.FrmTransaksiCreditNote
            i.TipeTransaksi = AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCreditNote1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiCreditNote1.Click
        Try
            Dim ij As AccModule.FrmTransaksiCreditNote
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiCreditNote)
                If IsNothing(ij) = False Then
                    If ij.TipeTransaksi = AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New AccModule.FrmTransaksiCreditNote
            i.TipeTransaksi = AccModule.ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiRecuring_Reminder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiRecuring_Reminder1.Click
        Try
            Dim ij As AccModule.FrmTransaksiRecuring_Reminder
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiRecuring_Reminder)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiRecuring_Reminder
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
            If i.dgvHeader.Rows.Count = 0 Then
                MsgBox("No recurring reminder available.", MsgBoxStyle.Information, "Information")
                i.Close()
            Else
                i.Close()
                'MsgBox("You've got a reminder for recurring journal.", MsgBoxStyle.Information, "Information")
                ' Increment the caption counter.
                Dim A As New AccModule.FrmTransaksiRecuring_Reminder
                A.UserCode = FrmLogin.UsernameTextBox.Text
                ' Set the caption to be unique.
                A.BackcolorAll = Core_Module.GetSMESSBackColorAll
                A.MdiParent = Me
                A.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiRecuring_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiRecuring1.Click
        Try
            Dim ij As AccModule.FrmTransaksiRecuring
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiRecuring)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New AccModule.FrmTransaksiRecuring
            ' Increment the caption counter.
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductAccountGroup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProductAccountGroup1.Click
        Try
            Dim ij As Core_Module.FrmMasterProductAccountGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterProductAccountGroup)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterProductAccountGroup
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKompetensi_Product1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKompetensi_Product1.Click
        Try
            Dim ij As FrmMasterKompetensi_Product
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterKompetensi_Product)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterKompetensi_Product
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterDownTime1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterDownTime1.Click
        Try
            Dim ij As Core_Module.FrmMasterDownTime
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterDownTime)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterDownTime
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterMachineGroup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterMachineGroup1.Click
        Try
            Dim ij As Core_Module.FrmMasterMachineGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterMachineGroup)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterMachineGroup
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterMachineData1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterMachineData1.Click
        Try
            Dim ij As Core_Module.FrmMasterMachineData
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterMachineData)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterMachineData
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterReject1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterReject1.Click
        Try
            Dim ij As Core_Module.FrmMasterReject
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterReject)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterReject
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkUnit1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterWorkUnit1.Click
        Try
            Dim ij As Core_Module.FrmMasterWorkUnit
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterWorkUnit)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterWorkUnit
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPR1.Click
        Try
            Dim ij As New Purchase_Module.FrmTransaksiPR
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPR)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New Purchase_Module.FrmTransaksiPR
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPR_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPR_Cancel1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPR_Cancel
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPR_Cancel)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPR_Cancel
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPR_PickUpOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPR_PickUpOrder1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPR_PickUpOrder
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPR_PickUpOrder)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPR_PickUpOrder
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterRefParty(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterRefParty1.Click
        Try
            Dim ij As FrmMasterRefParty
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterRefParty)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterRefParty
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporan_Purchase_Request_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporan_Purchase_Request1.Click
        Try

            Dim ij As Purchase_Module.FrmLaporan_Purchase_Request
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporan_Purchase_Request)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporan_Purchase_Request
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporan_Purchase_Request_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporan_Purchase_Request_Cancel1.Click
        Try

            Dim ij As Purchase_Module.FrmLaporan_Purchase_Request_Cancel
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporan_Purchase_Request_Cancel)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporan_Purchase_Request_Cancel
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporan_Purchase_Request_PickUp_Order_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporan_Purchase_Request_PickUp_Order1.Click
        Try

            Dim ij As Purchase_Module.FrmLaporan_Purchase_Request_PickUp_Order
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporan_Purchase_Request_PickUp_Order)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporan_Purchase_Request_PickUp_Order
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporan_PR_Balance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporan_PR_Balance1.Click
        Try

            Dim ij As Purchase_Module.FrmLaporan_PR_Balance
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporan_PR_Balance)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporan_PR_Balance
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmProductListReport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmProductListReport1.Click
        Try
            Dim ij As Core_Module.FrmMasterProductListReport
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterProductListReport)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterProductListReport
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanStockBundling1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanStockBundling1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_List
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_List)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_List
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOC_QuickDel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOC_QuickDel1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOC_QuickDel
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOC_QuickDel)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOC_QuickDel
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOC1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOC1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOC
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOC)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOC
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
