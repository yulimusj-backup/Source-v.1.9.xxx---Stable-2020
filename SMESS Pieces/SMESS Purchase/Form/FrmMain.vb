Option Explicit On
Option Strict Off

Imports Core_Module
Imports System.IO

Public Class FrmMain
    Public EW As E_mail_Module.ClsEmailWatcher
    Dim ObjComapanyData As New Core_Module.ClsMasterCompanyData
    Dim CloseForm As Boolean = False
    Dim ObjLicense As New ClsLicense
    Dim ObjHakAkses As New Core_Module.ClsHakAkses

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
            Me.NotifyIcon1.Visible = False
            Me.NotifyIcon1.Dispose()
            Me.EW.Enabled = False
            Me.EW.Dispose()
            Me.EW = Nothing
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
        'Dim ObjUpdateDB As New ClsUpdateDatebase
        Dim ObjRecurringReminder As New AccModule.ClsTransaksiRecurring
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
        'Dim ObjUpdate As New Download_Module.ClsAutoUpdate
        'Dim ObjVersion As New ClsVersionControl

        Try
            'MsgBox("Starting as base.", MsgBoxStyle.Information, "Information")

            'pastikan versi, untuk installer paling baru
            'If ObjVersion.GetVersion() = "" Then
            'ObjVersion.SetVersion("1.6.40")
            'End If
            'If ObjVersion.GetVersion_ClientUpdate() = "" Then
            '    ObjVersion.SetVersion_ClientUpdate("1.6.10")
            'End If
            'check for update via LAN
            ' Me.CheckForUpdateLAN(False)

            'start monitor
            'Me.Identifier = Me.ObjRandom.Next(0, 10 ^ 7)
            'Me.ObjMonitor.StartMonitor(FrmLogin.UsernameTextBox.Text, My.Computer.Name, _
            'Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)
            'Me.Timer1.Interval = Me.ObjMonitor.MonitorInterval
            'Me.Timer1.Enabled = True
            'If Me.ObjMonitor.ApakahLicenseSMESSCukup() = False Then
            '    MsgBox("Insufficient SMESS active license, please release some license.", _
            '    MsgBoxStyle.Information, "Information")
            '    Dim i As New Core_Module.FrmTransaksiMonitor
            '    i.ShowDialog()
            '    Me.Close()
            '    Exit Sub
            'End If

            Me.InventoryTransactionToolStripMenuItem.Text = "Purchase"
            Me.SalesTransactionToolStripMenuItem.Text = "Sales"
            Me.InventoryTransactionToolStripMenuItem1.Text = "Inventory"
            'Me.HRD_Trans_ToolStripMenuItem4.Text = "HR"
            Me.AccountingTransactionToolStripMenuItem.Text = "Accounting"

            'Me.Icon = My.Resources.desktop_icon_3_fw
            'cek versi database
            'If My.Settings.db_version <> ObjDB.GetDBVersion Then
            '    If ObjUpdateDB.UpdateDatabase(ObjDB.GetDBVersion) = False Then
            '        Me.CloseForm = True
            '        Me.Close()
            '        Exit Sub
            '    End If
            'End If

            'cek hak akses yang belum ada di dalam database
            Control.CheckForIllegalCrossThreadCalls = False
            'Me.BackgroundWorker1.RunWorkerAsync()
            If ObjMasterPriorityLevel.GetJml = 0 Then
                ObjMasterPriorityLevel.simpanBaru(1, "R", "Reguler", "USER")
                ObjMasterPriorityLevel.setDefaultByPrimary("USER", ClsMasterPriorityLevel.enumDefault.yes, ObjMasterPriorityLevel.getPrimary("R"))
            End If

            'cek visibilitas
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'load custom report form
            Dim ObjCR As New CustomSMESS.ClsTrans_CustomReport(FrmLogin.UsernameTextBox.Text)
            'masukkan hak akses untuk menu custom report
            'ObjCR.UpdateMenu()
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
            'Me.ObjCD.CekVisibilitasMenu(Me.MenuStrip1, FrmLogin.UsernameTextBox.Text)

            ''jika tidak ada stock taking aktif, maka stock taking process tidak perlu ditampilkan
            'Dim ObjST As New Inventory_Module.clsTransaksiStockTaking
            'If ObjST.CekApakahAdaYangMasihAktif(False) = False Then
            '    FrmTransaksiStockTakingProcess1.Visible = False
            '    frmTransaksiStockTaking_Print1.Visible = False
            'End If
            'ObjST = Nothing
            ''default menu yg muncul dan yg tidak muncul
            'Me.AturMenuDefault()

            'activates email watcher, untuk sementara email watcher dimatikan dulu, karen mengganggu proses debugging
            'EW = New E_mail_Module.ClsEmailWatcher(FrmLogin.UsernameTextBox.Text, Me, Me.NotifyIcon1, _
            'Me.AlertMeOnIncomingEmailToolStripMenuItem)
            'If My.Settings.Email_Wathcer = "1" Then
            '    EW.Enabled = True
            '    AlertMeOnIncomingEmailToolStripMenuItem.Checked = True
            'Else
            '    EW.Enabled = False
            '    AlertMeOnIncomingEmailToolStripMenuItem.Checked = False
            'End If

            Dim ijk As New FrmMainSMESS
            ijk.MdiParent = Me
            'ijk.BackColorAll = Core_Module.GetSMESSBackColorAll
            ijk.Show()

            CompanyName = Me.ObjComapanyData.GetCompanyName
            If CompanyName <> "" Then
                CompanyName = " - " + CompanyName
            End If
            Me.Text = "SMESS - Ver. 1.6.00" + _
               CompanyName + " - Connection: " + Core_Module.GetDBServer + "(" + Core_Module.GetDBName + ")"

            'coding berkaitan dengan licensing
            My.Settings.a = Me.ObjLicense.GetProductKey
            My.Settings.b = Me.ObjLicense.GetLicenseKey(Today, My.Settings.a, ClsLicense.EnumTipeLicense.Normal1)

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
            'Me.MenuStrip1.Items.Remove(POS_Menu)
            'Me.AccountingTransactionToolStripMenuItem.DropDownItems.Remove(Me.FixedAssetToolStripMenuItem)
            'Me.AccReportToolStripMenuItem.DropDownItems.Remove(FrmLapAsset_DepresiasiPerPeriod1)
            ''Me.AccountingTransactionToolStripMenuItem.DropDownItems.Remove(PaToolStripMenuItem)
            ''Me.SalesTransactionToolStripMenuItem.DropDownItems.Remove(FrmTransaksiSalesReturn_Quick1)
            ''Me.DataMasterToolStripMenuItem.DropDownItems.Remove(FrmMasterDepartement1)
            'Me.DataMasterToolStripMenuItem.DropDownItems.Remove(FrmMasterJob1)
            ''Me.ToolsToolStripMenuItem.DropDownItems.Remove(CheckForUpdateToolStripMenuItem)
            'Me.VersionControlToolStripMenuItem.Visible = True
            'Me.FrmUpdateFromLAN1.Visible = True
            ResetStockLMB2016ToolStripMenuItem.Visible = True
            If ObjIdentifier.GetIdentifier = "smess_lmb" Then
                UpdateM3LMBToolStripMenuItem.Visible = True
            End If

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
            'If ObjRecurringReminder.GetJmlReminderAktif(Today, Me.ObjUser.GetPrimary(FrmLogin.UsernameTextBox.Text)) > 0 Then
            '    FrmTransaksiRecuring_Reminder1.PerformClick()
            'End If

            'update nama user pada core module
            Core_Module.SetUser(FrmLogin.UsernameTextBox.Text)
            Me.MonthlyClosingToolsToolStripMenuItem.Visible = True
            ImportEFakturToolStripMenuItem.Visible = True
            'MsgBox("final update autoupdate.")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            obj = Nothing
            ObjLS = Nothing
            ObjReminder = Nothing
            ObjDB = Nothing
            ObjMasterPriorityLevel = Nothing
            'ObjUpdateDB = Nothing
            ObjRecurringReminder = Nothing
            'ObjVersion = Nothing
        End Try
    End Sub

    Private Sub ShowLicenseToEnter()
        Dim i As New FrmTransaksiLicensing

        Try

            'cek kembali, jika masih bermasalah maka tutup aplikasi
            If My.Settings.a <> Me.ObjLicense.GetProductKey Then
                If Me.ObjLicense.GetTipeLicense(My.Settings.b) = ClsLicense.EnumTipeLicense.Trial Then
                    My.Settings.a = Me.ObjLicense.GetProductKey
                    My.Settings.b = Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) - TimeSpan.FromDays(35), _
                    My.Settings.a, ClsLicense.EnumTipeLicense.Trial)
                    My.Settings.Save()

                    i.ShowDialog()
                Else
                    i.ShowDialog()

                    MsgBox("Your product key is invalid. This can happen if you change your hardware. Please contact your software vendor to obtain new license key.", _
                    MsgBoxStyle.Information, "Information")
                    Me.CloseForm = True
                    Me.Close()
                End If
            End If

            '2. cek expiry date
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Function CekTollStripMenu(ByVal i As ToolStripMenuItem, ByVal TBL As DSUser.TblHakAksesUserDataTable) As Boolean
        Dim obj As New ClsUser
        Dim j As ToolStripMenuItem
        Dim vs As Boolean = False
        Dim row As DSUser.TblHakAksesUserRow
        Dim int As Integer

        Try
            If i.DropDownItems.Count > 0 Then
                For int = 0 To i.DropDownItems.Count - 1
                    j = TryCast(i.DropDownItems.Item(int), ToolStripMenuItem)
                    If IsNothing(j) = False Then
                        If Me.CekTollStripMenu(j, TBL) Then
                            vs = True
                        End If
                    End If
                Next

                i.Visible = vs
                Return vs
            Else
                For Each row In TBL
                    If row.Nama = i.Name.Substring(0, i.Name.Length - 1) Then
                        i.Visible = True
                        Return True
                    End If
                    i.Visible = False
                Next
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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
            i.MainMenuStrip1 = Me.MenuStrip1
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        Try
            If MsgBox("Exit SMESS?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.CloseForm = True
                FrmLogin.Close()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            ' i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiGRN1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiGRN1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiGRN
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiGRN)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiGRN
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            ' i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseRequisition1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseRequisition1.Click
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.ObjStockBundling.TipeBundling = Inventory_Module.ClsTransaksiStockBundling.EnumTipeBundling.Purchase1
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            ' i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            Else
                Me.EW.Stop()
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AuthorizePOPurchaseStaffToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv11.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOOtorisasi2)
                If IsNothing(ij) = False Then
                    If ij.Text = "Authorize Purchase Order - Purchase Staff" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOOtorisasi2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 0
            i.CancelAuthorizationLevel = 1
            ' Set the caption to be unique.
            i.Text = "Authorize Purchase Order - Purchase Staff"
            'back color
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AuthorizePOPurchaseManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv21.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOOtorisasi2)
                If IsNothing(ij) = False Then
                    If ij.Text = "Authorize Purchase Order - Purchase Manager" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOOtorisasi2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 1
            i.CancelAuthorizationLevel = 2
            ' Set the caption to be unique.
            i.Text = "Authorize Purchase Order - Purchase Manager"
            'back color
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AuthorizePODirectorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv31.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPOOtorisasi2)
                If IsNothing(ij) = False Then
                    If ij.Text = "Authorize Purchase Order - Director" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiPOOtorisasi2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 2
            i.CancelAuthorizationLevel = 3
            ' Set the caption to be unique.
            i.Text = "Authorize Purchase Order - Director"
            'back color
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockInAdj
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me

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
            i.MdiParent = Me

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
            ' i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me

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
            ' i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me

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
            ' i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me

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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me

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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesListByInvoice1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesListByInvoice1.Click
        Try
            Dim ij As Sales_Module.FrmSalesReport
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmSalesReport)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmSalesReport
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            ' i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            'i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceDP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesInvoiceDP1.Click
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.BackcolorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            'i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasiSalesStaff1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiQuotationOtorisasiSalesStaff1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiQuotationOtorisasi2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiQuotationOtorisasi2)
                If IsNothing(ij) = False Then
                    If ij.AuthorizationLevel = 0 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiQuotationOtorisasi2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 0
            i.CancelAuthorizationLevel = 1
            ' Set the caption to be unique.
            i.Text = i.Text & " - Sales Staff Authorization"

            i.BackColorAll = Color.LightSkyBlue

            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasiSalesManager1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiQuotationOtorisasiSalesManager1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiQuotationOtorisasi2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiQuotationOtorisasi2)
                If IsNothing(ij) = False Then
                    If ij.AuthorizationLevel = 1 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiQuotationOtorisasi2
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 1
            i.CancelAuthorizationLevel = 2
            ' Set the caption to be unique.
            i.Text = i.Text & " - Sales Manager Authorization"
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiQuotationOtorisasiGeneralManager1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiQuotationOtorisasiGeneralManager1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiQuotationOtorisasi2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiQuotationOtorisasi2)
                If IsNothing(ij) = False Then
                    If ij.AuthorizationLevel = 2 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiQuotationOtorisasi2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.AuthorizationLevel = 2
            i.CancelAuthorizationLevel = 3
            ' Set the caption to be unique.
            i.Text = i.Text & " - General Manager Authorization"
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            'i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSettingInvoice1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiSettingInvoice
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiSettingInvoice)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiSettingInvoice
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSettingInvoiceDP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiSettingInvoiceDP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiSettingInvoiceDP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiSettingInvoiceDP
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSettingSalesReturn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiSettingSR
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiSettingSR)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiSettingSR
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSettingAOP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiSettingAOP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiSettingAOP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiSettingAOP
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesCycle22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesCycle22.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Sales
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Sales)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_Sales
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.SalesReturnPaymentJournal
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Purchase21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Purchase21.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Purchase
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal_Purchase)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal_Purchase
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.PurchaseReturnReceivement
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Adjust1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Adjust1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiJournal
            i.UserCOde = FrmLogin.UsernameTextBox.Text
            i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
            ' Set the caption to be unique.
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaskiSettingAOPR1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiSettingAOPR
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiSettingAOPR)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New AccModule.FrmTransaksiSettingAOPR
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.CloseForm = False Then
                If MsgBox("Exit SMESS?", MsgBoxStyle.YesNo, "Confimation") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
            i.BackColorAll = Color.LightSkyBlue
            i.MdiParent = Me
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
            ' i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TransactionNumberingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterNumbering1.Click
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

            ' Increment the caption counter.
            Dim i As New FrmMasterNumbering
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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
            i.Lokal = True
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAsset1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterAsset1.Click
        Dim i As New FrmMasterAsset
        i.ShowDialog()
    End Sub

    Private Sub MaterialUsageToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaterialUsageToolStripMenuItem1.Click
        Dim i As New Purchase_Module.FrmTransaksiAsset_ServiceUsage
        i.ShowDialog()
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanVAT2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanVAT2.Click
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
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

    Private Sub FrmTransaksiAOP_CreateByStockIn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAOP_CreateByStockIn1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiAOP_CreateByStockIn
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiAOP_CreateByStockIn)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTransaksiAOP_CreateByStockIn
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorALL = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MonthlyClosingToolsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MonthlyClosingToolsToolStripMenuItem.Click
        Try
            Dim ij As Purchase_Module.Form2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.Form2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.Form2
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorALL = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ResetStockLMB2016ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetStockLMB2016ToolStripMenuItem.Click
        Try
            Dim ij As New Inventory_Module.Form2
            ij.MdiParent = Me
            ij.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateM3LMBToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateM3LMBToolStripMenuItem.Click
        'Dim ij As CustomSMESS.FrmTransaksiUpdateVol
        'Dim j As Windows.Forms.Form

        'For Each j In Me.MdiChildren
        '    ij = TryCast(j, CustomSMESS.FrmTransaksiUpdateVol)
        '    If IsNothing(ij) = False Then
        '        ij.BringToFront()
        '        ij.Focus()
        '        Exit Sub
        '    End If
        'Next

        ''Static intCount As Integer
        'Dim i As New CustomSMESS.FrmTransaksiUpdateVol
        'i.UserCode = FrmLogin.UsernameTextBox.Text
        'i.BackColorAll = Core_Module.GetSMESSBackColorAll
        'i.MdiParent = Me
        'i.Text = "01. Update m3"
        'i.Show()
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
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
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

    Private Sub VATReportBinayaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VATReportBinayaToolStripMenuItem.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanVAT2_Custom
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanVAT2_Custom)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanVAT2_Custom
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

    Private Sub ImportCheckDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportCheckDataToolStripMenuItem.Click
        Try
            Dim ij As Purchase_Module.FrmImport_CheckData
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmImport_CheckData)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmImport_CheckData
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

    Private Sub ImportEFakturToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportEFakturToolStripMenuItem.Click
        Try
            Dim ij As Sales_Module.FrmImport_eFaktur
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmImport_eFaktur)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmImport_eFaktur
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
End Class