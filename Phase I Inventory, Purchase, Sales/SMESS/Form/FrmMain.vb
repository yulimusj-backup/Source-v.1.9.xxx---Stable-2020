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
    Dim ObjAutoDB As New Core_Module.ClsMasterBackUpDB
    'Dim NextExpiryDate As Date = #4/30/2019#
    Dim VersionCheckFreq As Integer = 0
    Dim StartVersionCheck As Boolean = False

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

    Private Function GetTipeLicenseRandom() As ClsLicense.EnumTipeLicense
        Dim i As Integer = 0
        Try
            i = Me.ObjRandom.Next(0, 9)
            Select Case i
                Case 0
                    Return ClsLicense.EnumTipeLicense.Normal1
                Case 1
                    Return ClsLicense.EnumTipeLicense.Normal2
                Case 2
                    Return ClsLicense.EnumTipeLicense.Normal3
                Case 3
                    Return ClsLicense.EnumTipeLicense.Normal4
                Case 4
                    Return ClsLicense.EnumTipeLicense.Normal4
                Case 5
                    Return ClsLicense.EnumTipeLicense.Normal5
                Case 6
                    Return ClsLicense.EnumTipeLicense.Normal6
                Case 7
                    Return ClsLicense.EnumTipeLicense.Normal7
                Case 8
                    Return ClsLicense.EnumTipeLicense.Normal8
                Case 9
                    Return ClsLicense.EnumTipeLicense.Normal9
            End Select

            Return ClsLicense.EnumTipeLicense.Normal1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsLicense.EnumTipeLicense.Normal1
        End Try
    End Function

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
        Dim ObjVersion As New ClsVersionControl
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
        Dim ObjVersionControl As New Core_Module.Clstrans_VersionControl
        Dim ObjCustom As New CustomSMESS.ClsIdentifier

        Try
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

            'Me.Icon = My.Resources.desktop_icon_3_fw
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
            ObjCR.UpdateMenu(False)
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
            Me.Text = "SMESS - Ver. " + ObjVersion.GetVersion + _
               CompanyName + " - Connection: " + Core_Module.GetDBServer + "(" + Core_Module.GetDBName + ")"

            'coding berkaitan dengan licensing
            '29/045/2020 by ahdar input quick license manager update
            '1. cek sudah pernah login atau belum
            If My.Settings.a <> "" Then
                Dim ObjVer As New Core_Module.Clstrans_VersionControl
                Dim NewLicenseKey As String = ""
                '2. cek komputer blacklist atau bukan
                If ObjVer.ApakahKompIsBlacklist(My.Settings.a) = True Then
                    MsgBox("The license in this computer is blacklisted. Please contact your vendor if this is not wanted.", MsgBoxStyle.Information, "Information")
                    Me.CloseForm = True
                    Me.Close()
                    Exit Sub
                End If

                NewLicenseKey = ObjVer.GetNewLicensiExpiry(My.Settings.a)
                If String.IsNullOrEmpty(NewLicenseKey) = False Then
                    'jika license tidak kosong maka set sebagai license baru
                    'artinya setiap login acuannya dari database bukan dari my.setting lagi 
                    My.Settings.b = NewLicenseKey
                    My.Settings.Save()
                Else
                    'jika kosong maka update sesuai license lama
                    ObjVer.UpdateNewLicenseKey(My.Settings.b, My.Settings.a)
                End If
            End If
            '

            'coding berkaitan dengan licensing
            If My.Settings.a = "" Then
                'masuk trial period
                My.Settings.a = Me.ObjLicense.GetProductKey
                My.Settings.b = Me.ObjLicense.GetLicenseKey(Today, My.Settings.a, ClsLicense.EnumTipeLicense.Trial)
                My.Settings.c = 400
                My.Settings.Save()

                MsgBox("This is a trial aplication. Please enter license key to get full aplication function", MsgBoxStyle.Information, "Information")
                Me.Text = Me.Text + " (Trial)"

                'coding berkaitan dengan licensing
                '29/045/2020 by ahdar input quick license manager update
                '1. cek sudah pernah login atau belum
                If My.Settings.a <> "" Then
                    Dim ObjVer As New Core_Module.Clstrans_VersionControl
                    Dim NewLicenseKey As String = ""
                    '2. cek komputer blacklist atau bukan
                    If ObjVer.ApakahKompIsBlacklist(My.Settings.a) = True Then
                        MsgBox("This computer not registered as a smess member.", MsgBoxStyle.Information, "Information")
                        Me.CloseForm = True
                        Me.Close()
                        Exit Sub
                    End If

                    NewLicenseKey = ObjVer.GetNewLicensiExpiry(My.Settings.a)
                    If String.IsNullOrEmpty(NewLicenseKey) = False Then
                        'jika license tidak kosong maka set sebagai license baru
                        'artinya setiap login acuannya dari database bukan dari my.setting lagi 
                        My.Settings.b = NewLicenseKey
                        My.Settings.Save()
                    Else
                        'jika kosong maka update sesuai license lama
                        ObjVer.UpdateNewLicenseKey(My.Settings.b, My.Settings.a)
                    End If
                End If
                '
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
                    If ObjCustom.GetIdentifier() = "smess_tbu" Then
                        'temporary fix u/ topaz
                        Dim Expiry As Date = Me.ObjLicense.GetExpiryDate(My.Settings.b)
                        If Today.Date >= Expiry.Date Then
                            If Expiry <= #12/31/2019# Then
                                'langsung ditimpa aj licensenya
                                My.Settings.b = Me.ObjLicense.GetLicenseKey(Expiry + TimeSpan.FromDays(1), _
                                My.Settings.a, Me.ObjLicense.GetTipeLicense(My.Settings.b) + 1, #12/31/2019#)
                                My.Settings.Save()
                                Sales_Module.SetLicense(My.Settings.b)
                                Purchase_Module.SetLicense(My.Settings.b)
                                Inventory_Module.SetLicense(My.Settings.b)
                                Core_Module.SetLicense(My.Settings.b)
                                AccModule.SetLicense(My.Settings.b)
                                HR.ModGlobalVariable.SetLicense(My.Settings.b)
                            Else
                                MsgBox("Your license has expired. Please contact your software vendor to obtain new license key.", _
                                MsgBoxStyle.Information, "Information")
                                Me.ShowLicenseToEnter()
                                'End If
                            End If
                        End If
                    Else
                        If Today > Me.ObjLicense.GetExpiryDate(My.Settings.b) Then
                            MsgBox("Your license has expired. Please contact your software vendor to obtain new license key.", _
                            MsgBoxStyle.Information, "Information")
                            Me.ShowLicenseToEnter()
                            'End If
                        End If
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
                        'End If
                    End If
                End If
            End If

            Core_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Inventory_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Purchase_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Sales_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            AccModule.ModGlobalVariable.SetLicense(My.Settings.b)
            E_mail_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            HR.ModGlobalVariable.SetLicense(My.Settings.b)

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
            Me.AccReportToolStripMenuItem.DropDownItems.Remove(Me.FrmLaporanVAT1)
            Me.AccReportToolStripMenuItem.DropDownItems.Remove(Me.FrmLaporanChequeAging1)
            Me.AccReportToolStripMenuItem.DropDownItems.Remove(Me.FrmLaporanChequePerPeriod1)

            Me.AccReportToolStripMenuItem.DropDownItems.Remove(FrmLapAsset_DepresiasiPerPeriod1)
            Me.SalesReportToolStripMenuItem.DropDownItems.Remove(FrmLaporanCommission1)
            Me.SalesReportToolStripMenuItem.DropDownItems.Remove(FrmLaporanQuotationDetail1)
            Me.SalesReportToolStripMenuItem.DropDownItems.Remove(FrmLaporanQuotationRealisasi1)
            Me.SalesReportToolStripMenuItem.DropDownItems.Remove(FrmLaporanQuotationRealisasiChart1)
            Me.SalesReportToolStripMenuItem.DropDownItems.Remove(FrmLaporanQuotationTrendChart1)
            'Me.ManajemenUserToolStripMenuItem.DropDownItems.Remove(FrmTrans_CutOffDB1)

            'If ObjIdentifier.GetIdentifier <> "smess_gemilang" Then
            '    Me.FrmMasterSalesPerson11.DropDownItems.Remove(FrmMasterCommission1)
            'End If
            Me.DataMasterToolStripMenuItem.DropDownItems.Remove(FrmMasterJob1)
            Me.VersionControlToolStripMenuItem.Visible = True
            Me.FrmUpdateFromLAN1.Visible = True

            'HR Transaksi
            'Me.DataMasterToolStripMenuItem.DropDownItems.Remove(Me.HumanResourceToolStripMenuItem)
            'Me.MenuStrip1.Items.Remove(Me.HRD_Trans_ToolStripMenuItem4)
            'Me.ReportsToolStripMenuItem3.DropDownItems.Remove(Me.HRD_Report_ToolStripMenuItem1)
            'Me.ToolsToolStripMenuItem.DropDownItems.Remove(FrmTrans_TaxBeginningBalance11)
            'Me.ToolsToolStripMenuItem.DropDownItems.Remove(LoanBeginningBalanceToolStripMenuItem)

            My.Settings.Save()

            'pastikan versi, untuk installer paling baru
            If ObjVersion.GetVersion() = "" Then
                ObjVersion.SetVersion("1.6.40")
            End If
            'check for update via LAN
            Me.StartVersionCheck = True
            Me.CheckForUpdateLAN(False)

            'start auto back up
            ObjAutoDB.StartBackUpScheduler(FrmLogin.UsernameTextBox.Text)
            'If FrmLogin.UsernameTextBox.Text = "USER" Then
            '    Me.CekAutoBackUpToolStripMenuItem.Visible = True
            'Else
            '    Me.FileToolStripMenuItem.DropDownItems.Remove(Me.CekAutoBackUpToolStripMenuItem)
            'End If

            'update setting untuk inventory movement
            Purchase_Module.UpdateSettingsPO()

            'jika user ybs memiliki reminder jurnal, maka keluarkan remindernya.
            If ObjRecurringReminder.GetJmlReminderAktif(Today, Me.ObjUser.GetPrimary(FrmLogin.UsernameTextBox.Text)) > 0 Then
                FrmTransaksiRecuring_Reminder1.PerformClick()
            End If

            'update nama user pada core module
            Core_Module.SetUser(FrmLogin.UsernameTextBox.Text)

            Me.ProductionTransactionToolStripMenuItem.Text = "Production"

            'Dim ij As New Core_Module.FrmUserLocationPrivilegeSetting
            'ij.UserCode = "USER"
            'ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            ''ij.ObjStockBundling.TipeBundling = Inventory_Module.ClsTransaksiStockBundling.EnumTipeBundling.Purchase
            'ij.ShowDialog()

            'form yang sementara dihide dulu/sudah obselete

            'kontrol menu untuk memastikan semua menu ada
            If FrmLogin.UsernameTextBox.Text.ToLower = "user" Then
                UpdateAllAccessToolStripMenuItem.Visible = True
            Else
                ManajemenUserToolStripMenuItem.DropDownItems.Remove(UpdateAllAccessToolStripMenuItem)
            End If
            If ObjIdentifier.GetIdentifier = "smess_lmb" And AccountingStartUpSettingToolStripMenuItem.Visible = False Then
                Me.AktifkanBalanceAR()
            End If

            If ObjIdentifier.GetIdentifier <> "smess_lmb" Then
                Me.AccountingTransactionToolStripMenuItem.DropDownItems.Remove(Me.MonthlyClosingTools1)
            End If

            If FrmLogin.UsernameTextBox.Text = "USER" Then
                Me.UpdateUtilityToolStripMenuItem.Visible = True
                FrmTransaksiLicensing21.Visible = True
                Me.UpdateAllAccessToolStripMenuItem.Visible = True
            Else
                Me.ToolsToolStripMenuItem.DropDownItems.Remove(Me.UpdateUtilityToolStripMenuItem)
                FrmTransaksiLicensing21.Visible = False
                Me.UpdateAllAccessToolStripMenuItem.Visible = False
            End If

            'update login time
            ObjVersionControl.SetLoginTimeStamp(My.Settings.a, FrmLogin.UsernameTextBox.Text, My.Computer.Name, Me.ObjLicense.GetExpiryDate(My.Settings.b))
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
            ObjVersion = Nothing
            ObjIdentifier = Nothing
            ObjVersionControl = Nothing
            ObjCustom = Nothing
        End Try
    End Sub

    Private Sub AktifkanBalanceAR()
        Dim ObjClosing As New Purchase_Module.ClsTransaksiMonthlyClosing

        Try
            'jika jumlah closing > jan 2018 = 0
            'aktifkan menu balance ar
            If ObjClosing.GetLastClosingPeriod.Date < #1/1/2018# Then
                AccountingStartUpSettingToolStripMenuItem.Visible = True
                FrmTransaksiSalesBalanceManual1.Visible = True
                FrmTransaksiSalesReturnBalanceManual1.Visible = True
                Me.FrmTransaksiPurchaseBalanceManual1.Visible = False
                Me.FrmTransaksiPurchaseReturnBalanceManual1.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub AturMenuDefault()
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
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
                Me.ToolsToolStripMenuItem.Visible = True
                FrmTransaksiMonitor.Visible = True
                Me.FrmPersonalSetting1.Visible = True
            End If

            If Me.ObjClosing.GetJml > 0 Then
                If ObjIdentifier.GetIdentifier = "smess_lmb" And AccountingStartUpSettingToolStripMenuItem.Visible = False Then
                    Me.AktifkanBalanceAR()
                Else
                    Me.AccountingStartUpSettingToolStripMenuItem.Visible = False
                    Me.FrmTransaksiInventoryBeginingBalance1.Visible = False
                    Me.FrmTransaksiStockInAdj_SetAmount1.Visible = False
                End If
            End If

            'Me.FrmSettingUpdate1.Visible = True
            'UploadToolStripMenuItem.Visible = True
            'FrmDownload1.Visible = True
            'VersionControlToolStripMenuItem.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjIdentifier = Nothing
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
            i.ObjStockBundling.TipeBundling = Inventory_Module.ClsTransaksiStockBundling.EnumTipeBundling.Purchase1
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
            i.ShowIcon = False
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
            i.ShowIcon = False
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

    Private Sub FrmTransaksiSOC1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSOCHeader.Click
        
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
        Dim LastVersionTxt As String = ""
        Dim DefaultVersion As String = "1.6.40"
        Dim OperatingVersion As String = ""

        Try
            Me.AturMenuDefault()

            If FrmLogin.UsernameTextBox.Text = "USER" Then
                If Core_Module.GetDBServer = "localhost" Or Core_Module.GetDBServer = "cien-asss" Then
                    Exit Sub
                End If
            End If

            If Me.VersionCheckFreq < 3 And Me.StartVersionCheck = True Then
                'MsgBox("Starting version check", MsgBoxStyle.Information, "Information")

                Dim ObjSetting As New Core_Module.ClsTransaksiSetting
                Dim ObjVersion As New ClsVersionControl
                Dim ObjUpdate1 As New ClsUpdateProgram
                Dim ObjVersionControl As New Core_Module.Clstrans_VersionControl

                LastVersionTxt = ObjVersion.GetVersion(DefaultVersion)
                OperatingVersion = ObjVersionControl.GetOperatingVersion
                If ObjUpdate1.ApakahVersionBerbeda(LastVersionTxt, OperatingVersion) = True Then
                    MsgBox("Operating version is: " + OperatingVersion + ", please contact your system administrator for version update.", MsgBoxStyle.Information, "Information")
                    Me.CloseForm = True
                    FrmLogin.Close()
                    Me.Close()

                    ObjSetting = Nothing
                    ObjVersion = Nothing
                    ObjUpdate1 = Nothing
                    ObjVersionControl = Nothing

                    Exit Sub
                End If

                Me.VersionCheckFreq += 1

                ObjSetting = Nothing
                ObjVersion = Nothing
                ObjUpdate1 = Nothing
                ObjVersionControl = Nothing
            End If
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

    Private Sub UpdateNumbering_Produksi()
        Dim objNum As New Core_Module.ClsMasterTransactionNumbering

        Try
            If objNum.UpdateNumbering("tbltrans_spl", "Production Overtime Order", _
                       Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL) = False Then
                Exit Sub
            End If
            If objNum.UpdateNumbering("tbltrans_mo_item_wo", "Work Order", _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.WO) = False Then
                Exit Sub
            End If
            If objNum.UpdateNumbering("tbltrans_mo_item", "Job Order", _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO) = False Then
                Exit Sub
            End If
            If objNum.UpdateNumbering("tbltrans_mo", "Manufacturing Order", _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO) = False Then
                Exit Sub
            End If
            If objNum.UpdateNumbering("tbltrans_sch_downtime", "Schedule Downtime", _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = False Then
                Exit Sub
            End If
            If objNum.UpdateNumbering("tbltransaksi_stockbundling", "Production Log", _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling) = False Then
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objNum = Nothing
        End Try
    End Sub

    Private Sub UpdateHakAkses(Optional ByVal vExitImmidiately As Boolean = True)
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjSalaryForm As New Core_Module.ClsMasterSalaryFormula
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
        Dim ObjAntrianCustom As New CustomSMESS.ClsAntrian_Custom  '2018-08-26: untuk memastikan setiap custom, memiliki antrian custom. source ada di lmb

        Try
            Me.ObjHakAkses.SetConnectionString(DTA)

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Quick_Otorisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Quick_Otorisasi", "Hak untuk membuat delivery order yang dibuat user lain.", "FrmTransaksiDO_Quick", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_Quick", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_Quick", "Hak akses untuk input, edit, dan hapus data Quick DO")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterTaxFacilityStampReference", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterTaxFacilityStampReference", "Hak akses untuk input, edit, dan hapus data facility stamp")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterTaxAddInfoReference", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterTaxAddInfoReference", "Hak akses untuk input, edit, dan hapus data add info")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanVAT2025") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanVAT2025", "Hak akses untuk melihat laporan VAT CoreTax 2025.", , )
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterTaxProductReference") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterTaxProductReference", "Hak akses untuk input, edit dan hapus master tax product reference.", , )
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterTaxMeasurementReference") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterTaxMeasurementReference", "Hak akses untuk input, edit dan hapus master tax unit of measurement reference.", , )
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoice_HakAksesEditHapus") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoice_HakAksesEditHapus", "Hak akses untuk edit dan hapus dengan user lain", "FrmTransaksiSalesInvoice", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If ObjIdentifier.GetIdentifier() = "smess_binaya" Then
                If Me.ObjHakAkses.ApakahSudahAda("FrmImport_Art") = False Then
                    Me.ObjHakAkses.Simpan("FrmImport_Art", "Hak akses untuk melihat laporan VAT Report Custom.", "FrmImport_Art", 0)
                Else
                    If vExitImmidiately = True Then Exit Sub
                End If

                If Me.ObjHakAkses.ApakahSudahAda("FrmImport_eFaktur") = False Then
                    Me.ObjHakAkses.Simpan("FrmImport_eFaktur", "Hak akses untuk melihat laporan VAT Report Custom.", "FrmImport_eFaktur", 0)
                Else
                    If vExitImmidiately = True Then Exit Sub
                End If

                If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanVAT2_Custom") = False Then
                    Me.ObjHakAkses.Simpan("FrmLaporanVAT2_Custom", "Hak akses untuk melihat laporan VAT Report Custom.", "FrmLaporanVAT2_Custom", 0)
                Else
                    If vExitImmidiately = True Then Exit Sub
                End If
            End If


            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoice_QuickByStockOut") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoice_QuickByStockOut", "Hak akses untuk membuat sales invoice otomatis berdasarkan stock out sales.", "FrmTransaksiSalesInvoice", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSOC_QuickDel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSOC_QuickDel", "Hak akses membuat transaksi sales order cancellation sekaligus.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanSalesDP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanSalesDP", "Hak akses untuk melihat laporan sales down payment.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanPurchaseDP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanPurchaseDP", "Hak akses untuk melihat laporan purchase down payment.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockInAdj_SetAmount", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockInAdj_SetAmount", "Hak akses untuk input dan edit begining nilai amount pada balance inventory", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_AddItembyCustom", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_AddItembyCustom", "Hak akses untuk input, edit, dan delete product berdasarkan custom input", "FrmTransaksiDO", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            ''custom admiral +taufik 'tidak jadi karena gabung dengan AddByCustomProduct
            'If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiAOP_Quick2_AddItembyCustom") = False Then
            '    Me.ObjHakAkses.Simpan("FrmTransaksiAOP_Quick2_AddItembyCustom", "Hak akses untuk input, edit dan delete product berdasarkan custom input", "FrmTransaksiAOP_Quick", 0)
            'Else
            '    If vExitImmidiately = True Then Exit Sub
            'End If
            ''custom admiral +aul 'tidak jadi karena gabung dengan AddByCustomProduct
            'If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoiceQuick_AddItembyCustom") = False Then
            '    Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoiceQuick_AddItembyCustom", "Hak akses untuk input, edit dan delete product berdasarkan custom input", "FrmTransaksiSalesInvoiceQuick", 0)
            'Else
            '    If vExitImmidiately = True Then Exit Sub
            'End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling_AddByCustomProduct", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockBundling", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockInAdj_AddItembyCustom") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockInAdj_AddItembyCustom", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockInAdj", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockOutAdj_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockOutAdj_AddByCustomProduct", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockOutAdj", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockOutSales_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockOutSales_AddByCustomProduct", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockOutSales", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockOutPurchaseReturn2_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockOutPurchaseReturn2_AddByCustomProduct", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockOutPurchaseReturn", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockTransfer__AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockTransfer__AddByCustomProduct", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockPindah", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockInPurchase_Custom") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockInPurchase_Custom", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockInPurchase", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'custom admiral
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockInSalesReturn_Custom") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockInSalesReturn_Custom", "Hak akses untuk menambahkan product berdasarkan custom input.", "FrmTransaksiStockInSalesReturn", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_AddByBarcode") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_AddByBarcode", "Hak akses untuk menambahkan product berdasarkan barcode pada DO.", "FrmTransaksiDO", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoiceQuick_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoiceQuick_AddByCustomProduct", "Hak akses untuk input, edit dan delete item invoice quick berdasarkan custom product", "FrmTransaksiSalesInvoiceQuick", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesReturn_Quick_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesReturn_Quick_AddByCustomProduct", "Hak akses untuk input, edit dan delete item sales return quick berdasarkan custom product", "FrmTransaksiSalesReturn_Quick", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPurchaseReturn_Quick_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPurchaseReturn_Quick_AddByCustomProduct", "Hak akses untuk input, edit dan delete item purchase return quick berdasarkan custom product", "FrmTransaksiPurchaseReturn_Quick", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiAOP_Quick2_AddByCustomProduct") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiAOP_Quick2_AddByCustomProduct", "Hak akses untuk input, edit dan delete item AOP quick berdasarkan custom product", "FrmTransaksiAOP_Quick", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterKaryawan_Kompetensi_Upd") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterKaryawan_Kompetensi_Upd", "Hak akses untuk update kompetensi karyawan.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Jurnal_VerifyProductAccount) = "" Then
                ObjSetting.SetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Jurnal_VerifyProductAccount, "1")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterCustomer_Reference") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterCustomer_Reference", "Hak akses untuk input, edit, dan hapus master customer reference.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanCustomerList") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanCustomerList", "Hak akses untuk melihat customer")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Gaji_MasaSblm") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Gaji_MasaSblm", "Hak akses untuk input, edit, dan hapus pendapatan netto masa sebelum.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDO_EditShipTo", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDO_EditShipTo", "Hak akses untuk mengedit alamat kirim pada DO", "FrmTransaksiDO", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPO_EditPriceStlhInv", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPO_EditPriceStlhInv", "Hak untuk mengubah harga setelah item so dirujuk oleh invoice supplier/approval of purchase.", "FrmTransaksiPO", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSO_EditPriceStlhInv", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSO_EditPriceStlhInv", "Hak akses untuk mengubah harga setelah item so dirujuk oleh invoice", "FrmTransaksiSO", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("frmMasterProduct_InputWeightNVolume") = False Then
                Me.ObjHakAkses.Simpan("frmMasterProduct_InputWeightNVolume", "Hak akses untuk input data berat dan volume pada suatu product.", "FrmMasterProduct", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("frmMasterProduct_ChangeControl") = False Then
                Me.ObjHakAkses.Simpan("frmMasterProduct_ChangeControl", "Hak akses untuk mengubah kontrol batch, expiry dan lokasi pada suatu product.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            'If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockOutAdj_QuickInputProductLevel") = False Then -> hold dulu sampe beneran siap dan tidak stgh matang
            '    Me.ObjHakAkses.Simpan("FrmTransaksiStockOutAdj_QuickInputProductLevel", "Hak akses untuk input data item stock out adjust menggunakan level produk.", "FrmTransaksiStockOutAdj", 0)
            'Else
            '    If vExitImmidiately = True Then Exit Sub
            'End If
            'If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockOutAdj_QuickInputBatchNo") = False Then -> hold dulu sampe beneran siap dan tidak stgh matang
            '    Me.ObjHakAkses.Simpan("FrmTransaksiStockOutAdj_QuickInputBatchNo", "Hak akses untuk input data item stock out adjust menggunakan data batch no.", "FrmTransaksiStockOutAdj", 0)
            'Else
            '    If vExitImmidiately = True Then Exit Sub
            'End If
            'If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockInAdj_QuickInputProductLevel") = False Then -> hold dulu sampe beneran siap dan tidak stgh matang
            '    Me.ObjHakAkses.Simpan("FrmTransaksiStockInAdj_QuickInputProductLevel", "Hak akses untuk input data item stock in adjust menggunakan level produk.", "FrmTransaksiStockInAdj", 0)
            'Else
            '    If vExitImmidiately = True Then Exit Sub
            'End If 

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiExpiryDateUpdate_Merge") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiExpiryDateUpdate_Merge", "Hak akses untuk menggabungkan data batch suatu product.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockTransfer_Receive_FullAkses") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockTransfer_Receive_FullAkses", "Hak akses untuk receive & cancel received", "FrmTransaksiStockTransfer_Receive", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSalaryFormula_Detail") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSalaryFormula_Detail", "Hak akses untuk input, edit, hapus salary formula detail.", "FrmMasterSalaryFormula1", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockTransfer_Receive") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockTransfer_Receive", "Hak akses untuk set status transit transaksi transfer lokasi.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockPindah_QuickInputBarcode") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockPindah_QuickInputBarcode", "Hak akses untuk input data item stock transfer menggunakan barcode.", "FrmTransaksiStockPindah", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockPindah_QuickInputDocRef") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockPindah_QuickInputDocRef", "Hak akses untuk input data item stock transfer menggunakan data dokumen lainnya.", "FrmTransaksiStockPindah", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling_QuickInputBarcode") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling_QuickInputBarcode", "Hak akses untuk input data item stock bundling menggunakan barcode.", "FrmTransaksiStockBundling", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling_QuickInputDocRef") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling_QuickInputDocRef", "Hak akses untuk input data item stock bundling menggunakan data dokumen lainnya.", "FrmTransaksiStockBundling", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockOutAdj_QuickInputBarcode") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockOutAdj_QuickInputBarcode", "Hak akses untuk input data item stock out adjust menggunakan barcode.", "FrmTransaksiStockOutAdj", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockOutAdj_QuickInputDocRef") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockOutAdj_QuickInputDocRef", "Hak akses untuk input data item stock out adjust menggunakan data dokumen lainnya.", "FrmTransaksiStockOutAdj", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockInAdj_QuickInputBarcode") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockInAdj_QuickInputBarcode", "Hak akses untuk input data item stock in adjust menggunakan barcode.", "FrmTransaksiStockInAdj", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockInAdj_QuickInputDocRef") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockInAdj_QuickInputDocRef", "Hak akses untuk input data item stock in adjust menggunakan data dokumen lainnya.", "FrmTransaksiStockInAdj", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiBalanceManual_DPSls") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiBalanceManual_DPSls", "Hak akses untuk input, edit, delete transaksi historical down payment sales balance.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiBalanceManual_DPPch") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiBalanceManual_DPPch", "Hak akses untuk input, edit, delete transaksi historical down payment purchase balance.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If ObjIdentifier.GetIdentifier = "smess_lmb" Then
                If Me.ObjHakAkses.ApakahSudahAda("MonthlyClosingTools") = False Then
                    Me.ObjHakAkses.Simpan("MonthlyClosingTools", "Hak akses untuk membuat laporan excel monthly closing.", "", 0)
                Else
                    If vExitImmidiately = True Then Exit Sub
                End If
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling2_OverlapSchedule") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling2_OverlapSchedule", "Hak akses untuk input log produksi dengan waktu produksi yang saling beririsan.", "FrmTransaksiStockBundling2", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling2_OverlapSchedule") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling2_OverlapSchedule", "Hak akses untuk input log produksi dengan waktu produksi yang saling beririsan.", "FrmTransaksiStockBundling2", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoice_OverridePrice") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoice_OverridePrice", "Hak akses untuk edit harga dan diskon pada transaksi invoice.", "FrmTransaksiSalesInvoice", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_StockBundling_Quick") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_StockBundling_Quick", "Hak akses untuk input transaksi production log berdasarkan nomor transaksi MO.", "FrmTrans_Attendance1", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Attendance_Spl") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Attendance_Spl", "Hak akses untuk mengubah duration lembur yang dibuat dari overtime order.", "FrmTrans_Attendance1", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPLApproval_Authorize") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPLApproval_Authorize", "Hak akses untuk otorisasi overtime order.", "FrmTrans_SPLApproval", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPLApproval_ShowAll") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPLApproval_ShowAll", "Hak akses untuk melihat semua transaksi overtime order.", "FrmTrans_SPLApproval", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_EditDetailOT") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_EditDetailOT", "Hak akses untuk edit detail ot transaksi surat perintah lembur.", "FrmTrans_SPL", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMaster_SPL_EmployeeCustom") = False Then
                Me.ObjHakAkses.Simpan("FrmMaster_SPL_EmployeeCustom", "Hak akses form setting otorisasi surat perintah lembur.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            Me.UpdateNumbering_Produksi()


            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterKaryawan_LihatKompetensi") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterKaryawan_LihatKompetensi", "Hak akses untuk melihat data kompetensi pada form master karyawan.", "FrmMasterKaryawan1", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Setting") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Setting", "Hak akses form setting otorisasi surat perintah lembur.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanOvertimeOrder") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanOvertimeOrder", "Hak akses form laporan surat perintah lembur.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPLApproval") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPLApproval", "Hak akses form approval surat perintah lembur.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL", "Hak akses form transaksi surat perintah lembur.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_New") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_New", "Hak akses untuk input transaksi surat perintah lembur.", "FrmTrans_SPL", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Edit") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Edit", "Hak akses untuk edit transaksi surat perintah lembur.", "FrmTrans_SPL", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Hapus") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Hapus", "Hak akses untuk hapus transaksi surat perintah lembur.", "FrmTrans_SPL", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Cetak") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Cetak", "Hak akses untuk cetak transaksi surat perintah lembur.", "FrmTrans_SPL", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_EditUserLain", "Hak akses untuk edit transaksi surat perintah lembur yang dibuat user lain.", "FrmTrans_SPL", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_HapusUserLain", "Hak akses untuk hapus transaksi surat perintah lembur yang dibuat user lain.", "FrmTrans_SPL", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Budget") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Budget", "Hak akses form transaksi budget lembur.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Budget_New") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Budget_New", "Hak akses untuk input transaksi budget lembur.", "FrmTrans_SPL_Budget", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Budget_Edit") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Budget_Edit", "Hak akses untuk edit transaksi budget lembur.", "FrmTrans_SPL_Budget", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Budget_Hapus") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Budget_Hapus", "Hak akses untuk hapus transaksi budget lembur.", "FrmTrans_SPL_Budget", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Budget_Cetak") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Budget_Cetak", "Hak akses untuk cetak transaksi budget lembur.", "FrmTrans_SPL_Budget", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Budget_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Budget_EditUserLain", "Hak akses untuk edit transaksi budget lembur yang dibuat user lain.", "FrmTrans_SPL_Budget", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SPL_Budget_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SPL_Budget_HapusUserLain", "Hak akses untuk hapus transaksi budget lembur yang dibuat user lain.", "FrmTrans_SPL_Budget", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_CashBank_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_CashBank_HapusUserLain", "Hak akses untuk hapus transaksi cash and bank yang dibuat user lain", "FrmTransaksiJournal_CashBank")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDebetNote_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDebetNote_HapusUserLain", "Hak akses untuk hapus transaksi debit note yang dibuat user lain", "FrmTransaksiDebetNote")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDebetNote_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDebetNote_EditUserLain", "Hak akses untuk edit transaksi debit note yang dibuat user lain", "FrmTransaksiDebetNote")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCreditNote_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCreditNote_HapusUserLain", "Hak akses untuk hapus transaksi credit note yang dibuat user lain", "FrmTransaksiCreditNote")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCreditNote_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCreditNote_EditUserLain", "Hak akses untuk edit transaksi credit note yang dibuat user lain", "FrmTransaksiCreditNote")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_Closing_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_Closing_HapusUserLain", "Hak akses untuk hapus transaksi end of year yang dibuat user lain", "FrmTransaksiJournal_Closing")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_Closing_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_Closing_EditUserLain", "Hak akses untuk edit transaksi end of year yang dibuat user lain", "FrmTransaksiJournal_Closing")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_EndOfMonth_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_EndOfMonth_HapusUserLain", "Hak akses untuk hapus transaksi end of month yang dibuat user lain", "FrmTransaksiJournal_EndOfMonth")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_EndOfMonth_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_EndOfMonth_EditUserLain", "Hak akses untuk edit transaksi end of month yang dibuat user lain", "FrmTransaksiJournal_EndOfMonth")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanPinjamanKaryawan") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanPinjamanKaryawan", "Hak akses untuk mengakses form laporan pinjaman karyawan")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_Purchase1_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_Purchase1_HapusUserLain", "Hak akses untuk hapus transaksi purchase payment yang dibuat user lain", "FrmTransaksiJournal_Purchase1")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_Purchase1_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_Purchase1_EditUserLain", "Hak akses untuk edit transaksi purchase payment yang dibuat user lain", "FrmTransaksiJournal_Purchase1")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesCycle1_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesCycle1_HapusUserLain", "Hak akses untuk hapus transaksi sales receivement yang dibuat user lain", "FrmTransaksiSalesCycle1")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesCycle1_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesCycle1_EditUserLain", "Hak akses untuk edit transaksi sales receivement yang dibuat user lain", "FrmTransaksiSalesCycle1")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_CashBank_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_CashBank_EditUserLain", "Hak akses untuk edit transaksi cash and bank yang dibuat user lain", "FrmTransaksiJournal_CashBank")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiGeneralJournal_HapusUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiGeneralJournal_HapusUserLain", "Hak akses untuk hapus transaksi general jurnal yang dibuat user lain", "FrmTransaksiGeneralJournal")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiGeneralJournal_EditUserLain") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiGeneralJournal_EditUserLain", "Hak akses untuk edit transaksi general jurnal yang dibuat user lain", "FrmTransaksiGeneralJournal")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanOtherPayment") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanOtherPayment", "Hak akses untuk mengakses form laporan other payment")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanStockBundling_LihatCost", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanStockBundling_LihatCost", "Hak akses untuk melihat cost pada laporan stock bundling.", "FrmLaporanStockBundling", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_PrintMOBlmOtorisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_PrintMOBlmOtorisasi", "Hak akses untuk print mo yang belum di otorisasi.", "FrmTrans_ManufacturingOrder", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_SF_Override_PReqETD", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_SF_Override_PReqETD", "Hak akses untuk mengedit tanggal janji pengiriman permintaan pembelian ke pihak terkait pada shipping forecast/material dashboard.", "FrmTrans_MaterialPlanningDashboard", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanSupplierList") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanSupplierList", "Hak akses untuk melihat Supplier")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanCustomerList") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanCustomerList", "Hak akses untuk melihat customer")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            'tambahan 5 sept 2017
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanKontrakKaryawan", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanKontrakKaryawan", "Hak akses untuk mengakses form laporan kontrak karyawan")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterKaryawan_TipeKontrak", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterKaryawan_TipeKontrak", "Hak akses untuk input, edit, dan delete master tipe kontrak")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_EditSpareQty", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_EditSpareQty", "Hak akses untuk mengubah qty spare.", "FrmTrans_MaterialPlanningDashboard", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("frmMasterProduct_BOM1") = False Then
                Me.ObjHakAkses.Simpan("frmMasterProduct_BOM1", "Hak akses untuk fitur edit bill of material secara keseluruhan")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_Authorize1") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_Authorize1", "Hak akses untuk otorisasi material planning dashboard level 1")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_Authorize2") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_Authorize2", "Hak akses untuk otorisasi material planning dashboard level 2")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_Authorize3") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_Authorize3", "Hak akses untuk otorisasi material planning dashboard level 3")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_AuthorizeRequired") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_AuthorizeRequired", "Material Planning Dashboard yang dibuat wajib di otorisasi.", "FrmTrans_MaterialPlanningDashboard")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_Track") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_Track", "Material Planning Dashboard track.", "FrmTrans_MaterialPlanningDashboard")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_CreateEditPO") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_CreateEditPO", "Hak akses untuk membuat atau menambahkan item pada PO", "FrmTrans_MaterialPlanningDashboard")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_CreateEditMO") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_CreateEditMO", "Hak akses untuk membuat atau menambahkan item pada MO.", "FrmTrans_MaterialPlanningDashboard")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_MP_EditOrderQty") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_MP_EditOrderQty", "Hak akses untuk mengubah qty order.", "FrmTrans_MaterialPlanningDashboard")
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_EditTargetFinish", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_EditTargetFinish", "Hak akses untuk input/edit tanggal target selesai mo.", "FrmTrans_ManufacturingOrder", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_AddItem", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_AddItem", "Hak akses untuk input/edit/hapus item baru transaksi manufacturing order.", "FrmTrans_ManufacturingOrder", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_EditTransaksi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_EditTransaksi", "Hak akses untuk edit transaksi manufacturing order yang dibuat user lain.", "FrmTrans_ManufacturingOrder", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_HapusTransaksi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_HapusTransaksi", "Hak akses untuk hapus transaksi manufacturing order yang dibuat user lain.", "FrmTrans_ManufacturingOrder", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_InputWIP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_InputWIP", "Hak akses untuk input,edit dan hapus item WIP.", "FrmTrans_ManufacturingOrder", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_InputAdditional", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_InputAdditional", "Hak akses untuk input,edit dan hapus item additional.", "FrmTrans_ManufacturingOrder", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanKontrakKaryawan") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanKontrakKaryawan", "Hak akses untuk membuat laporan kontrak kerja karyawan")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterMachineLine", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterMachineLine", "Hak akses untuk input, edit, dan delete data line produksi.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling_NoStrictLocation2", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling_NoStrictLocation2", "Hak akses untuk input item raw material dari lokasi diluar prioritas lokasi.", _
                "FrmTransaksiStockBundling2", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling_NoStrictLocation", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling_NoStrictLocation", "Hak akses untuk input item raw material dari lokasi diluar prioritas lokasi.", _
                "FrmTransaksiStockBundling", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ProductionSchedule_SendTakeControl", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ProductionSchedule_SendTakeControl", "Hak akses untuk melakukan penjadwalan produksi.", "FrmTrans_ProductionSchedule", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterWorkUnitGroup", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterWorkUnitGroup", "Hak akses untuk input, edit, dan delete data work unit group.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSetupTime", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSetupTime", "Hak akses untuk input, edit, dan delete data setup time.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterProductionRate", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterProductionRate", "Hak akses untuk input, edit, dan delete data referensi rate produksi.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_Authorize", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_Authorize", "Hak akses untuk mengotorisasi/batal otorisasi transaksi manufacturing order.", "", 0)
                Me.UpdateNumbering_Produksi()
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiStockBundling2", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiStockBundling2", "Hak akses untuk melakukan input, edit, hapus transaksi log produksi.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanProd_RMAllocation", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanProd_RMAllocation", "Hak akses untuk membuat laporan pick list.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanProd_PlanningReport", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanProd_PlanningReport", "Hak akses untuk membuat cetakan penjadwalan produksi.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanProd_RejectReport", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanProd_RejectReport", "Hak akses untuk membuat laporan reject.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanProd_RawMaterialUsage", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanProd_RawMaterialUsage", "Hak akses untuk membuat laporan penggunaan bahan baku.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanProd_Productivity", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanProd_Productivity", "Hak akses untuk membuat laporan produktivitas produksi.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanProd_Downtime", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanProd_Downtime", "Hak akses untuk membuat laporan downtime.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder_SetFinishJO", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder_SetFinishJO", "Hak akses untuk melakukan set finish job order.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ProductionSchedule", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ProductionSchedule", "Hak akses untuk melihat form penjadwalan produksi.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_ManufacturingOrder", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_ManufacturingOrder", "Hak akses untuk input, edit, dan delete transaksi manufacturing order.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_CutOffDB", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_CutOffDB", "Hak akses untuk melakukan pembagian database menjadi 2 berdasarkan periode tertentu.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanQuotationTrendChart", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanQuotationTrendChart", "Hak akses untuk melihat grafik trend penawaran penjualan.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanQuotationRealisasiChart", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanQuotationRealisasiChart", "Hak akses untuk melihat grafik penawaran penjualan.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanQuotationRealisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanQuotationRealisasi", "Hak akses untuk melihat laporan realisasi penawaran penjualan.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanQuotationDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanQuotationDetail", "Hak akses untuk melihat laporan penawaran penjualan.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Gaji_Override_FindAll", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Gaji_Override_FindAll", "Hak akses untuk melihat semua data yang diinput, termasuk yang dibuat oleh user lain.", "FrmTrans_Gaji_Override", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Gaji_OverrideA_FindAll", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Gaji_OverrideA_FindAll", "Hak akses untuk melihat semua data yang diinput, termasuk yang dibuat oleh user lain.", "FrmTrans_Gaji_OverrideA", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjHakAkses.ApakahSudahAda("FrmMasterCommission", DTA) = False Then
                ObjHakAkses.Simpan("FrmMasterCommission", "Hak akses untuk input, edit, dan delete master komisi.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_TaxRecalculate", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_TaxRecalculate", "Hak akses untuk input, edit, dan delete data perhitungan pajak.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Gaji_OverrideA", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Gaji_OverrideA", "Hak akses untuk input, edit, dan delete data other payment data.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjIdentifier.GetIdentifier() = "smess_gemilang" Then
                If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanCommission", DTA) = False Then
                    Me.ObjHakAkses.Simpan("FrmLaporanCommission", "Hak akses untuk melihat laporan komisi sales")
                Else
                    If vExitImmidiately = True Then Exit Sub
                End If
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanVAT2_Edit_FP_AOP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanVAT2_Edit_FP_AOP", "Hak akses untuk edit nomor faktur pajak AOP.", "FrmLaporanVAT2", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanVAT2_Edit_KodeFP_Inv", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanVAT2_Edit_KodeFP_Inv", "Hak akses untuk edit kode faktur pajak sales invoice.", "FrmLaporanVAT2", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanSalesReturnDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanSalesReturnDetail", "Hak akses untuk melihat laporan detail sales return.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanPurchaseReturnDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanPurchaseReturnDetail", "Hak akses untuk melihat laporan detail purchase return.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            Me.ObjHakAkses.SetConnectionString(DTA)
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanLogUser", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanLogUser", "Hak akses untuk melihat laporan perubahan pada salary formula.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If ObjSalaryForm.GetJmlParameter_ByPrimary(141) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(141, "Is Outstation", "nilai yes jika hari itu dinas luar", "", "yes/no", 0, 0, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSalaryForm.GetJmlParameter_ByPrimary(142) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(142, "Total Outstation Days Monthly", "Total outstation selama 1 bulan dalam hitungan hari", "", "Days", 1, 1, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSalaryForm.GetJmlParameter_ByPrimary(143) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(143, "Total Outstation Days Yearly", "Total outstation selama 1 tahun dalam hitungan hari", "", "Days", 2, 1, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSalaryForm.GetJmlParameter_ByPrimary(144) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(144, "Total Outstation Days Weekday Monthly", "Total outstation tipe weekday selama 1 bulan dalam hitungan hari", "", "Days", 1, 1, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSalaryForm.GetJmlParameter_ByPrimary(145) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(145, "Total Outstation Days Holiday Monthly", "Total outstation tipe holiday selama 1 bulan dalam hitungan hari", "", "Days", 1, 1, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSalaryForm.GetJmlParameter_ByPrimary(146) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(146, "Total Outstation Days Weekday Yearly", "Total outstation tipe weekday selama 1 tahun dalam hitungan hari", "", "Days", 2, 1, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSalaryForm.GetJmlParameter_ByPrimary(147) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(147, "Total Outstation Days Holiday Yearly", "Total outstation tipe holiday selama 1 tahun dalam hitungan hari", "", "Days", 2, 1, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSalaryForm.GetJmlParameter_ByPrimary(148) = 0 Then
                ObjSalaryForm.SimpanBaru_ParameterGaji(148, "Is Resign", "nilai yes jika karyawan telah tidak aktif", "", "yes/no", 1, 0, 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterProduct_ChangeAccountCode", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterProduct_ChangeAccountCode", "Hak akses untuk mengubah akun yang berhubungan dengan product yang bersangkutan.", _
                "FrmMasterProduct", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_Otorisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_Otorisasi", "Hak akses untuk mengotorisasi/ batal otorisasi Purch. Request.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiAOP_CreateByStockIn_CreateAOP", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiAOP_CreateByStockIn_CreateAOP", "Hak akses membuat transaksi aop berdasarkan stock in purchase yang sudah terjadi.", _
                "FrmTransaksiAOP_CreateByStockIn", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_LihatDetailPembelian", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_LihatDetailPembelian", "Hak akses untuk melihat detail pembelian pada transaksi material planning.", "FrmTrans_MaterialPlanningDashboard", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_OverrideEdit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_OverrideEdit", "Hak akses untuk mengedit transaksi purchase request yang dibuat oleh user lain.", "FrmTransaksiPR", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_OverrideDelete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_OverrideDelete", "Hak akses untuk menghapus transaksi purchase request yang dibuat oleh user lain.", "FrmTransaksiPR", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.UpdateFromLAN_RequireUAC, ClsTransaksiSetting.EnumNamaParameter.UpdateFromLAN_RequireUAC) = "" Then
                ObjSetting.SetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.UpdateFromLAN_RequireUAC, _
                ClsTransaksiSetting.EnumNamaParameter.UpdateFromLAN_RequireUAC, "1")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiAOP_CreateByStockIn", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiAOP_CreateByStockIn", "Hak akses untuk menampilkan semua stock yang belum di AOP.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmSalesReport", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmSalesReport", "Hak akses membuat laporan penjualan berdasarkan invoice penjualan.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanStockInOut", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanStockInOut", "Hak akses membuat laporan stock in/out.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanStockTransfer", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanStockTransfer", "Hak akses membuat laporan stock transfer.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPOC_QuickDel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPOC_QuickDel", "Hak akses membuat transaksi purchase order cancellation sekaligus.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_Override_Otorisasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_Override_Otorisasi", _
                "Hak akses untuk edit/hapus transaksi purchase request walaupun sudah diotorisasi.", _
                "FrmTransaksiPR", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanStockBundling", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanStockBundling", "Hak akses untuk melihat laporan stock bundling.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmUpdateFromLAN", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmUpdateFromLAN", "Hak akses untuk melakukan update via LAN.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmSettingUpdate", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmSettingUpdate", "Hak akses untuk mensetting lokasi folder yang digunakan untuk melakukan update via LAN.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmDownload", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmDownload", "Hak akses untuk mendownload update terakhir dari cloud.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmProductListReport", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmProductListReport", "Hak akses untuk laporan daftar product.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterMachineData", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterMachineData", "Hak akses untuk input, edit, dan hapus master mesin.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_PR_Balance", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_PR_Balance", "Hak akses untuk membuat laporan purchase request balance.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_Purchase_Request_PickUp_Order", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_Purchase_Request_PickUp_Order", "Hak akses untuk membuat laporan purchase request pick up order.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_Purchase_Request_Cancel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_Purchase_Request_Cancel", "Hak akses untuk membuat laporan purchase request cancellation.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporan_Purchase_Request", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporan_Purchase_Request", "Hak akses untuk membuat laporan purchase request.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterRefParty", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterRefParty", "Hak akses untuk input, edit, dan hapus master pihak yang melakukan purchase request.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_PickUpOrder", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_PickUpOrder", "Hak akses untuk input, edit, dan hapus transaksi purchase request pick up order.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR_Cancel", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR_Cancel", "Hak akses untuk input, edit, dan hapus transaksi purchase request cancellation.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiPR", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiPR", "Hak akses untuk input, edit, dan hapus transaksi purchase request.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterMachineGroup", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterMachineGroup", "Hak akses untuk input, edit, dan hapus master grup mesin.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterWorkUnit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterWorkUnit", "Hak akses untuk input, edit, dan hapus master work unit.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDownTime", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterDownTime", "Hak akses untuk input, edit, dan hapus master down time.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterReject", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterReject", "Hak akses untuk input, edit, dan hapus master reject.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterProductAccountGroup", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterProductAccountGroup", "Hak akses untuk input, edit, dan hapus master kelompok akun product.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterKompetensi_Product", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterKompetensi_Product", "Hak akses untuk input, edit, dan hapus master custom value untuk product.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiRecuring", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiRecuring", "Hak akses untuk input, edit, dan hapus transaksi recurring.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiRecuring_Reminder", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiRecuring_Reminder", "Hak akses untuk form reminder recurring.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDebetNote", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDebetNote", "Hak akses untuk input, edit, dan hapus transaksi debit note.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCreditNote1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCreditNote1", "Hak akses untuk input, edit, dan hapus transaksi credit note.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanPurchaseByInvoice", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanPurchaseByInvoice", "Hak akses untuk melihat laporan pembelian berdasarkan invoice supplier.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmSalesOrderReportByDateDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmSalesOrderReportByDateDetail", "Hak akses untuk melihat laporan detail sales order.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmPurchaseOrderReportByDateDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmPurchaseOrderReportByDateDetail", "Hak akses untuk melihat laporan detail purchase order.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanAkunProduct", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanAkunProduct", "Hak akses untuk melihat laporan akun produk.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("frmMasterProduct_Verifikasi", DTA) = False Then
                Me.ObjHakAkses.Simpan("frmMasterProduct_Verifikasi", "Hak akses untuk memverifikasi akun produk.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Attendance_Actual_In_Report", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Attendance_Actual_In_Report", "Hak akses untuk melihat data clock in karyawan.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_Attendance_Log", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_Attendance_Log", "Hak akses untuk melakukan import data dari server kehadiran.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterAttendanceServer", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterAttendanceServer", "Hak akses untuk input, edit, dan hapus data server kehadiran.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanAPPaid", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanAPPaid", "Hak akses untuk melihat laporan balance hutang.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanARPaid", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanARPaid", "Hak akses untuk melihat laporan balance piutang.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiDebetNote", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiDebetNote", "Hak akses untuk input, edit, dan delete transaksi debet note.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiCreditNote", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiCreditNote", "Hak akses untuk input, edit, dan delete transaksi credit note.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransactionQueue", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTransactionQueue", "Hak akses untuk mereset data dari tabel antrian.", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
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
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_New", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_New", "Hak akses untuk input transaksi baru material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Edit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Edit", "Hak akses untuk edit transaksi material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Delete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Delete", "Hak akses untuk hapus transaksi material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Print", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Print", "Hak akses untuk cetak transaksi material planning group.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_Apply", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_Apply", "Hak akses untuk mengaplikasikan data transaksi material planning group ke dalam parameter perhitungan.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_ShowDetail", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_ShowDetail", "Hak akses untuk menampilkan detail perhitungan data transaksi material planning group ke dalam parameter perhitungan.", "FrmTrans_MaterialPlanningDashboard_Data1", 1)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_OverrideEdit", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_OverrideEdit", "Hak akses untuk edit data transaksi material planning group yang dibuat user lain.", "FrmTrans_MaterialPlanningDashboard_Data1", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1_OverrideDelete", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1_OverrideDelete", "Hak akses untuk delete data transaksi material planning group yang dibuat user lain.", "FrmTrans_MaterialPlanningDashboard_Data1", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1", "Hak akses untuk membuka form pengelompokan perencanaan material")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmEmailSettings") = False Then
                Me.ObjHakAkses.Simpan("FrmEmailSettings", "Hak akses untuk input, edit, dan hapus transaksi setting email.")
            Else
                If vExitImmidiately = True Then Exit Sub
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanSalesPerformance") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanSalesPerformance", "Hak akses untuk melihat laporan performa tim sales")
            Else
                If vExitImmidiately = True Then Exit Sub
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
            ObjSetting = Nothing
            ObjSalaryForm = Nothing
            ObjIdentifier = Nothing
            ObjAntrianCustom = Nothing
        End Try
    End Sub

    Private Sub UpdateHakAkses_HR(Optional ByVal vExitImmidiately As Boolean = True)
        Dim DTA As New DSUserTableAdapters.tblmasterhakaksesTableAdapter

        Try
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterWorkShift1", DTA) = False Then
                Me.ObjHakAkses.Simpan("FrmMasterWorkShift1", "Hak akses untuk mengakses form master workshift", "", 0)
            Else
                If vExitImmidiately = True Then Exit Sub
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

    Private Sub FrmMasterSalaryFormula_OtherDeduction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
            Dim ij As HR.FrmLaporanOvertimeAnalysis2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanOvertimeAnalysis2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New HR.FrmLaporanOvertimeAnalysis2
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
                    If ij.Tipe = HR.ClsTransaksiSalary_Override.EnumType.SalaryOverride Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New HR.FrmTrans_Gaji_Override
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Tipe = HR.ClsTransaksiSalary_Override.EnumType.SalaryOverride
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

    Private Sub FrmSettingUpdate1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmSettingUpdate1.Click
        Dim ij As FrmSettingUpdate
        Dim j As Windows.Forms.Form

        For Each j In Me.MdiChildren
            ij = TryCast(j, FrmSettingUpdate)
            If IsNothing(ij) = False Then
                ij.BringToFront()
                ij.Focus()
                Exit Sub
            End If
        Next

        Dim i As New FrmSettingUpdate

        i.MdiParent = Me
        i.StartPosition = FormStartPosition.CenterParent
        i.Show()
    End Sub

    Private Sub FrmDownload1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmDownload1.Click
        Dim ObjAuto As New ClsUpdateProgram
        Dim Hasil As MsgBoxResult
        Dim FrmDL As New FrmDownload
        Dim FrmChk As New FrmAutoUpdate_Check

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If FrmChk.ShowDialog = Windows.Forms.DialogResult.OK Then
                Windows.Forms.Cursor.Current = Cursors.Arrow
                Hasil = MsgBox("We found a new version of SMESS, proceed download ?", MsgBoxStyle.YesNo, "Confirm")
                If Hasil = MsgBoxResult.Yes Then
                    If FrmDL.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        'If FrmChk.Download_SMESS = 1 Then
                        CheckForUpdateLAN(True)
                        'End If
                    Else
                        If FrmDL.Progres = FrmDownload.enumProgress.CancelledByUser Then
                            MsgBox("Process cancelled by user.", MsgBoxStyle.Information, "Information")
                        Else
                            If FrmDL.Progres = FrmDownload.enumProgress.Start Then
                                MsgBox("Something went wrong, please try again.", MsgBoxStyle.Information, "Information")
                            ElseIf FrmDL.Progres = FrmDownload.enumProgress.CheckFolder Then
                                MsgBox("Something went wrong during creating buffer folder, please try again.", MsgBoxStyle.Information, "Information")
                            ElseIf FrmDL.Progres = FrmDownload.enumProgress.CopyFile Then
                                MsgBox("Something went wrong during copying file to target folder, please try again.", MsgBoxStyle.Information, "Information")
                            ElseIf FrmDL.Progres = FrmDownload.enumProgress.DeleteFile Then
                                MsgBox("Something went wrong during deleting buffer file, please try again.", MsgBoxStyle.Information, "Information")
                            ElseIf FrmDL.Progres = FrmDownload.enumProgress.Download Then
                                MsgBox("Something went wrong during download, please try again.", MsgBoxStyle.Information, "Information")
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAuto = Nothing
            FrmChk = Nothing
            FrmDL = Nothing
        End Try
    End Sub

    'Private Sub CheckForUpdateLAN(ByVal vShowMsg As Boolean)
    '    Dim ObjAuto As New ClsUpdateProgram
    '    Dim Obj1 As New ClsAutoUpdate
    '    Dim ObjAutoUpdate As New ClsAutoUpdate

    '    Try
    '        If System.IO.Directory.Exists(Obj1.GetServerPath) = False Then
    '            If vShowMsg = True Then
    '                MsgBox("Server path doesn't exists.", MsgBoxStyle.Information, "Information")
    '            End If
    '        Else
    '            If ObjAuto.RunProsesDownloadUpdate(ClsUpdateProgram.EnumFungsi._UpdateViaServer, vShowMsg) = True Then
    '                MsgBox("Please wait, we are going to update your software.", MsgBoxStyle.Information, "Information")
    '                Me.ObjMonitor.EndMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)

    '                Dim proc As New System.Diagnostics.Process()
    '                proc = Process.Start(CurrentDirectory & "\Client Update Module.exe")

    '                Me.CloseForm = True
    '                FrmLogin.Close()
    '                Me.Close()
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjAuto = Nothing
    '        Obj1 = Nothing
    '    End Try
    'End Sub

    Private Sub CheckForUpdateLAN(ByVal vShowMsg As Boolean)
        Dim Frm As New FrmDownload3
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim dummy As String = ""
        Dim VersionDB As String = ""
        Dim ObjVersionControl As New Core_Module.Clstrans_VersionControl
        Dim ObjVersion As New ClsVersionControl
        Dim LastVersionTxt As String = ""
        Dim DefaultVersion As String = "1.6.40"
        Dim ObjUpdate1 As New ClsUpdateProgram
        Dim OperatingVersion As String = ""

        Try
            'cek version dari db dahulu
            VersionDB = ObjVersionControl.GetVersion(My.Settings.a)
            LastVersionTxt = ObjVersion.GetVersion(DefaultVersion)
            If ObjUpdate1.ApakahVersionBerbeda(LastVersionTxt, VersionDB) = True Then
                'jika true, maka version DB > last version
                'update last version
                'ini terjadi kalo abis update
                ObjVersion.SetVersion(VersionDB)
            ElseIf ObjUpdate1.ApakahVersionBerbeda(VersionDB, LastVersionTxt) = True Then
                'jika true, maka last version > version db
                'update version db
                'ini terjadi kalo lisence baru masuk atau setelah update
                ObjVersionControl.SetVersion(My.Settings.a, LastVersionTxt)
            End If

            If Frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                MsgBox("Please wait, we are going to update your software.", MsgBoxStyle.Information, "Information")
                Me.ObjMonitor.EndMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License, Me.Identifier)

                dummy = ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.UpdateFromLAN_RequireUAC, ClsTransaksiSetting.EnumNamaParameter.UpdateFromLAN_RequireUAC)
                If dummy = "0" Or dummy = "" Then
                    Dim proc As New System.Diagnostics.Process()
                    proc = Process.Start(CurrentDirectory & "\LAN Update.exe")
                Else
                    Dim proc As New System.Diagnostics.Process()
                    proc = Process.Start(CurrentDirectory & "\Client Update Module.exe")
                End If

                Me.StartVersionCheck = True
                Me.CloseForm = True
                FrmLogin.Close()
                Me.Close()
                Exit Sub
            End If

            Me.StartVersionCheck = True
            LastVersionTxt = ObjVersion.GetVersion(DefaultVersion)
            OperatingVersion = ObjVersionControl.GetOperatingVersion
            If ObjUpdate1.ApakahVersionBerbeda(LastVersionTxt, OperatingVersion) = True Then
                If Me.ApakahBolehTimpa = False Then
                    MsgBox("Operating version is: " + OperatingVersion + ", please contact your system administrator for version update.", MsgBoxStyle.Information, "Information")
                    Me.CloseForm = True
                    FrmLogin.Close()
                    Me.Close()
                    Exit Sub
                Else
                    MsgBox("Operating version is: " + OperatingVersion + ", please contact your system administrator if this is not wanted.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm = Nothing
            ObjSetting = Nothing
            ObjVersionControl = Nothing
            ObjVersion = Nothing
            ObjUpdate1 = Nothing
        End Try
    End Sub

    Private Function ApakahBolehTimpa() As Boolean
        Try
            If FrmLogin.UsernameTextBox.Text <> "USER" Then
                Return False
            End If
            If Me.ObjMonitor.GetJmlMonitor(ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License) > 3 Then
                Return False
            End If
            If Core_Module.GetDBServer.ToLower <> "localhost" And Core_Module.GetDBServer.ToLower <> "cien-asus" Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub UploadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmUpdateFromLAN1.Click
        Me.CheckForUpdateLAN(True)
    End Sub

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
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
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
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
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

    Private Sub FrmMasterDownTime1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterDowntime2.Click
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
            i.Usercode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOC1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOC1.Click
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

    Private Sub FrmLaporanStockInOut1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanStockInOut1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanStockInOut
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanStockInOut)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanStockInOut
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

    Private Sub FrmLaporanStockTransfer1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanStockTransfer1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanStockTransfer
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanStockTransfer)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanStockTransfer
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

    Private Sub FrmPersonalSetting1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmPersonalSetting1.Click
        Try
            Dim ij As FrmPersonalSetting
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmPersonalSetting)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmPersonalSetting
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

    Private Sub FrmTransaksiPR_Otorisasi1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPR_Otorisasi1.Click
        Try
            Dim ij As New Purchase_Module.FrmTransaksiPROtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTransaksiPROtorisasi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next
            Dim i As New Purchase_Module.FrmTransaksiPROtorisasi
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BAckColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanLogUser1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanLogUser1.Click
        Try
            Dim ij As HR.FrmLaporanLogUser
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanLogUser)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanLogUser
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

    Private Sub FrmLaporanPurchaseReturnDetail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseReturnDetail1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseReturnDetail
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseReturnDetail)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseReturnDetail
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

    Private Sub FrmLaporanSalesReturnDetail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesReturnDetail1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesReturnDetail
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesReturnDetail)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesReturnDetail
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

    Private Sub FrmLaporanCommission1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanCommission1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanCommission
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanCommission)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanCommission
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

    Private Sub FrmTrans_Gaji_OverrideA1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Gaji_OverrideA1.Click
        Try
            Dim ij As HR.FrmTrans_Gaji_Override
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Gaji_Override)
                If IsNothing(ij) = False Then
                    If ij.Tipe = HR.ClsTransaksiSalary_Override.EnumType.OtherPayment Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New HR.FrmTrans_Gaji_Override
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.Tipe = HR.ClsTransaksiSalary_Override.EnumType.OtherPayment
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_TaxRecalculate1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_TaxRecalculate1.Click
        Try
            Dim ij As HR.FrmTrans_TaxRecalculate
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_TaxRecalculate)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_TaxRecalculate
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

    Private Sub FrmMasterCommission_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCommission1.Click
        Try
            Dim ij As FrmMasterCommission
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterCommission)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterCommission
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

    Private Sub FrmLaporanQuotationDetail1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanQuotationDetail1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanQuotationDetail
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanQuotationDetail)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanQuotationDetail
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

    Private Sub FrmLaporanQuotationRealisasi1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanQuotationRealisasi1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanQuotationRealisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanQuotationRealisasi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanQuotationRealisasi
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

    Private Sub FrmLaporanQuotationRealisasiChart1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanQuotationRealisasiChart1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanQuotationRealisasiChart
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanQuotationRealisasiChart)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanQuotationRealisasiChart
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

    Private Sub FrmLaporanQuotationTrendChart1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanQuotationTrendChart1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanQuotationTrendChart
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanQuotationTrendChart)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanQuotationTrendChart
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

    Private Sub FrmTrans_CutOffDB1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_CutOffDB1.Click
        Try
            Dim ij As Purchase_Module.FrmTrans_CutOffDB
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmTrans_CutOffDB)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmTrans_CutOffDB
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

    Private Sub FrmMasterProductionRate1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProductionRate1.Click
        Try
            Dim ij As Core_Module.FrmMasterProductionRate
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterProductionRate)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterProductionRate
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

    Private Sub FrmTrans_ManufacturingOrder1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_ManufacturingOrder1.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_ManufacturingOrder
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_ManufacturingOrder)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_ManufacturingOrder
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

    Private Sub FrmTrans_ProductionSchedule1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_ProductionSchedule1.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_ProductionSchedule
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_ProductionSchedule)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_ProductionSchedule
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

    Private Sub FrmTransaksiStockBundling21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockBundling21.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_StockBundling
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_StockBundling)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_StockBundling
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

    Private Sub FrmTrans_ManufacturingOrder_SetFinishJO1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_ManufacturingOrder_SetFinishJO1.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_ManufacturingOrder_SetFinishJO
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_ManufacturingOrder_SetFinishJO)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_ManufacturingOrder_SetFinishJO
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

    Private Sub FrmTrans_ManufacturingOrder_Authorize1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_ManufacturingOrder_Authorize1.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_ManufacturingOrder_Authorize
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_ManufacturingOrder_Authorize)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_ManufacturingOrder_Authorize
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

    Private Sub FrmMasterWorkUnitGroup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterWorkUnitGroup1.Click
        Try
            Dim ij As Core_Module.FrmMasterWorkUnitGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterWorkUnitGroup)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterWorkUnitGroup
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

    Private Sub FrmMasterSetupTime1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSetupTime1.Click
        Try
            Dim ij As Core_Module.FrmMasterSetupTime
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterSetupTime)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterSetupTime
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

    Private Sub FrmLaporanProd_Productivity1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProd_Productivity1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_Productivity
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_Productivity)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_Productivity
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

    Private Sub FrmLaporanProd_RejectReport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProd_RejectReport1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_RejectReport
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_RejectReport)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_RejectReport
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

    Private Sub FrmLaporanProd_Downtime1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProd_Downtime1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_Downtime
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_Downtime)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_Downtime
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

    Private Sub FrmLaporanProd_RawMaterialUsage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProd_RawMaterialUsage1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_RawMaterialUsage
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_RawMaterialUsage)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_RawMaterialUsage
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

    Private Sub FrmLaporanProd_PlanningReport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProd_PlanningReport1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_PlanningReport
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_PlanningReport)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_PlanningReport
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

    Private Sub FrmMasterMachineLine1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterMachineLine1.Click
        Try
            Dim ij As FrmMasterMachineLine
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterMachineLine)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterMachineLine
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

    Private Sub CekAutoBackUpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("Next Back Up " + Format(Me.ObjAutoDB.RefTimeScheduler + TimeSpan.FromMilliseconds(Me.ObjAutoDB.BackUpScheduler.Interval), "dd MMM yyy HH.mm"), MsgBoxStyle.Information, "Information")
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Authorize11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_MaterialPlanningDashboard_MP_Authorize11.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize)
                If IsNothing(ij) = False Then
                    If ij.AuthLevel = 1 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.AuthLevel = 1
            i.Text = "Material Planning Authorize Level 1"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Authorize21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_MaterialPlanningDashboard_MP_Authorize21.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize)
                If IsNothing(ij) = False Then
                    If ij.AuthLevel = 2 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.AuthLevel = 2
            i.Text = "Material Planning Authorize Level 2"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Authorize31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_MaterialPlanningDashboard_MP_Authorize31.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize)
                If IsNothing(ij) = False Then
                    If ij.AuthLevel = 3 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_Authorize
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.AuthLevel = 3
            i.Text = "Material Planning Authorize Level 3"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_BOM11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmMasterProduct_BOM11.Click
        Try
            Dim ij As frmMasterProduct_BOM1
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, frmMasterProduct_BOM1)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New frmMasterProduct_BOM1
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Text = "Bill Of Material Tools"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKaryawan_TipeKontrak11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKaryawan_TipeKontrak1.Click
        Try
            Dim ij As Core_Module.FrmMasterKaryawan_TipeKontrak
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterKaryawan_TipeKontrak)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterKaryawan_TipeKontrak
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            'i.Text = "Bill Of Material Tools"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanKontrakKaryawan11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanKontrakKaryawan1.Click
        Try
            Dim ij As HR.FrmLaporanKontrakKaryawan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanKontrakKaryawan)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanKontrakKaryawan
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            'i.Text = "Bill Of Material Tools"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSupplierList1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FrmLaporanSupplierList1.Click
        Try
            Dim ij As FrmLaporanSupplierList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmLaporanSupplierList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmLaporanSupplierList
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            'i.Text = "Bill Of Material Tools"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanCustomerList1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FrmLaporanCustomerList1.Click
        Try
            Dim ij As FrmLaporanCustomerList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmLaporanCustomerList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmLaporanCustomerList
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Text = "Bill Of Material Tools"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanProd_RMAllocation1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProd_RMAllocation1.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_PickList
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_PickList)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_PickList
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

    Private Sub UpdateAllAccessToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateAllAccessToolStripMenuItem.Click
        Dim ObjCustomReport As New CustomSMESS.ClsTrans_CustomReport(FrmLogin.UsernameTextBox.Text)
        Dim ObjCustomTrans As New CustomSMESS.ClsTrans_CustomTrans(FrmLogin.UsernameTextBox.Text)

        Try
            If FrmLogin.UsernameTextBox.Text.ToLower = "user" Then
                Me.UpdateHakAkses(False)
                Me.UpdateHakAkses_HR(False)
                Me.UpdateTransNumbering()
                ObjCustomReport.UpdateMenu(True)
                ObjCustomTrans.UpdateMenu()

                MsgBox("Data updated", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCustomReport = Nothing
            ObjCustomTrans = Nothing
        End Try
    End Sub

    Private Sub UpdateTransNumbering()
        Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
        Dim TipeNumbering As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.AOP_Impor

        Try
            'untuk sementara SPL + DT ditutup dulu, karena SPL terkait produksi belum dikeluarkan pada form utama, sementara masih menggunakan penjadwal overtime yang lebih sederhana
            'If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL) = False Then
            '    ObjNumbering.SimpanBaru("Inventory - Overtime Order", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SPL, "tbltrans_spl", FrmLogin.UsernameTextBox.Text)
            'End If
            'If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime) = False Then
            '    ObjNumbering.SimpanBaru("Inventory - Downtime Schedule", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Schedule_Downtime, "tbltrans_sch_downtime", FrmLogin.UsernameTextBox.Text)
            'End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = False Then
                ObjNumbering.SimpanBaru("Sales - Delivery Order", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO, "tbltrans_do", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = False Then
                ObjNumbering.SimpanBaru("Inventory - Material Planning", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP, "tbltrans_material_planning", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note) = False Then
                ObjNumbering.SimpanBaru("Accounting - Credit Note", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note, "tbltrans_credit_note", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note) = False Then
                ObjNumbering.SimpanBaru("Accounting - Debit Note", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note, "tbltrans_credit_note", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = False Then
                ObjNumbering.SimpanBaru("Purchase - Purchase Order - Local", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal, "tbltrans_po", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) = False Then
                ObjNumbering.SimpanBaru("Purchase - Purchase Order - Import", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor, "tbltrans_po", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter) = False Then
                ObjNumbering.SimpanBaru("Accounting - Collection Letter", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.CollectionLetter, "tbltrans_surattagihan", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = False Then
                ObjNumbering.SimpanBaru("Sales - Invoice", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, "tbltrans_salesinvoice", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF) = False Then
                ObjNumbering.SimpanBaru("Inventory - Shipping Forecast", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._SF, "tbltrans_shipping_forecast", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO) = False Then
                ObjNumbering.SimpanBaru("Inventory - Manufacturing Order", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO, "tbltrans_mo", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling) = False Then
                ObjNumbering.SimpanBaru("Inventory - Bundling", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Bundling, "tbltransaksi_stockbundling", FrmLogin.UsernameTextBox.Text)
            End If
            If ObjNumbering.ApakahSudahAda(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = False Then
                ObjNumbering.SimpanBaru("Sales - Sales Order", ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO, "tbltrans_so", FrmLogin.UsernameTextBox.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanOtherPayment1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanOtherPayment1.Click
        Try
            Dim ij As HR.FrmLaporanOtherPayment
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanOtherPayment)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanOtherPayment
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

    Private Sub FrmLaporanPinjamanKaryawan1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPinjamanKaryawan1.Click
        Try
            Dim ij As HR.FrmLaporanPinjamanKaryawan
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanPinjamanKaryawan)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanPinjamanKaryawan
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

    Private Sub FrmTrans_SPL_Budget1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_SPL_Budget1.Click
        Try
            Dim ij As HR.FrmTrans_SPL_Budget
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_SPL_Budget)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_SPL_Budget
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

    Private Sub FrmTrans_SPL1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_SPL1.Click
        Try
            Dim ij As HR.FrmTrans_SPL
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_SPL)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_SPL
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

    Private Sub FrmTrans_SPLApproval1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_SPLApproval1.Click
        Try
            Dim ij As HR.FrmTrans_SPLApproval
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_SPLApproval)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_SPLApproval
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

    Private Sub FrmLaporanOvertimeOrder1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanOvertimeOrder1.Click
        Try
            Dim ij As HR.FrmLaporanOvertimeOrder
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmLaporanOvertimeOrder)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmLaporanOvertimeOrder
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

    Private Sub FrmTrans_SPL_Setting1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_SPL_Setting1.Click
        Try
            Dim ij As HR.FrmTrans_SPL_Setting
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_SPL_Setting)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_SPL_Setting
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

    Private Sub EmployeeGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMaster_SPL_EmployeeCustom1.Click
        Try
            Dim ij As Core_Module.FrmMaster_SPL_EmployeeCustom
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMaster_SPL_EmployeeCustom)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMaster_SPL_EmployeeCustom
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

    Private Sub FrmTrans_StockBundling_Quick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_StockBundling_Quick1.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_StockBundling_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_StockBundling_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_StockBundling_Quick
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

    Private Sub MonthlyClosingTools1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MonthlyClosingTools1.Click
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

    Private Sub FrmTransaksiBalanceManual_DPSls1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiBalanceManual_DPSls1.Click
        Try
            Dim ij As AccModule.FrmTransaksiBalanceManual_DP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiBalanceManual_DP)
                If IsNothing(ij) = False Then
                    If ij.Text = "Historical Down Payment Sales Balance" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New AccModule.FrmTransaksiBalanceManual_DP
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.vTipeDP = AccModule.ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_Invoice
            ' Set the caption to be unique.
            i.Text = "Historical Down Payment Sales Balance"
            'back color
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiBalanceManual_DPPch1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiBalanceManual_DPPch1.Click
        Try
            Dim ij As AccModule.FrmTransaksiBalanceManual_DP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiBalanceManual_DP)
                If IsNothing(ij) = False Then
                    If ij.Text = "Historical Down Payment Purchase Balance" Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            Dim i As New AccModule.FrmTransaksiBalanceManual_DP
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.vTipeDP = AccModule.ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP
            ' Set the caption to be unique.
            i.Text = "Historical Down Payment Purchase Balance"
            'back color
            i.BackCOlorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateUtilityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateUtilityToolStripMenuItem.Click
        Dim ij As New FrmUpdate

        Try
            ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockTransfer_Receive1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockTransfer_Receive1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockTransfer_Receive
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockTransfer_Receive)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockTransfer_Receive
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

    Private Sub FrmTransaksiExpiryDateUpdate_Merge1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiExpiryDateUpdate_Merge1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiExpiryDateUpdate_Merge
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiExpiryDateUpdate_Merge)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiExpiryDateUpdate_Merge
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

    Private Sub frmMasterProduct_ChangeControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmMasterProduct_ChangeControl1.Click
        Try
            Dim ij As Core_Module.frmMasterProduct_ChangeControl
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.frmMasterProduct_ChangeControl)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.frmMasterProduct_ChangeControl
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

    Private Sub FrmTrans_Gaji_MasaSblm1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Gaji_MasaSblm1.Click
        Try
            Dim ij As HR.FrmTrans_Gaji_MasaSblm
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmTrans_Gaji_MasaSblm)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmTrans_Gaji_MasaSblm
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

    Private Sub FrmMasterCustomer_Reference1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomer_Reference1.Click
        Try
            Dim ij As Core_Module.FrmMasterCustomer_Reference
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterCustomer_Reference)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterCustomer_Reference
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

    Private Sub FrmMasterKaryawan_Kompetensi_Upd1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKaryawan_Kompetensi_Upd1.Click
        Try
            Dim ij As Core_Module.FrmMasterKaryawan_Kompetensi_Upd
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterKaryawan_Kompetensi_Upd)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterKaryawan_Kompetensi_Upd
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

    Private Sub FrmTransaksiStockInAdj_SetAmount1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockInAdj_SetAmount1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInAdj_SetAmount
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockInAdj_SetAmount)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockInAdj_SetAmount
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.MdiParent = Me
            i.ShowIcon = False 

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseDP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanPurchaseDP1.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanPurchaseDP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanPurchaseDP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanPurchaseDP
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesDP1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FrmLaporanSalesDP1.Click
        Try
            Dim ij As Sales_Module.FrmLaporanSalesDP
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmLaporanSalesDP)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmLaporanSalesDP
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSOC1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSOC1.Click
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

    Private Sub FrmTransaksiSOC_QuickDel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSOC_QuickDel1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSOC_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiSOC_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiSOC_Quick
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

    Private Sub FrmLaporanVAT2_Custom1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanVAT2_Custom1.Click
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

    Private Sub FrmImport_eFaktur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmImport_eFaktur1.Click
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

    Private Sub FrmImport_Art_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmImport_Art1.Click
        Try
            Dim ij As Sales_Module.FrmImport_Art
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmImport_Art)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmImport_Art
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

    Private Sub LicenseToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LicenseToolStripMenuItem1.Click
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

    Private Sub ReqLicenseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReqLicenseToolStripMenuItem.Click
        Try
            Dim ij As Core_Module.FrmLicence_Add
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmLicence_Add)
                If IsNothing(ij) = False Then
                    If ij.IsBlackList = 0 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Core_Module.FrmLicence_Add
            ' Set the caption to be unique.
            i.IsBlackList = 0
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputLicenseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputLicenseToolStripMenuItem.Click
        Try
            Dim ij As Core_Module.FrmLicence_Input
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmLicence_Input)
                If IsNothing(ij) = False Then
                    If ij.IsBlackList = 0 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Core_Module.FrmLicence_Input
            ' Set the caption to be unique.
            i.IsBlackList = 0
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReqLicenseBlacklistToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReqLicenseBlacklistToolStripMenuItem.Click
        Try
            Dim ij As Core_Module.FrmLicence_Add
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmLicence_Add)
                If IsNothing(ij) = False Then
                    If ij.IsBlackList = 1 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Core_Module.FrmLicence_Add
            ' Set the caption to be unique.
            i.IsBlackList = 1
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputLicenseBlacklistToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputLicenseBlacklistToolStripMenuItem.Click
        Try
            Dim ij As Core_Module.FrmLicence_Input
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmLicence_Input)
                If IsNothing(ij) = False Then
                    If ij.IsBlackList = 1 Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                End If
            Next

            ' Increment the caption counter.
            Dim i As New Core_Module.FrmLicence_Input
            ' Set the caption to be unique.
            i.IsBlackList = 1
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLicensing21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiLicensing21.Click
        Try
            Dim ij As FrmTransaksiLicensing2
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmTransaksiLicensing2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            ' Increment the caption counter.
            Dim i As New FrmTransaksiLicensing2
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.ShowIcon = False
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TaxProductReferenceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterTaxProductReference1.Click
        Try
            Dim ij As Core_Module.FrmMasterTaxProductReference
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterTaxProductReference)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterTaxProductReference
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

    Private Sub TaxMeasurementReferenceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterTaxMeasurementReference1.Click
        Try
            Dim ij As Core_Module.FrmMasterTaxMeasurementReference
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterTaxMeasurementReference)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterTaxMeasurementReference
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

    Private Sub VATReportCoreTax2025ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanVAT20251.Click
        Try
            Dim ij As Purchase_Module.FrmLaporanVAT2025
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Purchase_Module.FrmLaporanVAT2025)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Purchase_Module.FrmLaporanVAT2025
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

    Private Sub FrmMasterTaxAddInfoReference1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterTaxAddInfoReference1.Click
        Try
            Dim ij As Core_Module.FrmMasterTaxAddInfoReference
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterTaxAddInfoReference)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterTaxAddInfoReference
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

    Private Sub FrmMasterTaxFacilityStampReference1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterTaxFacilityStampReference1.Click
        Try
            Dim ij As Core_Module.FrmMasterTaxFacilityStampReference
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterTaxFacilityStampReference)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterTaxFacilityStampReference
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

    Private Sub FrmTransaksiDO_Quick1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiDO_Quick1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiDO_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Sales_Module.FrmTransaksiDO_Quick)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Sales_Module.FrmTransaksiDO_Quick
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
