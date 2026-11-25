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
        'Dim i As ToolStripMenuItem
        Dim TBL As New DSUser.TblHakAksesUserDataTable
        Dim DTA As New DSUserTableAdapters.TblHakAksesUserTableAdapter
        'Dim int As Integer
        Dim ObjLS As New clsMasterListedSoftware
        Dim CompanyName As String = ""
        Dim PerbaruiLicense As Boolean = False
        Dim DateDiff As Integer = 0
        Dim ObjDB As New ClsDBVersion

        Try
            If Me.ObjHakAkses.ApakahSudahAda("FrmLaporanSalesPerformance") = False Then
                Me.ObjHakAkses.Simpan("FrmLaporanSalesPerformance", "Hak akses untuk melihat laporan performa tim sales")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesBudget") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesBudget", "Hak akses untuk melihat transaksi budget penjualan")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSalesPersonGroup") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSalesPersonGroup", "Hak akses untuk melihat master sales person group")
            End If


            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiSalesInvoiceSettingPajak") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiSalesInvoiceSettingPajak", "Hak akses untuk melakukan setting faktuk pajak")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterSalesPerson") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterSalesPerson", "Hak akses untuk input, edit, dan delete master sales person")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_SalesDashboard") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_SalesDashboard", "Hak akses untuk mengakses sales dashboard")
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

            'fil tabel
            DTA.Fill(TBL, FrmLogin.UsernameTextBox.Text)

            'For int = 0 To Me.MenuStrip1.Items.Count - 1
            'i = TryCast(Me.MenuStrip1.Items(int), ToolStripMenuItem)
            'If IsNothing(i) = False Then
            'Me.CekTollStripMenu(i, TBL)
            'End If
            'Next

            Windows.Forms.Cursor.Current = Cursors.Arrow

            Me.FileToolStripMenuItem.Visible = True
            Me.KeluarToolStripMenuItem.Visible = True

            'activates email watcher, untuk sementara email watcher dimatikan dulu, karen mengganggu proses debugging
            EW = New E_mail_Module.ClsEmailWatcher(FrmLogin.UsernameTextBox.Text, Me, Me.NotifyIcon1, _
            Me.AlertMeOnIncomingEmailToolStripMenuItem)
            'EW.Enabled = True

            Me.WindowsToolStripMenuItem.Visible = True
            Me.TileHorizontallyToolStripMenuItem.Visible = True
            Me.TileVericallyToolStripMenuItem.Visible = True
            Me.CascadeToolStripMenuItem.Visible = True
            Me.CloseAllWindowsToolStripMenuItem.Visible = True
            Me.InternalMessageToolStripMenuItem.Visible = True
            Me.AlertMeOnIncomingEmailToolStripMenuItem.Visible = True
            Me.Test1.Visible = True

            'jika tidak ada stock taking aktif, maka stock taking process tidak perlu ditampilkan
            Dim ObjST As New Inventory_Module.clsTransaksiStockTaking
            If ObjST.CekApakahAdaYangMasihAktif(False) = False Then
                FrmTransaksiStockTakingProcess1.Visible = False
            End If
            ObjST = Nothing

            Dim ijk As New FrmMainSMESS
            ijk.MdiParent = Me
            ijk.Show()

            CompanyName = Me.ObjComapanyData.GetCompanyName
            If CompanyName <> "" Then
                CompanyName = " - " + CompanyName
            End If
            Me.Text = "SMESS - Ver. " + CStr(My.Application.Info.Version.Major) + "." + _
            CStr(My.Application.Info.Version.Minor) + "." + CStr(My.Application.Info.Version.Build) + _
            CompanyName

            Me.LicenseToolStripMenuItem.Visible = True

            'coding berkaitan dengan licensing
            'If Today > Me.ObjLicense.GetExpiryDate(My.Settings.b) Then
            My.Settings.a = Me.ObjLicense.GetProductKey
            My.Settings.b = Me.ObjLicense.GetLicenseKey(Today, My.Settings.a, ClsLicense.EnumTipeLicense.Normal1)
            'End If

            Core_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Inventory_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Sales_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            E_mail_Module.ModGlobalVariable.SetLicense(My.Settings.b)

            My.Settings.Save()
            ImportEfakturToolStripMenuItem.Visible = True
            ImportArtToolStripMenuItem.Visible = True
            'Me.Show17()
            'Dim ij As New CustomSMESS.frmlaporan
            'ij.UserCode = "USER"
            'ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            'ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTA.Dispose()
            obj = Nothing
            ObjLS = Nothing
            ObjDB = Nothing
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
            'i.BackColorAll = Color.MistyRose
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
            'i.BackColorAll = Color.MistyRose
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
            ' i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
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
            i.BackColorAll = Color.MistyRose
            i.MdiParent = Me
            i.Show()
            i = Nothing
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
            'i.BackColorAll = Color.MistyRose
            i.MdiParent = Me
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub StockCardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockCard1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockCard
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockCard)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockCard
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.UserCode = FrmLogin.UsernameTextBox.Text
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.BackcolorAll = Core_Module.GetSMESSBackColorAll
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
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

            i.BackColorAll = Core_Module.GetSMESSBackColorAll

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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            'i.BackColorAll = Core_Module.GetSMESSBackColorAll 
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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

    Private Sub FrmMain_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Try
            Dim gr As Graphics = e.Graphics

            gr.DrawString("Master Data", Me.Font, Brushes.White, 100, 100)
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
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalesPerson1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesPerson11.Click

    End Sub

    Private Sub QuickSalesReturnToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesReturn_Quick1.Click
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
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmMasterProduct_PriceList1.Click
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
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanOtorisasiTransaksi1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanOtorisasiTransaksi1.Click
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

    Private Sub ImportEfakturToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportEfakturToolStripMenuItem.Click
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

    Private Sub ImportArtToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportArtToolStripMenuItem.Click
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
End Class