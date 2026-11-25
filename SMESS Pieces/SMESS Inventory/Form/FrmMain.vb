Option Explicit On
Option Strict Off

Imports Core_Module
Imports System.IO

Public Class FrmMain
    Public EW As E_mail_Module.ClsEmailWatcher
    Dim ObjComapanyData As New Core_Module.ClsMasterCompanyData
    Dim CloseForm As Boolean = False
    Dim ObjLicense As New ClsLicense
    Dim ObjHakAkses As New ClsHakAkses

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
        Dim i As ToolStripMenuItem
        Dim TBL As New DSUser.TblHakAksesUserDataTable
        Dim DTA As New DSUserTableAdapters.TblHakAksesUserTableAdapter
        Dim int As Integer
        Dim ObjLS As New clsMasterListedSoftware
        Dim CompanyName As String = ""
        Dim PerbaruiLicense As Boolean = False
        Dim DateDiff As Integer = 0
        Dim ObjDB As New ClsDBVersion
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier

        Try
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

            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard_Data1") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard_Data1", "Hak akses untuk membuka form pengelompokan perencanaan material")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTrans_MaterialPlanningDashboard") = False Then
                Me.ObjHakAkses.Simpan("FrmTrans_MaterialPlanningDashboard", "Hak akses untuk membuka form perencanaan material")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterProses") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterProses", "Hak akses untuk input, edit, dan delete master sales proses")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterBackUpDB") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterBackUpDB", "Hak akses untuk input, edit, dan delete rencana back up database")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("frmTransaksiStockTaking_Print") = False Then
                Me.ObjHakAkses.Simpan("frmTransaksiStockTaking_Print", "Hak akses untuk cetak form kosong stock taking")
            End If

            If Me.ObjHakAkses.ApakahSudahAda("FrmManajemenHakAkses_Override_Edit") = False Then
                Me.ObjHakAkses.Simpan("FrmManajemenHakAkses_Override_Edit", "Hak akses untuk mengedit data transaksi hak akses yang dibuat oleh user lain", "FrmManajemenHakAkses", 0)
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmManajemenHakAkses_Override_Delete") = False Then
                Me.ObjHakAkses.Simpan("FrmManajemenHakAkses_Override_Delete", "Hak akses untuk menghapus data transaksi hak akses yang dibuat oleh user lain", "FrmManajemenHakAkses", 0)
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
            'Dim nt As New Threading.Thread(AddressOf DoFirstActivitySales)
            'nt.Start()

            'fil tabel
            DTA.Fill(TBL, FrmLogin.UsernameTextBox.Text)

            For int = 0 To Me.MenuStrip1.Items.Count - 1
                i = TryCast(Me.MenuStrip1.Items(int), ToolStripMenuItem)
                If IsNothing(i) = False Then
                    Me.CekTollStripMenu(i, TBL)
                End If
            Next

            Windows.Forms.Cursor.Current = Cursors.Arrow

            Me.FileToolStripMenuItem.Visible = True
            Me.KeluarToolStripMenuItem.Visible = True

            'activates email watcher, untuk sementara email watcher dimatikan dulu, karen mengganggu proses debugging
            EW = New E_mail_Module.ClsEmailWatcher(FrmLogin.UsernameTextBox.Text, Me, Me.NotifyIcon1, _
            Me.AlertMeOnIncomingEmailToolStripMenuItem)
            'EW.Enabled = True

            Me.WindowsToolStripMenuItem.Visible = True
            ProductionTransactionToolStripMenuItem.Visible = True
            Me.TileHorizontallyToolStripMenuItem.Visible = True
            Me.TileVericallyToolStripMenuItem.Visible = True
            Me.CascadeToolStripMenuItem.Visible = True
            Me.CloseAllWindowsToolStripMenuItem.Visible = True
            Me.InternalMessageToolStripMenuItem.Visible = True
            Me.AlertMeOnIncomingEmailToolStripMenuItem.Visible = True
            Me.Test1.Visible = True
            Me.FrmTrans_ManufacturingOrder1.Visible = True
            Me.FrmTrans_ProductionSchedule1.Visible = True
            Me.FrmTransaksiStockBundling21.Visible = True
            Me.FrmTrans_ManufacturingOrder_SetFinishJO1.Visible = True
            Me.FrmTrans_ManufacturingOrder_Authorize1.Visible = True
            Me.FrmMasterWorkShift1.Visible = True
            Me.FrmMasterWorkUnitGroup1.Visible = True
            Me.FrmMasterSetupTime1.Visible = True
            Me.ProductionReportToolStripMenuItem.Visible = True
            Me.FrmLaporanProd_RMAllocation1.Visible = True
            Me.FrmLaporanProd_Productivity1.Visible = True
            Me.FrmLaporanProd_RejectReport1.Visible = True
            Me.FrmLaporanProd_RawMaterialUsage1.Visible = True
            Me.FrmLaporanProd_Downtime1.Visible = True
            Me.FrmLaporanProd_PlanningReport1.Visible = True
            Me.DefaultSettingToolStripMenuItem.Visible = True
            Me.FrmTrans_SPL1.Visible = True
            If ObjIdentifier.GetIdentifier = "smess_lmb" Then
                UpdateM3LMBToolStripMenuItem.Visible = True
            End If
            'jika tidak ada stock taking aktif, maka stock taking process tidak perlu ditampilkan
            Dim ObjST As New Inventory_Module.clsTransaksiStockTaking
            If ObjST.CekApakahAdaYangMasihAktif(False) = False Then
                FrmTransaksiStockTakingProcess1.Visible = False
                frmTransaksiStockTaking_Print1.Visible = False
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
            My.Settings.a = Me.ObjLicense.GetProductKey
            My.Settings.b = Me.ObjLicense.GetLicenseKey(Today, My.Settings.a, ClsLicense.EnumTipeLicense.Normal1)

            'update license bila expiry
            If Today > Me.ObjLicense.GetExpiryDate(My.Settings.b) Then
                My.Settings.b = Me.ObjLicense.GetLicenseKey(Me.ObjLicense.GetExpiryDate(My.Settings.b) + TimeSpan.FromDays(1), _
                                      My.Settings.a, Me.ObjLicense.GetTipeLicense(My.Settings.b) + 1)
            End If


            Core_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Inventory_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            E_mail_Module.ModGlobalVariable.SetLicense(My.Settings.b)

            'Dim ijklm As New Purchase_Module.FrmCekMovement
            'ijklm.ShowDialog()
            'Me.ManajemenUserToolStripMenuItem.Visible = True
            'Me.FrmManajemenHakAkses1.Visible = True
            'Me.FrmLaporanPembelian1.Visible = True
            'Me.FrmLaporanPenjualan1.Visible = True
            'My.Settings.a = "58DCC - S9MEX -   098 - XPMVQ - UACVP"
            'My.Settings.b = "6cT0O40ru2Sc9pq1pf1n90so0"
            Me.FileToolStripMenuItem.Visible = True
            Me.ManajemenUserToolStripMenuItem.Visible = True
            Me.FrmManajemenHakAkses1.Visible = True

            My.Settings.Save()

            'start auto back up
            'Dim ObjAutoDB As New Core_Module.ClsMasterBackUpDB
            'ObjAutoDB.StartBackUpScheduler(FrmLogin.UsernameTextBox.Text)

            Dim ij As New Inventory_Module.Form2
            'ij.UserCode = "USER"
            'ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            'ij.ObjStockBundling.TipeBundling = Inventory_Module.ClsTransaksiStockBundling.EnumTipeBundling.Purchase
            ij.MdiParent = Me
            ij.Show()

            UserLocationAccessToolStripMenuItem.Visible = True
            ProductionLogInputByJOToolStripMenuItem.Visible = True

            ToolsToolStripMenuItem.Visible = True
            FrmTrans_ProductionSchedule_Reorder_LMB1.Visible = True
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

    Private Function CekTollStripMenu(ByVal i As ToolStripMenuItem, ByVal TBL As DSUser.tblhakaksesUserDataTable) As Boolean
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
            i.MdiParent = Me
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
            i.BackColorAll = Color.FromArgb(170, 255, 190)
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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

            ' Increment the caption counter.
            Dim i As New Inventory_Module.FrmTransaksiStockInAdj
            ' Set the caption to be unique.
            i.MdiParent = Me
            i.BeginingBalance = True
            'i.BackColorAll = Color.FromArgb(170, 255, 190)
            i.UserCode = FrmLogin.UsernameTextBox.Text
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
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackUpDBToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterBackUpDB1.Click
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
            i.Show()
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

    Private Sub FrmMasterPriorityLevel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterPriorityLevel1.Click
        Try
            Dim ij As Core_Module.FrmMasterPriorityLevel
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterPriorityLevel)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterPriorityLevel
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

    Private Sub MaterialPlanningGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_MaterialPlanningDashboard_Data11.Click
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

            Dim i As New Core_Module.FrmLaporanOtorisasiTransaksi
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

    Private Sub FrmLaporanStockBundling1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanStockBundling1.Click
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

    Private Sub FrmMasterWorkShift1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterWorkShift1.Click
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

    Private Sub FrmLaporanProd_RMAllocation1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLaporanProd_RMAllocation1.Click
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

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterMachineLine1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterMachineLine1.Click
        Try
            Dim ij As Core_Module.FrmMasterMachineLine
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmMasterMachineLine)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmMasterMachineLine
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

    Private Sub DefaultSettingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DefaultSettingToolStripMenuItem.Click
        Try
            Dim ij As Inventory_Module.FrmSettingProduksi
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmSettingProduksi)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmSettingProduksi
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

    Private Sub FrmTrans_MaterialPlanningDashboard_Authorize31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_MaterialPlanningDashboard_MP_Authorize31__.Click
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
                    'If ij.AuthLevel = 3 Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                    'End If
                End If
            Next

            Dim i As New frmMasterProduct_BOM1
            ' Increment the caption counter.
            i.UserCode = FrmLogin.UsernameTextBox.Text
            ' Set the caption to be unique.
            i.BackColorAll = Core_Module.GetSMESSBackColorAll
            'i.AuthLevel = 3
            'i.Text = "Material Planning Authorize Level 3"
            i.MdiParent = Me
            i.ShowIcon = False

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UserLocationAccessToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserLocationAccessToolStripMenuItem.Click
        Try
            Dim ij As Core_Module.FrmUserLocationPrivilegeSetting
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Core_Module.FrmUserLocationPrivilegeSetting)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Core_Module.FrmUserLocationPrivilegeSetting
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

     
    Private Sub OvertimeOrderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_SPL1.Click
        Try
            Dim ij As Inventory_Module.FrmTrans_SPL
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTrans_SPL)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTrans_SPL
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

    Private Sub ProductionLogInputByJOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductionLogInputByJOToolStripMenuItem.Click
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

    Private Sub FrmTrans_ProductionSchedule_Reorder_LMB1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_ProductionSchedule_Reorder_LMB1.Click
        Dim ij As New Inventory_Module.FrmTrans_ProductionSchedule_Reorder_LMB

        Try
            ij.ShowDialog()
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
End Class