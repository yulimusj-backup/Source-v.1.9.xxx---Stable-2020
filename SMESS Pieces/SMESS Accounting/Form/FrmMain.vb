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
            'Me.EW.Enabled = False
            'Me.EW.Dispose()
            'Me.EW = Nothing
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
        Dim ObjReminder As New AccModule.ClsTransaksiReminder
        Dim ObjDB As New ClsDBVersion

        Try
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterAssetGroup") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterAssetGroup", "Hak akses untuk input, edit, dan hapus data master group asset")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterAsset") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterAsset", "Hak akses untuk input, edit, dan hapus data master asset")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmLapAsset_DepresiasiPerPeriod") = False Then
                Me.ObjHakAkses.Simpan("FrmLapAsset_DepresiasiPerPeriod", "Hak akses untuk melihat laporan depresiasi aset")
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
            'EW = New E_mail_Module.ClsEmailWatcher(FrmLogin.UsernameTextBox.Text, Me, Me.NotifyIcon1, _
            'Me.AlertMeOnIncomingEmailToolStripMenuItem)
            'EW.Enabled = True

            Me.WindowsToolStripMenuItem.Visible = True
            Me.TileHorizontallyToolStripMenuItem.Visible = True
            Me.TileVericallyToolStripMenuItem.Visible = True
            Me.CascadeToolStripMenuItem.Visible = True
            Me.CloseAllWindowsToolStripMenuItem.Visible = True
            Me.InternalMessageToolStripMenuItem.Visible = True
            Me.AlertMeOnIncomingEmailToolStripMenuItem.Visible = True
            Me.Test1.Visible = True
            Me.OvertimeOrderToolStripMenuItem.Visible = True
            Me.FrmTrans_SPL1.Visible = True
            Me.FrmTrans_SPLApproval1.Visible = True
            Me.FrmTrans_SPL_Budget1.Visible = True

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

            Core_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            AccModule.ModGlobalVariable.SetLicense(My.Settings.b)
            E_mail_Module.ModGlobalVariable.SetLicense(My.Settings.b)
            Me.FrmMasterSettingPajak1.Visible = True
            FrmTrans_TaxRecalculate1.Visible = True

            'reminder
            If ObjReminder.GetJml(Today, FrmLogin.UsernameTextBox.Text) > 0 Then
                MsgBox("You have reminder for today", MsgBoxStyle.Information, "Information")
                Me.FrmTransaksiReminder1.PerformClick()
            End If

            My.Settings.Save()

            If Me.ObjHakAkses.ApakahSudahAda("FrmMasterDefaultValues") = False Then
                Me.ObjHakAkses.Simpan("FrmMasterDefaultValues", "Hak akses untuk melakukan setting default nilai tukar mata uang asing")
            End If
            If Me.ObjHakAkses.ApakahSudahAda("FrmTransaksiJournal_CashBank") = False Then
                Me.ObjHakAkses.Simpan("FrmTransaksiJournal_CashBank", "Hak akses untuk input, edit, dan delete jurnal kas dan bank")
            End If

            Me.FrmTrans_TaxRecalculate1.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTA.Dispose()
            obj = Nothing
            ObjLS = Nothing
            ObjReminder = Nothing
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

    Private Sub FrmTransaksiGeneralJournal1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiGeneralJournal1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, AccModule.FrmTransaksiJournal)
                If IsNothing(ij) = False Then
                    If ij.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal And _
                    ij.TipeJournal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping Then
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
            i.TipeJournal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping
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

    Private Sub FrmTransaksiSalesCycle22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub FrmTransaksiJournal_Purchase21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub FrmMasterCompanyData1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub FrmSettingAP_AR1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    Private Sub ExchangeRateDefaultValuesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterDefaultValues1.Click
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndOfMonthJournalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_EndOfMonth1.Click
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub OtherToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Acquisition_Revaluation_OtherB1.Click
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLapAsset_DepresiasiPerPeriod1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmLapAsset_DepresiasiPerPeriod1.Click
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSettingPajak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSettingPajak11.Click
        Try
            Dim ij As HR.FrmMasterSettingPajak
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmMasterSettingPajak)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmMasterSettingPajak
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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub FrmTransaksiJournal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal14.Click
    '    Try
    '        Dim ij As AccModule.FrmTransaksiJournal
    '        Dim j As Windows.Forms.Form

    '        For Each j In Me.MdiChildren
    '            ij = TryCast(j, AccModule.FrmTransaksiJournal)
    '            If IsNothing(ij) = False Then
    '                ij.BringToFront()
    '                ij.Focus()
    '                Exit Sub
    '            End If
    '        Next

    '        Dim i As New AccModule.FrmTransaksiJournal
    '        ' Increment the caption counter.
    '        i.UserCOde = FrmLogin.UsernameTextBox.Text
    '        ' Set the caption to be unique.
    '        i.BackcolorAll = Core_Module.GetSMESSBackColorAll
    '        i.TipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
    '        i.TipeJournal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping
    '        i.MdiParent = Me
    '        i.Show()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

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
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MasterSlipDescriptionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As FrmMasterSalaryFormula_SlipDesc
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMasterSalaryFormula_SlipDesc)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMasterSalaryFormula_SlipDesc
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
            i.MdiParent = Me
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
                i.Close()
            Else
                i.Close()
                MsgBox("You've got a reminder for recurring journal !", MsgBoxStyle.Information, "Information")
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

    Private Sub TaxSettingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSettingPajak1.Click
        Try
            Dim ij As HR.FrmMasterSettingPajak
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, HR.FrmMasterSettingPajak)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New HR.FrmMasterSettingPajak
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
End Class