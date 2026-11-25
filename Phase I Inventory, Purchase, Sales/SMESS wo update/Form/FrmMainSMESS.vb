Option Explicit On
Option Strict Off

Imports Core_Module
Imports System.IO

Public Class FrmMainSMESS
    Dim mParentForm As FrmMain
    Dim RefColor1 As Color = core_module.GetSMESSColor
    Dim RefColor1_Border As Color = core_module.GetSMESSColor_Border
    Public CloseForm As Boolean = False
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub btnPurchase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPurchase.Click
        Me.PnlPurchase.BringToFront()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSales.Click
        Me.PnlSales.BringToFront()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInventory.Click
        Me.PnlInventory.BringToFront()
    End Sub

    Private Sub btnMaster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMaster.Click
        Me.PnlMasterData.BringToFront()
    End Sub

    Private Sub FrmMainSMESS_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason <> CloseReason.MdiFormClosing Then
            If Me.CloseForm = False Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub FrmMainSMESS_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Try
            Me.CekHakAkses()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekHakAkses()
        Dim TBL As New DSUser.tblhakaksesUserDataTable
        Dim DTA As New DSUserTableAdapters.tblhakaksesUserTableAdapter

        Try
            Me.mParentForm = Me.MdiParent
            'cek visibilitas
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'fil tabel
            DTA.Fill(TBL, FrmLogin.UsernameTextBox.Text)

            Me.CekHakAkses(Me.PnlInventory, TBL)
            Me.CekHakAkses(Me.PnlMasterData, TBL)
            Me.CekHakAkses(Me.PnlPurchase, TBL)
            Me.CekHakAkses(Me.PnlSales, TBL)
            Me.CekHakAkses(Me.PnlAccounting, TBL)
            Me.CekHakAkses(Me.PnlHRD, TBL)
            'Me.CekHakAkses(Me.GroupBox1, TBL)
            'Me.CekHakAkses(Me.GroupBox2, TBL)
            'Me.CekHakAkses(Me.GroupBox3, TBL)
            Me.CekHakAkses(Me.Panel1, TBL)
            Me.CekHakAkses(Me.Panel2, TBL)
            Me.CekHakAkses(Me.Panel3, TBL)
            Me.CekHakAkses(Me.Panel4, TBL)
            Me.CekHakAkses(Me.Panel5, TBL)
            Me.CekHakAkses(Me.Panel6, TBL)

            'jika tidak ada stock taking aktif, maka stock taking process tidak perlu ditampilkan
            Dim ObjST As New Inventory_Module.clsTransaksiStockTaking
            If ObjST.CekApakahAdaYangMasihAktif(False) = False Then
                FrmTransaksiStockTakingProcess1.Visible = False
                frmTransaksiStockTaking_Print1.Visible = False
            End If
            ObjST = Nothing

            'hide form login awal, komputer yang lambat kadang tidak melakukan proses hide form login
            FrmLogin.Hide()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Errror")
        Finally
            TBL.Dispose()
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub FrmMainSMESS_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = 1000
            Me.Timer1.Interval = 2000

            Me.UpdateTampilan()
            Me.CekHakAkses()
            'Me.rboAP.Checked = True
            'Me.GroupBox4.Visible = False
            Control.CheckForIllegalCrossThreadCalls = False
            Me.TxtReminderAPAR_FilterCode.CharacterCasing = CharacterCasing.Upper
            Me.cboPrintOption.SelectedIndex = 0
            Me.Timer1.Enabled = True

            Me.FrmTrans_Gaji11.Text = "Salary Calculation"
            Me.FrmTransaksiStockOutPurchaseReturn1.Text = "Stock Out Purc. Return"
            Me.FrmTransaksiStockOutPurchaseReturn2.Text = "Stock Out Purc. Return"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CekHakAkses(ByVal i As Windows.Forms.GroupBox, ByVal TBL As DSUser.tblhakaksesUserDataTable) As Boolean
        Dim obj As New ClsUser
        Dim j As Windows.Forms.Control
        Dim btn As Windows.Forms.Button
        Dim row As DSUser.tblhakaksesUserRow

        Try
            For k As Integer = 0 To i.Controls.Count - 1
                j = i.Controls(k)
                btn = TryCast(j, Windows.Forms.Button)
                If IsNothing(btn) = False Then
                    btn.Enabled = False
                    btn.Cursor = Cursors.Arrow
                    btn.BackColor = Color.WhiteSmoke
                    btn.FlatAppearance.BorderSize = 2
                    btn.FlatAppearance.BorderColor = Me.RefColor1

                    For Each row In TBL
                        If row.Nama = btn.Name.Substring(0, btn.Name.Length - 1) Then
                            btn.Enabled = True
                            btn.Cursor = Cursors.Hand
                            btn.FlatAppearance.BorderSize = 1
                            btn.FlatAppearance.BorderColor = Color.White
                            btn.BackColor = Me.RefColor1

                            Exit For
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function CekHakAkses(ByVal i As Windows.Forms.Panel, ByVal TBL As DSUser.tblhakaksesUserDataTable) As Boolean
        Dim obj As New ClsUser
        Dim j As Windows.Forms.Control
        Dim btn As Windows.Forms.Button
        Dim row As DSUser.tblhakaksesUserRow
        Dim Lbl As Label

        Try
            For k As Integer = 0 To i.Controls.Count - 1
                j = i.Controls(k)
                btn = TryCast(j, Windows.Forms.Button)
                If IsNothing(btn) = False Then
                    btn.Enabled = False
                    btn.Cursor = Cursors.Arrow
                    btn.FlatAppearance.BorderSize = 2
                    btn.FlatAppearance.BorderColor = Me.RefColor1
                    btn.BackColor = Color.WhiteSmoke

                    For Each row In TBL
                        If row.Nama = btn.Name.Substring(0, btn.Name.Length - 1) Then
                            btn.Enabled = True
                            btn.Cursor = Cursors.Hand
                            btn.FlatAppearance.BorderSize = 1
                            btn.FlatAppearance.BorderColor = Color.White
                            btn.BackColor = Me.RefColor1

                            Exit For
                        End If
                    Next
                Else
                    Lbl = TryCast(j, Label)
                    If IsNothing(Lbl) = False Then
                        If Lbl.Name.ToLower.StartsWith("frm") Then
                            Lbl.ForeColor = Color.Black
                            Lbl.Cursor = Cursors.Arrow
                        End If
                        For Each row In TBL
                            If row.Nama = Lbl.Name.Substring(0, Lbl.Name.Length - 1) Then
                                Lbl.Cursor = Cursors.Hand
                                Lbl.ForeColor = Color.White
                                If Lbl.Text = "Location" Then
                                    Lbl.ForeColor = RefColor1
                                End If

                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnAccounting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccounting.Click
        Me.PnlAccounting.BringToFront()
    End Sub

    Private Sub btnInternalMessage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInternalMessage.Click
        Try
            Me.PnlHRD.BringToFront()
            'Dim j As E_mail_Module.FrmMainEmail
            'Dim k As Windows.Forms.Form
            'For Each k In Me.mParentForm.MdiChildren
            '    j = TryCast(k, E_mail_Module.FrmMainEmail)
            '    If IsNothing(j) = False Then
            '        'activates form
            '        j.BringToFront()
            '        j.Focus()
            '        'keluar
            '        Exit Sub
            '    End If
            'Next
            'Dim i As New E_mail_Module.FrmMainEmail
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            'i.MdiParent = Me.mParentForm   i.ShowIcon = False
            'i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiGeneralJournal1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiGeneralJournal1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesCycle22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Sales
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Purchase21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiJournal_Purchase
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Adjust1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Closing1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInPurchase2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockInPurchase2.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInPurchase2
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInSalesReturn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockInSalesReturn2.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInSalesReturn2
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInAdj1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockInAdj1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockInAdj
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            ' Set the caption to be unique.
            '  i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockOutSales2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockOutSales2.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockOutSales
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockOutPurchaseReturn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockOutPurchaseReturn2.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockOutAdj1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockOutAdj1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockOutAdj
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            ' Set the caption to be unique.
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
            i = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockCard1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiStockCard1.Click
        Try
            Dim ij As Inventory_Module.FrmTransaksiStockCard
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKondisiStock1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_SalesDashboard1.Click
        Try
            Dim ij As Sales_Module.FrmTrans_SalesDashboard
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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

            i.BackColorAll = Me.BackColorAll

            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceDP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesInvoiceDP1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesInvoiceDP
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmTransaksiStockInSalesReturn2)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmTransaksiStockInSalesReturn2
            i.UserCode = FrmLogin.UsernameTextBox.Text
            'i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            'i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BAckColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOOtorisasiLv11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv11.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOOtorisasiLv21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv21.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOOtorisasiLv31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPOOtorisasiLv31.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPOOtorisasi
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            'i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOPDP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As Purchase_Module.FrmTransaksiAOPDP
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiGRN1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As Purchase_Module.FrmTransaksiGRN
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            ' i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPurchaseReturn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPurchaseReturn1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPurchaseReturn
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            'i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Lokal = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiMonthlyClosing1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiMonthlyClosing2
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BAckColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProductGroup1.Click
        Try
            Dim ij As FrmMasterProductGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterBrand1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterBrand1.Click
        Try
            Dim ij As FrmMasterBrand
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackCOlorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSubBrand1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSubBrand1.Click
        Try
            Dim ij As FrmMasterSubBrand
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCategory1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCategory1.Click
        Try
            Dim ij As frmMasterCategory
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProduct1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProduct1.Click
        Try
            Dim ij As frmMasterProduct
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomerGroup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerGroup1.Click
        Try
            Dim ij As FrmMasterCustomerGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomerType1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerType1.Click
        Try
            Dim ij As FrmMasterCustomerType
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomerSegmentation1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerSegmentation1.Click
        Try
            Dim ij As FrmMasterCustomerSegmentation
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalesArea1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesArea1.Click
        Try
            Dim ij As FrmMasterSalesArea
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomer1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomer1.Click
        Try
            Dim ij As FrmMasterCustomer
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSupplier1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSupplier1.Click
        Try
            Dim ij As FrmMasterSupplier
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterUnit1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterUnit1.Click
        Try
            Dim ij As FrmMasterUnit
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCurrency1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCurrency1.Click
        Try
            Dim ij As FrmMasterCurrency2
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAkun1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterAkun1.Click
        Try
            Dim ij As AccModule.FrmMasterAkun
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_View1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_View1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal_View
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackCOlorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesInvoiceQuick1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesInvoiceQuick1.Click
        Try
            Dim ij As Sales_Module.FrmTransaksiSalesInvoiceQuick
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiAOP_Quick1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiAOP_Quick1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiAOP_Quick2
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPurchaseReturn_Quick2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ij As Purchase_Module.FrmTransaksiPurchaseReturn_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnReminder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReminder.Click
        Me.PnlReminder.BringToFront()
        Me.rboAP.Checked = False
        Me.rboAP.Checked = True
    End Sub

    Private Sub SetRectangleShape(ByVal X1 As Integer, ByVal X2 As Integer, ByVal Y1 As Integer, ByVal Y2 As Integer, ByVal Space As Integer, _
    ByVal vRecShape As PowerPacks.RectangleShape)
        Try
            vRecShape.Width = X2 - X1 + (2 * Space)
            vRecShape.Height = Y2 - Y1 + (2 * Space)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateTampilan()
        Dim RefColor As Color = Core_Module.GetSMESSColor
        Dim BorderColor As Color = Core_Module.GetSMESSColor_Border
        Dim Space As Integer = 7

        Try
            Me.Label12.ForeColor = RefColor
            Me.btnMaster.FlatAppearance.BorderColor = BorderColor
            Me.btnPurchase.FlatAppearance.BorderColor = BorderColor
            Me.btnSales.FlatAppearance.BorderColor = BorderColor
            Me.btnAccounting.FlatAppearance.BorderColor = BorderColor
            Me.btnReminder.FlatAppearance.BorderColor = BorderColor
            Me.btnInternalMessage.FlatAppearance.BorderColor = BorderColor
            Me.btnInventory.FlatAppearance.BorderColor = BorderColor

            Me.LblSales.BackColor = RefColor
            Me.LblAccounting.BackColor = RefColor
            Me.LblPurchase.BackColor = RefColor
            Me.LblMasterData1.BackColor = RefColor
            Me.LblHRD.BackColor = RefColor
            Me.LblInventory.BackColor = RefColor
            Me.LblReminder.BackColor = RefColor
            Me.LblMasterData.ForeColor = RefColor
            Me.Label1.ForeColor = RefColor
            Me.Label2.ForeColor = RefColor
            Me.Label8.ForeColor = RefColor
            Me.Label3.ForeColor = RefColor
            Me.Label10.ForeColor = RefColor
            Me.Label11.ForeColor = RefColor
            Me.Label13.ForeColor = RefColor
            Me.Label9.ForeColor = RefColor

            'master product
            Me.LblProductMaster.BackColor = RefColor
            Me.FrmMasterProduct2.BackColor = RefColor
            Me.FrmMasterCategory2.BackColor = RefColor
            Me.FrmMasterSubBrand2.BackColor = RefColor
            Me.FrmMasterBrand2.BackColor = RefColor
            Me.LblProductMaster.BackColor = RefColor
            Me.FrmMasterProductGroup2.BackColor = RefColor
            'master HR
            Me.LblHR.BackColor = RefColor
            Me.FrmMasterKompetensi11.BackColor = RefColor
            Me.FrmMasterKaryawan12.BackColor = RefColor
            Me.FrmMasterCuti11.BackColor = RefColor
            Me.FrmMasterSalaryFormula_OtherDeduction11.BackColor = RefColor
            Me.FrmMasterJobPosition12.BackColor = RefColor
            Me.FrmMasterSalaryFormula11.BackColor = RefColor
            'master Customer
            Me.LblCustomer.BackColor = RefColor
            Me.FrmMasterCustomer2.BackColor = RefColor
            Me.FrmMasterCustomerGroup2.BackColor = RefColor
            Me.FrmMasterSalesArea2.BackColor = RefColor
            Me.FrmMasterCustomerSegmentation2.BackColor = RefColor
            Me.FrmMasterCustomerType2.BackColor = RefColor
            'master inventory
            Me.LblInventoryMaster.BackColor = RefColor
            Me.FrmMasterGudang2.BackColor = RefColor
            Me.FrmMasterStorageArea2.ForeColor = RefColor
            Me.FrmMasterStorageArea2.BackColor = Color.White
            'master accounting
            Me.LblAccountingMaster.BackColor = RefColor
            Me.FrmTransaksiBudget2.BackColor = RefColor
            Me.FrmMasterAkun2.BackColor = RefColor
            'other master
            Me.LblOtherMaster.BackColor = RefColor
            Me.FrmMasterCurrency2.BackColor = RefColor
            Me.FrmMasterDepartement2.BackColor = RefColor
            Me.FrmMasterProses1.BackColor = RefColor
            Me.FrmTransaksiSalesBudget1.BackColor = RefColor
            Me.FrmMasterSalesPersonGroup1.BackColor = RefColor
            Me.FrmMasterSalesPerson1.BackColor = RefColor
            Me.SetPosisi(Me.RectangleShape1, Me.FrmMasterSupplier2)
            Me.SetPosisi(Me.RectangleShape2, Me.FrmMasterDepartement2)
            Me.SetPosisi(Me.RectangleShape3, Me.FrmMasterSalesPersonGroup1)
            Me.SetPosisi(Me.RectangleShape4, Me.FrmMasterUnit2)
            Me.SetPosisi(Me.RectangleShape5, Me.FrmMasterSalesPerson1)
            Me.SetPosisi(Me.RectangleShape6, Me.FrmMasterCurrency2)
            Me.SetPosisi(Me.RectangleShape7, Me.FrmTransaksiSalesBudget1)
            Me.SetPosisi(Me.RectangleShape8, Me.FrmMasterProses1)
            Me.RectangleShape4.Left = Me.RectangleShape1.Left
            Me.RectangleShape6.Left = Me.RectangleShape1.Left
            Me.RectangleShape8.Left = Me.RectangleShape1.Left
            Me.RectangleShape3.Left = Me.RectangleShape2.Left
            Me.RectangleShape5.Left = Me.RectangleShape2.Left
            Me.RectangleShape7.Left = Me.RectangleShape2.Left

            'purchase cycle
            Me.RectangleShape18.BorderColor = RefColor
            Me.RectangleShape18.Width = Me.FrmTransaksiPOOtorisasiLv11.Width + 2 * Space
            Me.RectangleShape18.Height = Me.Label16.Height + Me.FrmTransaksiPOOtorisasiLv11.Height + _
            Me.FrmTransaksiPOOtorisasiLv21.Height + Me.FrmTransaksiPOOtorisasiLv31.Height + 3 * Space
            Me.RectangleShape18.Location = New Point(Me.FrmTransaksiPOOtorisasiLv11.Left - Space, Me.Label16.Top + Me.Label16.Height / 2 + 1)

            Me.RectangleShape19.BorderColor = RefColor
            Me.RectangleShape19.Width = Me.FrmTransaksiAOPLokal1.Width + 2 * Space
            Me.RectangleShape19.Height = Me.Label17.Height + Me.FrmTransaksiAOPLokal1.Height + _
            Me.FrmTransaksiAOPRImport1.Height + 3 * Space
            Me.RectangleShape19.Location = New Point(Me.FrmTransaksiAOPLokal1.Left - Space, Me.Label17.Top + Me.Label17.Height / 2 + 1)

            Me.RectangleShape16.BorderColor = RefColor
            Me.RectangleShape16.Location = New Point(FrmTrans_MaterialPlanningDashboard1.Left - Space, Me.Label6.Top + Me.Label6.Height / 2 + 1)
            Me.SetRectangleShape(FrmTrans_MaterialPlanningDashboard1.Left, FrmTransaksiPurchaseReturn_Quick1.Right + Space, _
            Me.Label6.Top + Me.Label6.Height / 2 + 1, FrmTrans_MaterialPlanningDashboard1.Bottom, Space, Me.RectangleShape16)

            Me.Label5.ForeColor = RefColor
            Me.Label6.ForeColor = RefColor
            Me.Label16.ForeColor = RefColor
            Me.Label17.ForeColor = RefColor
            Me.FrmTransaksiPO1.BackColor = RefColor
            Me.FrmTransaksiPOOtorisasiLv11.BackColor = RefColor
            Me.FrmTransaksiPOOtorisasiLv21.BackColor = RefColor
            Me.FrmTransaksiPOOtorisasiLv31.BackColor = RefColor
            Me.FrmTransaksiPurchaseReturn1.BackColor = RefColor
            Me.FrmTransaksiPOC1.BackColor = RefColor
            Me.FrmTransaksiPackingList1.BackColor = RefColor
            Me.FrmTransaksiStockInPurchase1.BackColor = RefColor
            Me.FrmTransaksiStockOutPurchaseReturn1.BackColor = RefColor
            Me.FrmTransaksiAOPLokal1.BackColor = RefColor
            Me.FrmTransaksiAOPRImport1.BackColor = RefColor
            Me.FrmTransaksiMonthlyClosing1.BackColor = RefColor
            Me.FrmTransaksiHistoricalPurchase1.BackColor = RefColor
            Me.FrmTransaksiAOP_Quick1.BackColor = RefColor
            Me.FrmTransaksiPurchaseReturn_Quick1.BackColor = RefColor

            'sales cycle
            Me.RectangleShape17.BorderColor = RefColor
            Me.RectangleShape17.Location = New Point(FrmTransaksiQuotationOtorisasiSalesStaff1.Left - Space - 2, Me.Label15.Top + Me.Label15.Height / 2 + 1)
            Me.SetRectangleShape(FrmTransaksiQuotationOtorisasiSalesStaff1.Left, FrmTransaksiQuotationOtorisasiSalesStaff1.Right + Space, _
            FrmTransaksiQuotationOtorisasiSalesStaff1.Top, FrmTransaksiQuotationOtorisasiGeneralManager1.Bottom + Space, Space, Me.RectangleShape17)

            Me.RectangleShape15.BorderColor = RefColor
            Me.RectangleShape15.Location = New Point(FrmTrans_SalesDashboard1.Left - Space, Me.Label7.Top + Me.Label7.Height / 2 + 1)
            Me.SetRectangleShape(FrmTrans_SalesDashboard1.Left, FrmTransaksiSalesReturn_Quick1.Right + Space + Space, _
            FrmTrans_SalesDashboard1.Top, FrmTrans_SalesDashboard1.Bottom + Space + Space, Space, Me.RectangleShape15)

            Me.Label4.ForeColor = RefColor
            Me.Label7.ForeColor = RefColor
            Me.Label15.ForeColor = RefColor
            Me.FrmTransaksiQuotation1.BackColor = RefColor
            Me.FrmTransaksiSO1.BackColor = RefColor
            Me.FrmTransaksiQuotationOtorisasiGeneralManager1.BackColor = RefColor
            Me.FrmTransaksiQuotationOtorisasiSalesManager1.BackColor = RefColor
            Me.FrmTransaksiQuotationOtorisasiSalesStaff1.BackColor = RefColor
            Me.FrmTransaksiDO1.BackColor = RefColor
            Me.FrmTransaksiStockOutSales1.BackColor = RefColor
            Me.FrmTransaksiSalesInvoice1.BackColor = RefColor
            Me.FrmTransaksiSalesReturn1.BackColor = RefColor
            Me.FrmTransaksiStockInSalesReturn1.BackColor = RefColor
            Me.FrmTransaksiSalesInvoiceQuick1.BackColor = RefColor
            Me.FrmTransaksiSalesReturn_Quick1.BackColor = RefColor
            Me.FrmTransaksiSalesInvoiceDP1.BackColor = RefColor
            Me.FrmTransaksiBookingInv1.BackColor = RefColor
            Me.FrmTrans_SalesDashboard1.BackColor = RefColor

            'inventory cycle
            Me.RectangleShape11.BorderColor = RefColor
            Me.RectangleShape11.Location = New Point(Me.RectangleShape11.Left, Me.Label10.Top + Me.Label10.Height / 2 + 1)
            Me.RectangleShape12.BorderColor = RefColor
            Me.RectangleShape12.Location = New Point(Me.RectangleShape12.Left, Me.Label2.Top + Me.Label2.Height / 2 + 1)
            Me.RectangleShape13.BorderColor = RefColor
            Me.RectangleShape13.Location = New Point(Me.RectangleShape13.Left, Me.Label8.Top + Me.Label8.Height / 2 + 1)
            Me.RectangleShape14.BorderColor = RefColor
            Me.RectangleShape14.Location = New Point(Me.RectangleShape14.Left, Me.Label3.Top + Me.Label3.Height / 2 + 1)

            Me.SetRectangleShape(FrmTransaksiStockInPurchase2.Left, frmTransaksiStockTaking_Print1.Right + Space + Space, _
            FrmTransaksiStockInPurchase2.Top, FrmTransaksiStockInPurchase2.Bottom + Space + Space, Space, Me.RectangleShape11)
            Me.SetRectangleShape(FrmTransaksiStockInPurchase2.Left, frmTransaksiStockTaking_Print1.Right + Space + Space, _
            FrmTransaksiStockOutSales2.Top, FrmTransaksiStockOutSales2.Bottom + Space + Space, Space, Me.RectangleShape12)
            Me.SetRectangleShape(FrmTransaksiStockInPurchase2.Left, frmTransaksiStockTaking_Print1.Right + Space + Space, _
            frmTransaksiStockTaking1.Top, frmTransaksiStockTaking1.Bottom + Space + Space, Space, Me.RectangleShape13)
            Me.SetRectangleShape(FrmTransaksiStockInPurchase2.Left, frmTransaksiStockTaking_Print1.Right + Space + Space, _
            FrmTransaksiStockCard1.Top, FrmTransaksiStockCard1.Bottom + Space + Space, Space, Me.RectangleShape14)

            Me.FrmTransaksiStockInPurchase2.BackColor = RefColor
            Me.FrmTransaksiStockInAdj1.BackColor = RefColor
            Me.FrmTransaksiStockInSalesReturn2.BackColor = RefColor
            Me.FrmTransaksiStockOutSales1.BackColor = RefColor
            Me.FrmTransaksiStockOutPurchaseReturn2.BackColor = RefColor
            Me.FrmTransaksiStockOutAdj1.BackColor = RefColor
            Me.FrmTransaksiStockBundling1.BackColor = RefColor
            Me.FrmTransaksiStockPindah1.BackColor = RefColor
            Me.FrmMasterKondisiStock1.BackColor = RefColor
            Me.FrmTransaksiStockCard1.BackColor = RefColor
            Me.frmTransaksiStockTaking1.BackColor = RefColor
            Me.frmTransaksiStockTag1.BackColor = RefColor
            Me.FrmTransaksiStockTakingProcess1.BackColor = RefColor
            Me.frmTransaksiStockTaking_Print1.BackColor = RefColor

            'accounting cycle
            Me.RectangleShape10.BorderColor = RefColor
            Me.RectangleShape10.Location = New Point(FrmTransaksiClosing1.Left - Space, Me.Label9.Top + Me.Label9.Height / 2 + 5)
            Me.SetRectangleShape(FrmTransaksiClosing1.Left, FrmTransaksiReminder1.Right + Space, FrmTransaksiClosing1.Top, _
            FrmTransaksiClosing1.Bottom + Space, Space, Me.RectangleShape10)
            Me.FrmTransaksiJournal_CashBank1.BackColor = RefColor
            Me.FrmTransaksiJournal_Closing1.BackColor = RefColor
            Me.FrmTransaksiJournal_EndOfMonth1.BackColor = RefColor
            Me.FrmTransaksiJournal_Purchase11.BackColor = RefColor
            Me.FrmTransaksiJournal_View1.BackColor = RefColor
            Me.FrmTransaksiGeneralJournal1.BackColor = RefColor
            Me.FrmTransaksiSalesCycle12.BackColor = RefColor
            Me.FrmTransaksiClosing1.BackColor = RefColor
            Me.FrmTransaksiTagihan1.BackColor = RefColor
            Me.FrmTransaksiReminder1.BackColor = RefColor

            'HRD cycle
            Me.RectangleShape9.BorderColor = RefColor
            Me.RectangleShape9.Location = New Point(FrmTrans_Attendance_DataImport11.Left - Space, Me.Label14.Top + Me.Label14.Height / 2 + 1)
            Me.SetRectangleShape(FrmTrans_Attendance_DataImport11.Left, FrmTrans_Gaji_Quick11.Right + Space, _
            FrmTrans_Attendance_DataImport11.Top, FrmTrans_Attendance_DataImport11.Bottom + Space + Space, Space, Me.RectangleShape9)

            Me.Label14.ForeColor = RefColor
            Me.Label22.ForeColor = RefColor
            Me.FrmMasterWorkShift11.BackColor = RefColor
            Me.FrmMasterWorkShift_Allocate11.BackColor = RefColor
            Me.FrmTrans_Attendance11.BackColor = RefColor
            Me.FrmTransaksiLoan11.BackColor = RefColor
            Me.FrmTransaksiLoan_Payment11.BackColor = RefColor
            Me.FrmTransaksiCuti_Pengajuan11.BackColor = RefColor
            Me.FrmTransaksiCuti_Approval11.BackColor = RefColor
            Me.FrmTrans_Gaji11.BackColor = RefColor
            Me.FrmTrans_Payment11.BackColor = RefColor
            Me.FrmTrans_Export_eSPT11.BackColor = RefColor
            Me.FrmTrans_Attendance_DataImport11.BackColor = RefColor
            Me.FrmTrans_Gaji_Quick11.BackColor = RefColor

            Me.PictureBox1.BackColor = RefColor
            Me.Panel1.BackColor = RefColor
            Me.Panel2.BackColor = RefColor
            Me.Panel3.BackColor = RefColor
            Me.Panel4.BackColor = RefColor
            Me.Panel5.BackColor = RefColor
            Me.Panel6.BackColor = RefColor

            Me.Panel1.Visible = True
            Me.Panel2.Visible = True
            Me.Panel3.Visible = True
            Me.Panel4.Visible = True
            Me.Panel5.Visible = True
            Me.Panel6.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMaster_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMaster.MouseEnter
        Me.btnMaster.FlatAppearance.BorderSize = 2
    End Sub

    Private Sub btnMaster_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMaster.MouseLeave
        Me.btnMaster.FlatAppearance.BorderSize = 0
    End Sub

    Private Sub btnReminder_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReminder.MouseEnter
        Me.btnReminder.FlatAppearance.BorderSize = 2
    End Sub

    Private Sub btnReminder_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReminder.MouseLeave
        Me.btnReminder.FlatAppearance.BorderSize = 0
    End Sub

    Private Sub btnInventory_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInventory.MouseEnter
        Me.btnInventory.FlatAppearance.BorderSize = 2
    End Sub

    Private Sub btnInventory_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInventory.MouseLeave
        Me.btnInventory.FlatAppearance.BorderSize = 0
    End Sub


    Private Sub btnInternalMessage_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInternalMessage.MouseEnter
        Me.btnInternalMessage.FlatAppearance.BorderSize = 2
    End Sub

    Private Sub btnInternalMessage_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInternalMessage.MouseLeave
        Me.btnInternalMessage.FlatAppearance.BorderSize = 0
    End Sub

    Private Sub btnPurchase_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPurchase.MouseEnter
        Me.btnPurchase.FlatAppearance.BorderSize = 2
    End Sub

    Private Sub btnPurchase_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPurchase.MouseLeave
        Me.btnPurchase.FlatAppearance.BorderSize = 0
    End Sub

    Private Sub btnSales_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSales.MouseEnter
        Me.btnSales.FlatAppearance.BorderSize = 2
    End Sub

    Private Sub btnSales_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSales.MouseLeave
        Me.btnSales.FlatAppearance.BorderSize = 0
    End Sub

    Private Sub btnAccounting_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAccounting.MouseEnter
        Me.btnAccounting.FlatAppearance.BorderSize = 2
    End Sub

    Private Sub btnAccounting_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAccounting.MouseLeave
        Me.btnAccounting.FlatAppearance.BorderSize = 0
    End Sub

    Private Sub FrmTransaksiPurchaseReturn_Quick1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiPurchaseReturn_Quick1.Click
        Try
            Dim ij As Purchase_Module.FrmTransaksiPurchaseReturn_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKondisiStock1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKondisiStock1.Click
        Try
            Dim ij As Inventory_Module.FrmMasterKondisiStock
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_Closing1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiJournal_Closing1.Click
        Try
            Dim ij As AccModule.FrmTransaksiJournal
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterWorkShift11.Click
        Try
            Dim ij As Core_Module.FrmMasterWorkShift
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_Allocate11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterWorkShift_Allocate11.Click
        Try
            Dim ij As Core_Module.FrmMasterWorkShift_Allocate
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Attendance11.Click
        Try
            Dim ij As HR.FrmTrans_Attendance
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Gaji11.Click
        Try
            Dim ij As HR.FrmTrans_Gaji
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLoan11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiLoan11.Click
        Try
            Dim ij As HR.FrmTransaksiLoan
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiLoan_Payment11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiLoan_Payment11.Click
        Try
            Dim ij As HR.FrmTransaksiLoan_Payment
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Payment11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Payment11.Click
        Try
            Dim ij As HR.FrmTransaksiPayment
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Export_eSPT11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Export_eSPT11.Click
        Try
            Dim ij As HR.FrmTrans_Export_eSPT
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            'i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiCuti_Pengajuan11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiCuti_Pengajuan11.Click
        Dim ObjPengajuanCuti As New HR.ClsTrans_Cuti_Pengajuan
        Dim ObjUser As New Core_Module.ClsUser
        Try
            Dim ij As HR.FrmTransaksiCuti_Pengajuan
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
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

    Private Sub FrmTransaksiCuti_Approval11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiCuti_Approval11.Click
        Try
            Dim ij As HR.FrmTransaksiCuti_Approval
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_DataImport11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Attendance_DataImport11.Click
        Try
            Dim ij As HR.FrmTrans_Attendance_DataImport
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_Quick11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Gaji_Quick11.Click
        Try
            Dim ij As HR.FrmTrans_Gaji_Quick
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.Timer1.Enabled = False
            Me.PnlMasterData.BringToFront()
            Me.Panel7.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProductGroup2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProductGroup2.Click
        Try
            If Me.FrmMasterProductGroup2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If
            Dim ij As FrmMasterProductGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterBrand2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterBrand2.Click
        Try
            If FrmMasterBrand2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterBrand
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackCOlorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSubBrand2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSubBrand2.Click
        Try
            If Me.FrmMasterBrand2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterSubBrand
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackcolorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCategory2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCategory2.Click
        Try
            If Me.FrmMasterCategory2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As frmMasterCategory
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProduct2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProduct2.Click
        Try
            If Me.FrmMasterProduct2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As frmMasterProduct
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblHR_EmployeeData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKaryawan12.Click
        Try
            If FrmMasterKaryawan12.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Core_Module.FrmMasterKaryawan
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterJobPosition12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterJobPosition12.Click
        Try
            If FrmMasterJobPosition12.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Core_Module.FrmMasterJobPosition
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKompetensi11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterKompetensi11.Click
        Try
            If Me.FrmMasterKompetensi11.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Core_Module.FrmMasterKompetensi
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalaryFormula11.Click
        Try
            If FrmMasterSalaryFormula11.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Core_Module.FrmMasterSalaryFormula
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalaryFormula_OtherDeduction11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalaryFormula_OtherDeduction11.Click
        Try
            If FrmMasterSalaryFormula_OtherDeduction11.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Core_Module.FrmMasterSalaryFormula_OtherDeduction
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCuti11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCuti11.Click
        Try
            If FrmMasterCuti11.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Core_Module.FrmMasterCuti
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalesArea2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesArea2.Click
        Try
            If FrmMasterSalesArea2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterSalesArea
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomerGroup2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerGroup2.Click
        Try
            If FrmMasterCustomerGroup2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterCustomerGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomerType2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerType2.Click
        Try
            If FrmMasterCustomerType2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterCustomerType
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomerSegmentation2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomerSegmentation2.Click
        Try
            If FrmMasterCustomerSegmentation2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterCustomerSegmentation
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterGudang2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterGudang2.Click
        Try
            If FrmMasterGudang2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterGudang
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterStorageArea2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterStorageArea2.Click
        Try
            If FrmMasterStorageArea2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterStorageArea
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterAkun2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterAkun2.Click
        Try
            If FrmMasterAkun2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As AccModule.FrmMasterAkun
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetPosisi(ByVal vRec As Microsoft.VisualBasic.PowerPacks.RectangleShape, ByVal Lbl As Label)
        Try
            vRec.Location = New Point(Lbl.Left + (Lbl.Width / 2) - (vRec.Size.Width / 2), Lbl.Top + (Lbl.Height / 2) - (vRec.Height / 2))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiBudget2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiBudget2.Click
        Try
            If FrmTransaksiBudget2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As AccModule.FrmMasterBudget
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSupplier2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSupplier2.Click
        Try
            If FrmMasterSupplier2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterSupplier
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterDepartement2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterDepartement2.Click
        Try
            If FrmMasterDepartement2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterDepartement
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterUnit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterUnit2.Click
        Try
            If FrmMasterUnit2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterUnit
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalesPersonGroup1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesPersonGroup1.Click
        Try
            If FrmMasterSalesPersonGroup1.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterSalesPersonGroup
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCurrency2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCurrency2.Click
        Try
            If FrmMasterCurrency2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterCurrency2
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterSalesPerson1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterSalesPerson1.Click
        Try
            If FrmMasterSalesPerson1.Cursor = Cursors.Arrow Then
                Exit Sub
            End If
            Dim ij As FrmMasterSalesPerson
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterProses1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProses1.Click
        Try
            If FrmMasterProses1.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Core_Module.FrmMasterProses
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.UserCode = FrmLogin.UsernameTextBox.Text
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSalesBudget1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTransaksiSalesBudget1.Click
        Try
            If FrmTransaksiSalesBudget1.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As Sales_Module.FrmTransaksiSalesBudget
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCustomer2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterCustomer2.Click
        Try
            If FrmMasterCustomer2.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As FrmMasterCustomer
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_PriceList1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmMasterProduct_PriceList1.Click
        Try
            If frmMasterProduct_PriceList1.Cursor = Cursors.Arrow Then
                Exit Sub
            End If

            Dim ij As frmMasterProduct_PriceList
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmTransaksiStockTaking_Print1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles frmTransaksiStockTaking_Print1.Click
        Try
            Dim ij As Inventory_Module.frmTransaksiStockTaking_Print
            Dim j As Windows.Forms.Form

            For Each j In Me.mParentForm.MdiChildren
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
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.mParentForm
            i.ShowIcon = False
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class