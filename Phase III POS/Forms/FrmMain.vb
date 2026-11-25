Imports Core_Module, AccModule
Imports System.IO

Public Class FrmMain
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
            For Each i In Me.MdiChildren
                i.Close()
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

    Private Function CekTollStripMenu(ByVal i As ToolStripMenuItem, ByVal TBL As DSUser.tblhakaksesUserDataTable) As Boolean
        Dim obj As New ClsUser
        Dim j As ToolStripMenuItem
        Dim vs As Boolean = False
        Dim row As DSUser.tblhakaksesUserRow
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

    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub FrmMasterProduct1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterProduct1.Click
        Try
            Dim ij As FrmMast_Cabang
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMast_Cabang)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMast_Cabang
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterDiskon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmMasterDiskon1.Click
        Try
            Dim ij As FrmMast_Discount_Main
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMast_Discount_Main)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMast_Discount_Main
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SessionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SessionToolStripMenuItem.Click
        Try
            Dim ij As FrmMast_Session
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMast_Session)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMast_Session
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub POSSettingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POSSettingToolStripMenuItem.Click
        Try
            Dim ij As FrmTrans_POS_Settings
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmTrans_POS_Settings)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmTrans_POS_Settings
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SessionToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SessionToolStripMenuItem1.Click
        Try
            Dim ij As FrmTrans_Session
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmTrans_Session)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmTrans_Session
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub POSPaymentDisbursementToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POSPaymentDisbursementToolStripMenuItem.Click
        FrmTrans_POS_PaymentDisbursment.ShowDialog()
    End Sub

    Private Sub POSPaymentDisbursemenEditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POSPaymentDisbursemenEditToolStripMenuItem.Click
        FrmTrans_POS_PaymentDisbursment_Edit.ShowDialog()
    End Sub

    Private Sub PaymentOptionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FrmMast_PaymentOption.ShowDialog()
    End Sub

    Private Sub MemberToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MemberToolStripMenuItem.Click
        Try
            Dim ij As FrmMast_Member
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMast_Member)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMast_Member
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FrmMast_Discount_Main.ShowDialog()
    End Sub

    Private Sub PayoutDropCashToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PayoutDropCashToolStripMenuItem.Click
        Try
            Dim ij As FrmMast_PayOutDropCash
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMast_PayOutDropCash)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMast_PayOutDropCash
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PaymentOptionToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PaymentOptionToolStripMenuItem1.Click
        Try
            Dim ij As FrmMast_PaymentOption
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmMast_PaymentOption)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmMast_PaymentOption
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub POSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POSToolStripMenuItem.Click
        Dim ObjSession As New ClsTrans_Session
        Try
            Dim ij As FrmTrans_POS
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmTrans_POS)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmTrans_POS
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            If ObjSession.GetJmlByCashierCode(i.UserCode) > 0 Then
                i.Show()
            Else
                MsgBox(i.UserCode.ToString + " doesn't have an active session, you cannot access transaction POS.", MsgBoxStyle.Information, "Information")
                i.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSession = Nothing
        End Try
    End Sub

    Private Sub PengecekanFormatDecimalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PengecekanFormatDecimalToolStripMenuItem.Click
        Try
            Dim ij As FrmA_CekRounding
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmA_CekRounding)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmA_CekRounding
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ProductToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductToolStripMenuItem.Click
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
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Data_Transfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FrmTrans_Data_Transfer1.Click
        Try
            Dim ij As FrmTrans_Data_Transfer
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, FrmTrans_Data_Transfer)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    Exit Sub
                End If
            Next

            Dim i As New FrmTrans_Data_Transfer
            ' Increment the caption counter.
            'i.UserCode = FrmLogin.UsernameTextBox.Text
            i.UserCode = "USER"
            ' Set the caption to be unique.
            i.BackColorAll = Color.LightGoldenrodYellow
            i.MdiParent = Me
            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class