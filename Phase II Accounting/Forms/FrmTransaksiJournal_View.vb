Option Explicit On
Option Strict Off

Public Class FrmTransaksiJournal_View
    Public UserCode As String = ""
    Public BackCOlorAll As Color = Color.LightSeaGreen
    Dim ObjJurnal As New ClsTransaksiJournal
    Dim TBL As New DseTransaksiJournal_Cetak.TblCetak_JurnalDataTable
    Dim buffer As String = ""
    Dim ObjAkun As New ClsMasterAkun
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Private Sub cboFilterBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFilterBy.SelectedIndexChanged
        Try
            If Me.cboFilterBy.SelectedIndex = 0 Then
                Me.TxtFrom.Visible = False
                Me.TxtFrom.TabStop = False
                Me.TxtTo.Visible = False
                Me.TxtTo.TabStop = False
                Me.btnFindFrom.Visible = False
                Me.BtnFindTo.Visible = False
                Me.dtpFrom.Visible = True
                Me.dtpFrom.TabStop = True
                Me.dtpTo.Visible = True
                Me.dtpTo.TabStop = True
            Else
                Me.TxtFrom.Visible = True
                Me.TxtFrom.TabStop = True
                Me.TxtTo.Visible = True
                Me.TxtTo.TabStop = True
                Me.btnFindFrom.Visible = True
                Me.BtnFindTo.Visible = True
                Me.dtpFrom.Visible = False
                Me.dtpFrom.TabStop = False
                Me.dtpTo.Visible = False
                Me.dtpTo.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiJournal_View_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_Journal_View, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_Journal_View_Format, Me.cboFormat.SelectedIndex.ToString, Me.UserCode)
            If Me.chkShowRemarks.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_Journal_View, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_Journal_View_ShowRemarks, "1", Me.UserCode)
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_Journal_View, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_Journal_View_ShowRemarks, "0", Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiJournal_View_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Setting As String = ""

        Try
            Me.cboFilterBy.SelectedIndex = 0
            Me.cboFormat.SelectedIndex = 0

            If Me.ObjJurnal.GetJml(ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan) > 0 Or _
            Me.ObjJurnal.GetJml(ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Loan_Payment) > 0 Or _
            Me.ObjJurnal.GetJml(ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.HR_Payroll_Payment) > 0 Then
                Me.clbTipeJournal.Items.Add("Loan Journal")
                Me.clbTipeJournal.Items.Add("Loan Payment Journal")
                Me.clbTipeJournal.Items.Add("Salary Payment Journal")
            End If
            Me.chkJournalType.Checked = True
            Me.llbCheckAll.Location = New Point(Me.llbCheckAll.Location.X - 15, 12)
            Me.llbUncheckAll.Location = New Point(Me.llbUncheckAll.Location.X - 15, 12)

            Setting = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_Journal_View, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_Journal_View_Format, Me.UserCode)
            If Setting <> "" Then
                Me.cboFormat.SelectedIndex = Setting
            End If
            Setting = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_Journal_View, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_Journal_View_ShowRemarks, Me.UserCode)
            If Setting <> "" Then
                If Setting = "0" Then
                    Me.chkShowRemarks.Checked = False
                Else
                    Me.chkShowRemarks.Checked = True
                End If
            End If

            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            'Dim CheckedItem As String = ""
            'For i As Integer = 0 To clbTipeJournal.CheckedItems.Count - 1
            'CheckedItem = clbTipeJournal.CheckedItems.Item(i)
            'Next
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Dim jumlahcheck As Integer = 0
            For i As Integer = 0 To clbTipeJournal.CheckedItems.Count - 1
                jumlahcheck += i + 1
            Next
            Me.TxtJournalDescription.Text = Me.TxtJournalDescription.Text.Trim
            Me.TxtRemarksItem.Text = Me.TxtRemarksItem.Text.Trim

            If Me.cboFormat.SelectedIndex = 0 Then
                If jumlahcheck > 0 Then
                    Me.ObjJurnal.CetakJournal_OriginalCurrency(Me.dtpFrom.Value, Me.dtpTo.Value, TBL, Me.crvJournal, Me.clbTipeJournal, Me.chkShowRemarks.Checked, _
                     Me.TxtJournalDescription.Text, Me.TxtRemarksItem.Text, Me.dgvItem, Me, Me.UserCode)
                Else
                    MsgBox("Please choose journal type", MsgBoxStyle.Information, "Information")
                    Me.crvJournal.ReportSource = Nothing
                    Me.crvJournal.Refresh()
                    Exit Sub
                End If
            Else
                If jumlahcheck > 0 Then
                    If Me.cboFormat.SelectedIndex = 1 Then
                        Me.ObjJurnal.CetakJournal(Me.dtpFrom.Value, Me.dtpTo.Value, TBL, Me.crvJournal, Me.clbTipeJournal, True, Me.chkShowRemarks.Checked, _
                        Me.TxtJournalDescription.Text, Me.TxtRemarksItem.Text, Me.dgvItem, Me, Me.UserCode)
                    Else
                        Me.ObjJurnal.CetakJournal(Me.dtpFrom.Value, Me.dtpTo.Value, TBL, Me.crvJournal, Me.clbTipeJournal, False, Me.chkShowRemarks.Checked, _
                        Me.TxtJournalDescription.Text, Me.TxtRemarksItem.Text, Me.dgvItem, Me, Me.UserCode)
                    End If
                Else
                    MsgBox("Please choose journal type", MsgBoxStyle.Information, "Information")
                    Me.crvJournal.ReportSource = Nothing
                    Me.crvJournal.Refresh()
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub llbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.clbTipeJournal.Items.Count - 1
                Me.clbTipeJournal.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll.LinkClicked
        Try
            For i As Integer = 0 To Me.clbTipeJournal.Items.Count - 1
                Me.clbTipeJournal.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkJournalType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkJournalType.CheckedChanged
        Try
            If Me.chkJournalType.Checked = True Then
                For i As Integer = 0 To Me.clbTipeJournal.Items.Count - 1
                    Me.clbTipeJournal.SetItemCheckState(i, CheckState.Checked)
                Next
                Me.clbTipeJournal.Enabled = False
                Me.llbCheckAll.Enabled = False
                Me.llbUncheckAll.Enabled = False
            Else
                For i As Integer = 0 To Me.clbTipeJournal.Items.Count - 1
                    Me.clbTipeJournal.SetItemCheckState(i, CheckState.Unchecked)
                Next
                Me.clbTipeJournal.Enabled = True
                Me.llbCheckAll.Enabled = True
                Me.llbUncheckAll.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColAccount.Index Then
                e.Cancel = True
                Exit Sub
            Else
                'col account
                If e.ColumnIndex = Me.ColAccount.Index Then
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColAccount.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                        Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                    End If
                    Exit Sub
                End If
            Else
                ''data account number
                If e.ColumnIndex = Me.ColAccount.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New FrmMasterAkun_Cari
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                            Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                        ElseIf ij.mPrimaryAkun = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJournalDescription.GotFocus, TxtRemarksItem.GotFocus
        Dim Txt As Windows.Forms.TextBox

        Try
            Txt = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    If Txt.ReadOnly = False Then
                        Txt.BackColor = Me.BackCOlorAll
                    Else
                        Txt.BackColor = Me.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJournalDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJournalDescription.LostFocus, TxtRemarksItem.LostFocus
        Dim Txt As Windows.Forms.TextBox

        Try
            Txt = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True Then
                    If Txt.ReadOnly = False Then
                        Txt.BackColor = Color.White
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class