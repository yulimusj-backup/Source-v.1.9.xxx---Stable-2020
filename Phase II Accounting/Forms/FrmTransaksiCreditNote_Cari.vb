Option Explicit On
Option Strict Off
Public Class FrmTransaksiCreditNote_cari
    Public PrimaryTransCN As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim Obj As New ClsTransaksiCreditNote
    Dim ObjJournal As New ClsTransaksiJournal
    Dim objFTransCN As New FrmTransaksiCreditNote
    Dim objTransCN As New ClsTransaksiCreditNote
    Dim TBL As New DseTransaksiCreditNote_TransferJournal.tbltrans_credit_note_viewDataTable
    Dim Kriteria As Core_Module.ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByNoTransaksi
    Public TipeTransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi
    Dim UserCode As String = ""
    Dim FilterCreatedBy As String = ""
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.PrimaryTransCN = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.PrimaryTransCN = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmCariMasterCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ListCreatedBy As New List(Of String)
        Dim TipeAkses As ClsTransaksiJournal.EnumJournalAccesType = ClsTransaksiJournal.EnumJournalAccesType.FullAccess
        Dim Dummy As String = ""

        Try
            Me.ShowIcon = False
            Me.txtCari.Select()
            Me.cboCriteria.Items.Add("Document No")
            Me.cboCriteria.Items.Add("Date")
            Me.cboCriteria.Items.Add("No Jurnal")
            Me.cboCriteria.Items.Add("Description Jurnal")
            Me.cboCriteria.Items.Add("Reference Party Code")
            Me.cboCriteria.Items.Add("Reason")

            If Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE Then
                'Me.cboCriteria.Items.Add("Reference Party Name")
                Me.Text = "Find Credit Note"
            ElseIf Me.TipeTransaksi = ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE Then
                'Me.cboCriteria.Items.Add("Reference Party Code")
                'Me.cboCriteria.Items.Add("Reference Party Name")
                Me.Text = "Find Debit Note"
            End If
            Me.cboCriteria.SelectedIndex = 0
            Me.txtCari.Focus()
            Me.cboCriteria.SelectedIndex = 0
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.UserCode = Core_Module.GetUser
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive)
            If Dummy <> "" And Dummy <> "0" Then
                TipeAkses = Me.ObjJournal.GetLedgerAccess(Me.UserCode)
                If TipeAkses = ClsTransaksiJournal.EnumJournalAccesType.Respected_User_Only Then
                    If Me.UserCode <> "" Then
                        Me.FilterCreatedBy = "[Created By] = '" + Me.UserCode + "'"
                    End If

                ElseIf TipeAkses = ClsTransaksiJournal.EnumJournalAccesType.Selected_Users_Only Then
                    Me.ObjJournal.GetLedgerAccess(Me.UserCode, ListCreatedBy)
                    For Each UserCode In ListCreatedBy
                        If Me.FilterCreatedBy = "" Then
                            If Me.UserCode <> "" Then
                                Me.FilterCreatedBy = "[Created By] = '" + UserCode + "'"
                            End If
                        Else
                            Me.FilterCreatedBy = Me.FilterCreatedBy + " or [Created By] = '" + UserCode + "'"
                        End If
                    Next
                ElseIf TipeAkses = ClsTransaksiJournal.EnumJournalAccesType.FullAccess Then
                    Me.FilterCreatedBy = ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListCreatedBy = Nothing
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.Kriteria, Me.tslQuery, TipeTransaksi, Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date, Me.FilterCreatedBy)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    '
                End If
            ElseIf e.KeyValue = 40 And Me.dgvHasil.RowCount > 0 Then
                Me.dgvHasil.Focus()
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.Rows.Item(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, Me.Kriteria, Me.tslQuery, TipeTransaksi, Me.dtpDari.Value, dtpSampai.Value, Me.FilterCreatedBy)
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
                '
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHasil.CellContentDoubleClick
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvHasil.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.Pilih()
                Me.Close()
            ElseIf e.KeyValue = 38 Then
                If Me.dgvHasil.SelectedRows.Count > 0 And Me.dgvHasil.SelectedRows(0).Index = 0 Then
                    Me.txtCari.Focus()
                End If
            ElseIf e.KeyValue = 9 Then
                Me.txtCari.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.PrimaryTransCN = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.Get_Jml(Me.txtCari.Text, Me.Kriteria) < 150 Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, Me.Kriteria, Me.tslQuery, TipeTransaksi, Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date, Me.FilterCreatedBy)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    '
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtCari_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.GotFocus
        txtCari.BackColor = BackColorAll
    End Sub
    Private Sub TxtCari_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.LostFocus
        txtCari.BackColor = Me.txtprimarygroup.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.SelectedItem = "No Jurnal" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.Kriteria = EnumSearchCriteria.ByAccountNo
            ElseIf Me.cboCriteria.SelectedItem = "Date" Then
                Me.LblFrom.Visible = True
                Me.lblTo.Visible = True
                Me.dtpDari.Value = Now
                Me.dtpSampai.Value = Now
                Me.dtpDari.Visible = True
                Me.dtpSampai.Visible = True
                Me.txtCari.Visible = False
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.Kriteria = EnumSearchCriteria.ByDate
            ElseIf Me.cboCriteria.SelectedItem = "Description Jurnal" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.Kriteria = EnumSearchCriteria.ByDescJurnal
            ElseIf Me.cboCriteria.SelectedItem = "Document No" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.Kriteria = EnumSearchCriteria.ByNoTransaksi
                'ElseIf Me.cboCriteria.SelectedIndex = 4 Then
                '    Me.LblFrom.Visible = False
                '    Me.lblTo.Visible = False
                '    Me.dtpDari.Visible = False
                '    Me.dtpSampai.Visible = False
                '    Me.txtCari.Visible = True
                '    Me.txtCari.CharacterCasing = CharacterCasing.Normal
                '    Me.Kriteria = EnumSearchCriteria.ByReferenceName
            ElseIf Me.cboCriteria.SelectedItem = "Reason" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.Kriteria = EnumSearchCriteria.ByRemarks
            ElseIf Me.cboCriteria.SelectedItem = "Reference Party Code" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.Kriteria = EnumSearchCriteria.ByReferenceCode
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub dgvHasil_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvHasil.MouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub dgvHasil_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvHasil.RowHeaderMouseDoubleClick
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class