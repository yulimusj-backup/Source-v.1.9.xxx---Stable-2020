Option Explicit On
Option Strict Off

Public Class FrmTransaksiJournal_CariPurchase
    Public PrimaryPP As Double = -1
    Public BackColorAll As Color = Color.LightSeaGreen
    Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
    Public TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping

    Dim Obj As New ClsTransaksiJournal
    Dim TBL As New DseTransaksiJournal.tbltrans_jurnal1DataTable
    Dim FindCriteria As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByReferenceCode
    Dim FilterCreatedBy As String
    Dim UserCode As String = ""
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.PrimaryPP = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.PrimaryPP = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.PrimaryPP = -1
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
            Me.cboCriteria.Items.Add("No Journal")
            Me.cboCriteria.Items.Add("Date")
            Me.cboCriteria.Items.Add("Description")
            Me.cboCriteria.Items.Add("Supplier Code")
            Me.cboCriteria.Items.Add("Supplier Name")
            Me.cboCriteria.Items.Add("AOP Number")
            Me.cboCriteria.Items.Add("Supplier Invoice Number")
            Me.cboCriteria.Items.Add("Purchase Return Number")
            Me.cboCriteria.SelectedIndex = 0

            Me.txtCari.Focus()
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If

            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.UserCode = Core_Module.GetUser
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Jurnal_IsLedgerAccessActive)
            If Dummy <> "" And Dummy <> "0" Then
                TipeAkses = Me.Obj.GetLedgerAccess(Me.UserCode)
                If TipeAkses = ClsTransaksiJournal.EnumJournalAccesType.Respected_User_Only Then
                    Me.FilterCreatedBy = "[Created By] = '" + Me.UserCode + "'"
                ElseIf TipeAkses = ClsTransaksiJournal.EnumJournalAccesType.Selected_Users_Only Then
                    Me.Obj.GetLedgerAccess(Me.UserCode, ListCreatedBy)
                    For Each UserCode In ListCreatedBy
                        If Me.FilterCreatedBy = "" Then
                            Me.FilterCreatedBy = "[Created By] = '" + UserCode + "'"
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
                Obj.IsiDGV_PurchasePayment(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.TipeJournal, Me.dtpDari.Value, Me.dtpSampai.Value, Me.TipeJurnal2, Me.FilterCreatedBy)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    Me.dgvHasil.Rows(0).Selected = True
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
            Obj.IsiDGV_PurchasePayment(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.TipeJournal, Me.dtpDari.Value, Me.dtpSampai.Value, Me.TipeJurnal2, Me.FilterCreatedBy)
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
                Me.dgvHasil.Rows(0).Selected = True
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

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml(Me.txtCari.Text, Me.FindCriteria, Me.TipeJournal, Me.dtpDari.Value, _
            Me.dtpSampai.Value, Me.TipeJurnal2) < 500 Then
                Obj.IsiDGV_PurchasePayment(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.TipeJournal, Me.dtpDari.Value, Me.dtpSampai.Value, Me.TipeJurnal2, Me.FilterCreatedBy)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                    Me.dgvHasil.Rows(0).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCari_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.GotFocus
        txtCari.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCari_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.LostFocus
        txtCari.BackColor = Me.txtPrimaryGroup.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.Text = "No Journal" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByReferenceCode
            ElseIf Me.cboCriteria.Text = "Date" Then
                Me.LblFrom.Visible = True
                Me.lblTo.Visible = True
                Me.dtpDari.Visible = True
                Me.dtpSampai.Visible = True
                Me.txtCari.Visible = False
                Me.FindCriteria = EnumSearchCriteria.ByDate
            ElseIf Me.cboCriteria.Text = "Description" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.ByDescription
            ElseIf Me.cboCriteria.Text = "Supplier Name" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.BySupplierName
            ElseIf Me.cboCriteria.Text = "Supplier Code" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.BySupplierCode
            ElseIf Me.cboCriteria.Text = "AOP Number" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByAOPNumber
            ElseIf Me.cboCriteria.Text = "Supplier Invoice Number" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.BySupplierInvoiceNo
            ElseIf Me.cboCriteria.Text = "Purchase Return Number" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByPurchaseReturnNo
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpSampai_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpSampai.KeyDown
        If e.KeyValue = 13 Then
            Me.btnFind.PerformClick()
        End If
    End Sub

    Private Sub dtpDari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDari.KeyDown
        If e.KeyValue = 13 Then
            Me.btnFind.PerformClick()
        End If
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