Public Class FrmTransaksiRecurring_Cari
    Public PrimaryPP As Double = -1
    Public BackColorAll As Color = Color.LightSeaGreen
    'Public TipeJournal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.AdjustmenJournal
    'Public TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping

    Dim Obj As New ClsTransaksiRecurring
    Dim TBL As New DseTransaksiRecurring.tblCariDataTable
    Dim FindCriteria As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByReferenceCode

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
        Try
            Me.ShowIcon = False
            Me.cboCriteria.Items.Clear()
            Me.cboCriteria.Items.Add("No Journal")
            Me.cboCriteria.Items.Add("Date")
            Me.cboCriteria.Items.Add("Description")
            'If Me.TipeJournal = ClsTransaksiJournal.EnumTipeJournal.PurchasePaymentJournal Then
            '    Me.cboCriteria.Items.Add("Supplier Code")
            '    Me.cboCriteria.Items.Add("Supplier Name")
            '    Me.cboCriteria.Items.Add("AOP Number")
            '    Me.cboCriteria.Items.Add("Supplier Invoice Number")
            '    Me.cboCriteria.Items.Add("Purchase Return Number")
            'Else
            '    Me.cboCriteria.Items.Add("Customer Code")
            '    Me.cboCriteria.Items.Add("Customer Name")
            '    Me.cboCriteria.Items.Add("Invoice Number")
            '    Me.cboCriteria.Items.Add("Sales Return Number")
            'End If
            Me.cboCriteria.SelectedIndex = 0
            Me.TBL.Rows.Clear()
            Me.txtCari.Focus()
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If

            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.dtpDari.Value, Me.dtpSampai.Value)
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

    Public Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.dtpDari.Value, Me.dtpSampai.Value)
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
            If Me.Obj.GetJml(Me.txtCari.Text, Me.FindCriteria, Me.dtpDari.Value, _
            Me.dtpSampai.Value) < 500 Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.dtpDari.Value, Me.dtpSampai.Value)
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