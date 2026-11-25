Option Strict Off
Option Explicit On
Public Class FrmMasterProductAccountGroup_Cari
    Public Primary As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim Obj As New ClsMasterProductAccountGroup
    Dim TBL As New DseMasterProductAccountGroup.DataTable1DataTable
    Dim Kriteria As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByCode
    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Primary = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.Primary = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.Kriteria, Me.tslQuery)
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
            Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, Me.Kriteria, Me.tslQuery)
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
            Me.Primary = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml(Me.txtCari.Text, Me.Kriteria) < 150 Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, Me.Kriteria, Me.tslQuery)
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
            Me.txtCari.Left = Me.CboGroup.Left
            Me.txtCari.Width = Me.btnFind.Left - 5 - Me.txtCari.Left

            If Me.cboCriteria.SelectedIndex = 0 Then
                Me.CboGroup.Visible = False
                Me.CboGroup.SelectedIndex = -1
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.Kriteria = EnumSearchCriteria.ByCode
            ElseIf Me.cboCriteria.SelectedIndex = 1 Then
                Me.CboGroup.SelectedIndex = -1
                Me.CboGroup.Visible = False
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.Kriteria = EnumSearchCriteria.ByDescription
            ElseIf Me.cboCriteria.SelectedIndex = 2 Then
                Me.txtCari.Left = Me.CboGroup.Left + Me.CboGroup.Width + 5
                Me.txtCari.Width = Me.btnFind.Left - 5 - Me.txtCari.Left
                Me.CboGroup.Visible = True
                Me.CboGroup.SelectedIndex = 0
                If Me.CboGroup.SelectedIndex = 0 Then
                    Me.txtCari.CharacterCasing = CharacterCasing.Upper
                    Me.Kriteria = EnumSearchCriteria.ByGroupProductGroup
                ElseIf Me.CboGroup.SelectedIndex = 1 Then
                    Me.txtCari.CharacterCasing = CharacterCasing.Upper
                    Me.Kriteria = EnumSearchCriteria.ByGroupBrand
                ElseIf Me.CboGroup.SelectedIndex = 2 Then
                    Me.txtCari.CharacterCasing = CharacterCasing.Upper
                    Me.Kriteria = EnumSearchCriteria.ByGroupSubBrand
                ElseIf Me.CboGroup.SelectedIndex = 3 Then
                    Me.txtCari.CharacterCasing = CharacterCasing.Upper
                    Me.Kriteria = EnumSearchCriteria.ByGroupCategory
                End If
            ElseIf Me.cboCriteria.SelectedIndex = 3 Then
                Me.CboGroup.SelectedIndex = -1
                Me.CboGroup.Visible = False
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.Kriteria = EnumSearchCriteria.ByRemarks
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

    Private Sub FrmMasterCorrFact_Cari_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ShowIcon = False
            Me.txtCari.Select()
            Me.cboCriteria.SelectedIndex = 0
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'Me.StartPosition = FormStartPosition.CenterScreen
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboGroup.SelectedIndexChanged
        If Me.cboCriteria.SelectedIndex = 2 Then
            Me.CboGroup.Visible = True
            'Me.txtCari.Location = New System.Drawing.Point(260, 12)
            'Me.txtCari.Size = New System.Drawing.Size(252, 20)

            If Me.CboGroup.SelectedIndex = 0 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.Kriteria = EnumSearchCriteria.ByGroupProductGroup
            ElseIf Me.CboGroup.SelectedIndex = 1 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.Kriteria = EnumSearchCriteria.ByGroupBrand
            ElseIf Me.CboGroup.SelectedIndex = 2 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.Kriteria = EnumSearchCriteria.ByGroupSubBrand
            ElseIf Me.CboGroup.SelectedIndex = 3 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.Kriteria = EnumSearchCriteria.ByGroupCategory
            End If
        End If
    End Sub
End Class