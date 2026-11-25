''@rosadinnn
''2015-04-15

Option Explicit On
Option Strict Off
Public Class FrmEmailSettings_Cari
    Public primaryAccount As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim obj As New ClsEmailSettings
    Dim criteria As ClsEmailSettings.EnumEmailSearchCriteria = ClsEmailSettings.EnumEmailSearchCriteria.ByAccountCode
    Dim tbl As New DseEmailSettings.DataTable1DataTable
    Public UserCode As String = ""
    Public IsAccountWithUserCode As Boolean = False
    Public Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.primaryAccount = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.primaryAccount = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmEmailSettings_Cari_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.txtCari.Focus()
            Me.cboCriteria.SelectedIndex = 0
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.GotFocus
        Try
            txtCari.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                obj.CariDetail2(Me.txtCari.Text, Me.dgvHasil, Me.tbl, Me.criteria, Me.tslQuery, Me.IsAccountWithUserCode, Me.UserCode)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
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

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            obj.CariDetail2(Me.txtCari.Text, Me.dgvHasil, Me.tbl, Me.criteria, Me.tslQuery, Me.IsAccountWithUserCode, Me.UserCode)
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_CellContentDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHasil.CellContentDoubleClick
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

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.primaryAccount = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCari.LostFocus
        Try
            txtCari.BackColor = Me.txtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            'If obj.GetJml(Me.txtCari.Text, Me.criteria) < 150 Then
            obj.CariDetail2(Me.txtCari.Text, Me.dgvHasil, Me.tbl, Me.criteria, Me.tslQuery, Me.IsAccountWithUserCode, Me.UserCode)
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.SelectedIndex = 0 Or Me.cboCriteria.SelectedIndex = 1 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
            ElseIf Me.cboCriteria.SelectedIndex = 2 Or Me.cboCriteria.SelectedIndex = 3 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
            End If
            If Me.cboCriteria.SelectedIndex = 0 Then
                Me.criteria = ClsEmailSettings.EnumEmailSearchCriteria.ByAccountCode
            ElseIf Me.cboCriteria.SelectedIndex = 1 Then
                Me.criteria = ClsEmailSettings.EnumEmailSearchCriteria.ByUserCode
            ElseIf Me.cboCriteria.SelectedIndex = 2 Then
                Me.criteria = ClsEmailSettings.EnumEmailSearchCriteria.ByEmailFrom
            ElseIf Me.cboCriteria.SelectedIndex = 3 Then
                Me.criteria = ClsEmailSettings.EnumEmailSearchCriteria.ByDescription
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvHasil.MouseDoubleClick
        Try
            Me.Pilih()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class