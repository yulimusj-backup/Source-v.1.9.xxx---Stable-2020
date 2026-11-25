Option Explicit On
Option Strict Off

Public Class FrmTrans_SchDownTime_Cari
    Public backColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public primary As Double = -1

    Dim obj As New ClsTrans_SchDownTime
    Dim tbl As New DseTrans_SchDownTime.tbltrans_sch_downtime1DataTable
    Dim criteria As ClsTrans_SchDownTime.EnumSearchBy = ClsTrans_SchDownTime.EnumSearchBy.DOCUMENT_NO

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Primary = CDbl(Me.dgvHasil.SelectedRows(0).Cells(0).Value)
            Else
                Me.Primary = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_SchDownTime_Cari_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.tbl, criteria, Me.tslQuery, dtpDari.Value, dtpSampai.Value)
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

    Private Sub btnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.tbl, criteria, Me.tslQuery, dtpDari.Value, dtpSampai.Value)
            If Me.dgvHasil.Rows.Count > 0 Then
                Me.dgvHasil.ClearSelection()
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
            If Me.obj.GetJml(Me.txtCari.Text, Me.criteria, dtpDari.Value, dtpSampai.Value) < 150 Then
                obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.tbl, criteria, Me.tslQuery, dtpDari.Value, dtpSampai.Value)
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
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
        txtCari.BackColor = Me.txtPrimarySPL.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.SelectedIndex = 0 Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.criteria = ClsTrans_SchDownTime.EnumSearchBy.DOCUMENT_NO
            ElseIf Me.cboCriteria.SelectedIndex = 1 Then
                Me.LblFrom.Visible = True
                Me.lblTo.Visible = True
                Me.dtpDari.Visible = True
                Me.dtpSampai.Visible = True
                Me.txtCari.Visible = False
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.criteria = ClsTrans_SchDownTime.EnumSearchBy.DATE_
            ElseIf Me.cboCriteria.SelectedIndex = 2 Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.criteria = ClsTrans_SchDownTime.EnumSearchBy.REASON
            ElseIf Me.cboCriteria.SelectedIndex = 3 Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.criteria = ClsTrans_SchDownTime.EnumSearchBy.REMARKS
            ElseIf Me.cboCriteria.SelectedIndex = 4 Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.criteria = ClsTrans_SchDownTime.EnumSearchBy.WORK_UNIT_CODE
            ElseIf Me.cboCriteria.SelectedIndex = 5 Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.criteria = ClsTrans_SchDownTime.EnumSearchBy.WORK_UNIT_DESC
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