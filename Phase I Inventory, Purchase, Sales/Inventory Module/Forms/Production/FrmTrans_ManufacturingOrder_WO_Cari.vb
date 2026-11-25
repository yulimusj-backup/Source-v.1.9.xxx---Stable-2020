Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_WO_Cari
    Public Identifier As String = ""
    Public primaryMO As Double = -1
    Public PrimaryBundling As Double = -1
    Public IdentifierJO As String = ""

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vStatus1 As New ClsTrans_ManufacturingOrder.EnumPilihStatusWO
    Public vStatus2 As New ClsTrans_ManufacturingOrder.EnumPilihStatusWO
    Public vStatus3 As New ClsTrans_ManufacturingOrder.EnumPilihStatusWO
    Dim Obj As New ClsTrans_StockBundling
    Dim TBL As New DseTrans_StockBundling.tbltrans_mo_item_wo_findDataTable
    Dim FindCriteria As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByNoTransaksi

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.Identifier = Me.dgvHasil.SelectedRows(0).Cells(1).Value
                Me.IdentifierJO = Me.dgvHasil.SelectedRows(0).Cells(8).Value
            Else
                Me.primaryMO = -1
                Me.Identifier = ""
                Me.IdentifierJO = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmCariMasterCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
                Obj.IsiDgvByWo(PrimaryBundling, Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                FindCriteria, Me.tslQuery, vStatus1, vStatus2, vStatus3, Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date)
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
            Obj.IsiDgvByWo(PrimaryBundling, Me.txtCari.Text, Me.dgvHasil, Me.TBL, FindCriteria, Me.tslQuery, vStatus1, vStatus2, vStatus3, _
            Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date)

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
            Me.primaryMO = -1
            Me.Identifier = ""
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJmlByWo(Me.txtCari.Text, Me.cboCriteria.SelectedIndex, vStatus1, vStatus2, vStatus3, Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date) < 150 Then
                Obj.IsiDgvByWo(PrimaryBundling, Me.txtCari.Text, Me.dgvHasil, Me.TBL, FindCriteria, Me.tslQuery, vStatus1, vStatus2, vStatus3, Me.dtpDari.Value.Date, Me.dtpSampai.Value.Date)
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
        txtCari.BackColor = Me.txtPrimaryGroup.BackColor
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
                Me.FindCriteria = EnumSearchCriteria.ByNoTransaksi
            ElseIf Me.cboCriteria.SelectedIndex = 1 Then
                Me.LblFrom.Visible = True
                Me.lblTo.Visible = True
                Me.dtpDari.Visible = True
                Me.dtpSampai.Visible = True
                Me.txtCari.Visible = False
                Me.FindCriteria = EnumSearchCriteria.ByTanggalTransaksi
            ElseIf Me.cboCriteria.SelectedIndex = 2 Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.ByRemarks
            Else
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByJONumber
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