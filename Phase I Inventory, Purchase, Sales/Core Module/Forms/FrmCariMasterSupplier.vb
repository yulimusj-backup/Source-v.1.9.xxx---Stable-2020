Option Explicit On
Option Strict Off

Public Class FrmCariMasterSupplier
    Public PrimarySupplier As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Dim Obj As New clsMasterSupplier
    Dim TBL As New dseMasterSupplier.tblmaster_supplier1DataTable
    Dim Kriteria As Core_Module.ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByCode
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.PrimarySupplier = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.PrimarySupplier = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmCariMasterSupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""

        Try
            Me.ShowIcon = False
            If Me.UserCode = "" Then
                Me.UserCode = My.Settings.user_code
            End If

            Me.txtCari.Focus()
            Dummy = Me.ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.MasterSupplier, _
            ClsTransaksiSettingUser.EnumNamaParameter.MasterSupplier_DefaultFindBy, Me.UserCode)
            If Dummy = "" Or Dummy = CInt(Core_Module.clsMasterSupplier.EnumDefaultFindBy_SettingUser.ByGlobalSetting).ToString Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MasterSupplier, _
                ClsTransaksiSetting.EnumNamaParameter.MasterSupplier_DefaultFindBy)
                If Dummy = "" Then
                    Me.cboCriteria.SelectedIndex = 0
                Else
                    If CInt(Dummy) = CInt(Core_Module.clsMasterSupplier.EnumDefaultFindBy.ByDesc) Then
                        Me.cboCriteria.SelectedIndex = 1
                    Else
                        Me.cboCriteria.SelectedIndex = 0
                    End If
                End If
            Else
                If CInt(Dummy) = CInt(Core_Module.clsMasterSupplier.EnumDefaultFindBy_SettingUser.ByDesc) Then
                    Me.cboCriteria.SelectedIndex = 1
                Else
                    Me.cboCriteria.SelectedIndex = 0
                End If
            End If

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
                Me.cboCriteria.SelectedIndex, Me.tslQuery)
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
            Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, Me.cboCriteria.SelectedIndex, Me.tslQuery)
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
            Me.PrimarySupplier = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml(Me.txtCari.Text, Me.cboCriteria.SelectedIndex) < 150 Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, Me.cboCriteria.SelectedIndex, Me.tslQuery)
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
        txtCari.BackColor = Me.txtPrimaryGroup.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.SelectedIndex = 0 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
            ElseIf Me.cboCriteria.SelectedIndex = 1 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
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