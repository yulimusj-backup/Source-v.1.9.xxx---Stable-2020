Option Explicit On
Option Strict Off

Imports system.drawing, System.Windows.Forms
Public Class FrmTrans_Gaji_Override_Cari_Formula
    Public mPrimary As Double = -1
    Public mPrimaryUser As Double = -1
    Public mIdentifierFormula As String = "-1-1"
    Public TypePencarian As ClsTransaksiSalary_Override.EnumFormulaType = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim Obj As New ClsTransaksiSalary_Override
    Dim TBL_Formula As New DseTransaksiSalary_Override.tblcarimaster_salary_formula_view1DataTable
    Dim TBL_Other As New DseTransaksiSalary_Override.tblcarimaster_salary_formula_view2DataTable
    Dim FindCriteria As Core_Module.EnumSearchCriteria = Core_Module.EnumSearchCriteria.ByComponentName

    Private Sub frmCariMasterKaryawan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cboCriteria.Items.Add("Component Name")
            Me.cboCriteria.Items.Add("Salary Formula")

            Me.txtCari.Select()
            Me.cboCriteria.SelectedIndex = 0
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                If Me.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula Then
                    Obj.IsiDGV_Formula(Me.txtCari.Text, Me.dgvHasil, Me.TBL_Formula, _
                    Me.FindCriteria, Me.tslQuery, Me.mPrimaryUser, Me.dtpSampai.Value, Me.dtpDari.Value)
                Else
                    Obj.IsiDGV_OtherDeductionFormula(Me.txtCari.Text, Me.dgvHasil, Me.TBL_Other, _
                    Me.FindCriteria, Me.tslQuery, Me.dtpSampai.Value, Me.dtpDari.Value)
                End If

                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                End If
            ElseIf e.KeyValue = 40 And Me.dgvHasil.RowCount > 0 Then
                Me.dgvHasil.Select()
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
            If Me.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula Then
                Obj.IsiDGV_Formula(Me.txtCari.Text, Me.dgvHasil, Me.TBL_Formula, Me.FindCriteria, Me.tslQuery, _
                Me.mPrimaryUser, Me.dtpSampai.Value, Me.dtpDari.Value)
            Else
                Obj.IsiDGV_OtherDeductionFormula(Me.txtCari.Text, Me.dgvHasil, Me.TBL_Other, Me.FindCriteria, Me.tslQuery, _
                Me.dtpSampai.Value, Me.dtpDari.Value)
            End If

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
            Me.mPrimary = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula Then
                If Me.Obj.GetJml_Formula(Me.txtCari.Text, Me.FindCriteria, Me.mPrimaryUser, Me.dtpSampai.Value, Me.dtpDari.Value) < 500 Then
                    Obj.IsiDGV_Formula(Me.txtCari.Text, Me.dgvHasil, Me.TBL_Formula, Me.FindCriteria, Me.tslQuery, _
                     Me.mPrimaryUser, Me.dtpSampai.Value, Me.dtpDari.Value)
                    If Me.dgvHasil.Rows.Count > 0 Then
                        Me.dgvHasil.ClearSelection()
                    End If
                End If
            Else
                If Me.Obj.GetJml_OtherDeductionFormula(Me.txtCari.Text, Me.FindCriteria, Me.dtpSampai.Value, Me.dtpDari.Value) < 500 Then
                    Obj.IsiDGV_OtherDeductionFormula(Me.txtCari.Text, Me.dgvHasil, Me.TBL_Other, Me.FindCriteria, Me.tslQuery, _
                    Me.dtpSampai.Value, Me.dtpDari.Value)
                    If Me.dgvHasil.Rows.Count > 0 Then
                        Me.dgvHasil.ClearSelection()
                    End If
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
            If Me.cboCriteria.Text = "Component Name" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = Core_Module.EnumSearchCriteria.ByComponentName
            ElseIf Me.cboCriteria.Text = "Salary Formula" Then
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False
                Me.txtCari.Visible = True
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = Core_Module.EnumSearchCriteria.BySalaryFormula
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Pilih()
        Try
            If Me.TypePencarian = ClsTransaksiSalary_Override.EnumFormulaType.Salary_Formula Then
                If Me.dgvHasil.SelectedRows.Count > 0 Then
                    Me.mPrimary = Me.dgvHasil.SelectedRows(0).Cells(3).Value
                    Me.mIdentifierFormula = Me.dgvHasil.SelectedRows(0).Cells(2).Value
                Else
                    Me.mPrimary = -1
                    Me.mIdentifierFormula = "-1-1"
                End If
            Else
                If Me.dgvHasil.SelectedRows.Count > 0 Then
                    Me.mPrimary = Me.dgvHasil.SelectedRows(0).Cells(0).Value
                    Me.mIdentifierFormula = Me.dgvHasil.SelectedRows(0).Cells(0).Value
                Else
                    Me.mPrimary = -1
                    Me.mIdentifierFormula = "-1-1"
                End If
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