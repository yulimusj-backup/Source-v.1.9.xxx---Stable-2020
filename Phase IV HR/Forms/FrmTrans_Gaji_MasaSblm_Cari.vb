Option Explicit On
Option Strict Off

Imports System.Drawing
Public Class FrmTrans_Gaji_MasaSblm_Cari
    Public Frm As New FrmTrans_Gaji_MasaSblm
    Public PrimaryTransaksi As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim Obj As New ClsTrans_Gaji_MasaSblm
    Dim TBL As New DseTrans_Gaji_MasaSblm.tbltrans_salary_prev_income_cariDataTable
    Dim FindCriteria As ClsTrans_Gaji_MasaSblm.EnumSearch = ClsTrans_Gaji_MasaSblm.EnumSearch.By_EmployeeCode

    Dim EmpCode As String = "Employee Code"
    Dim EmpName As String = "Employee Name"
    Dim NoTrans As String = "A1 Form No"
    Dim TglTrans As String = "Date"
    Dim Employer As String = "Employer"
    Dim Remarks As String = "Remarks"

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.PrimaryTransaksi = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.PrimaryTransaksi = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterCuti_Cari_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Me.txtCari.Focus()
            Me.cboCriteria.Items.Add(Me.EmpCode)
            Me.cboCriteria.Items.Add(Me.EmpName)
            Me.cboCriteria.Items.Add(Me.NoTrans)
            Me.cboCriteria.Items.Add(Me.TglTrans)
            Me.cboCriteria.Items.Add(Me.Employer)
            Me.cboCriteria.Items.Add(Me.Remarks)
            Me.cboCriteria.SelectedIndex = 0

            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If
            Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CancelButton = Me.btnClose
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            'Me.DataTable1TableAdapter.Fill(Me.DseTransaksiLoan.DataTable1)
            If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.dtpDari.Value, Me.dtpSampai.Value, Me.Frm)
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
            Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
            Me.FindCriteria, Me.tslQuery, Me.dtpDari.Value, Me.dtpSampai.Value, Me.Frm)
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
            Me.PrimaryTransaksi = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml(Me.txtCari.Text, Me.dgvHasil, Me.FindCriteria, Me.dtpDari.Value, Me.dtpSampai.Value) < 150 Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.FindCriteria, Me.tslQuery, Me.dtpDari.Value, Me.dtpSampai.Value, Me.Frm)
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
        txtCari.BackColor = Me.txtprimarygroup.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.Text = Me.TglTrans Then
                Me.txtCari.Visible = False
                Me.LblFrom.Visible = True
                Me.lblTo.Visible = True
                Me.dtpDari.Visible = True
                Me.dtpSampai.Visible = True

                Me.FindCriteria = ClsTrans_Gaji_MasaSblm.EnumSearch.By_Date
            Else
                Me.txtCari.Visible = True
                Me.LblFrom.Visible = False
                Me.lblTo.Visible = False
                Me.dtpDari.Visible = False
                Me.dtpSampai.Visible = False

                If Me.cboCriteria.Text = Me.EmpCode Then
                    Me.FindCriteria = ClsTrans_Gaji_MasaSblm.EnumSearch.By_EmployeeCode
                ElseIf Me.cboCriteria.Text = Me.EmpName Then
                    Me.FindCriteria = ClsTrans_Gaji_MasaSblm.EnumSearch.By_EmployeeName
                ElseIf Me.cboCriteria.Text = Me.Employer Then
                    Me.FindCriteria = ClsTrans_Gaji_MasaSblm.EnumSearch.By_Employer
                ElseIf Me.cboCriteria.Text = Me.NoTrans Then
                    Me.FindCriteria = ClsTrans_Gaji_MasaSblm.EnumSearch.By_NoTransaksi
                ElseIf Me.cboCriteria.Text = Me.Remarks Then
                    Me.FindCriteria = ClsTrans_Gaji_MasaSblm.EnumSearch.By_Remarks
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