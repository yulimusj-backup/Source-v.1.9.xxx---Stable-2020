Option Explicit On
Option Strict Off

Public Class FrmTransaksiBalanceManual_DP_Cari
    Public mPrimaryDept As Double = -1
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vTipeDP As ClsTransaksiBalanceManual_DP.EnumTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP
    Public TBLAop_DP As New DseTransaksiBalanceManual_DP.tbl_aop_dp_findDataTable
    Public TBLInv_DP As New DseTransaksiBalanceManual_DP.tbl_inv_dp_findDataTable

    Dim FindCriteria As EnumSearchCriteria = EnumSearchCriteria.ByCustomerCode
    Dim Obj As New ClsTransaksiBalanceManual_DP

    Dim Criteria_1 As String = ""
    Dim Criteria_2 As String = ""
    Dim Criteria_3 As String = ""
    Dim Criteria_4 As String = ""
    Dim Criteria_5 As String = ""

    Private Sub frmCariMasterCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim AddCriteria_4 As Boolean = True

        Try
            If vTipeDP = ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP Then
                Me.Text = "Find Beginning Balance Purchase Down Payment Data"
                Me.Criteria_1 = "Supplier Code"
                Me.Criteria_2 = "Supplier Name"
                Me.Criteria_3 = "AOP DP No"
                Me.Criteria_4 = "Supplier Invoice No"
                Me.Criteria_5 = "PO No"
            Else
                Me.Text = "Find Beginning Balance Sales Down Payment Data"
                Me.Criteria_1 = "Customer Code"
                Me.Criteria_2 = "Customer Name"
                Me.Criteria_3 = "Invoice DP No"
                Me.Criteria_5 = "SO No"
                AddCriteria_4 = False
            End If

            Me.cboCriteria.Items.Add(Criteria_1)
            Me.cboCriteria.Items.Add(Criteria_2)
            Me.cboCriteria.Items.Add(Criteria_3)
            If AddCriteria_4 = True Then
                Me.cboCriteria.Items.Add(Criteria_4)
            End If
            Me.cboCriteria.Items.Add(Criteria_5)

            Me.txtCari.Focus()
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
               Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBLAop_DP, Me.TBLInv_DP, Me.FindCriteria, Me.tslQuery, Me.vTipeDP)
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
            Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBLAop_DP, Me.TBLInv_DP, Me.FindCriteria, Me.tslQuery, Me.vTipeDP)
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
            Me.mPrimaryDept = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml(Me.txtCari.Text, Me.FindCriteria, Me.vTipeDP) < 500 Then
                Obj.IsiDGV(Me.txtCari.Text, Me.dgvHasil, Me.TBLAop_DP, Me.TBLInv_DP, Me.FindCriteria, Me.tslQuery, Me.vTipeDP)
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
            If Me.cboCriteria.Text = Me.Criteria_1 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByCustomerCode
            ElseIf Me.cboCriteria.Text = Me.Criteria_2 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.FindCriteria = EnumSearchCriteria.ByCustomerName
            ElseIf Me.cboCriteria.Text = Me.Criteria_3 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.ByInvoiceNumber
            ElseIf Me.cboCriteria.Text = Me.Criteria_4 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.BySupplierInvoiceNo
            ElseIf Me.cboCriteria.Text = Me.Criteria_5 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.FindCriteria = EnumSearchCriteria.BySalesOrderNo
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Pilih()
        Try
            If Me.dgvHasil.SelectedRows.Count > 0 Then
                Me.mPrimaryDept = Me.dgvHasil.SelectedRows(0).Cells(0).Value
            Else
                Me.mPrimaryDept = -1
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