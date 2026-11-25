Option Explicit On
Option Strict Off

Public Class frmCariMasterProduct
    Public PrimaryProduct As Double = -1
    Public LocationControlledOnly As Boolean = False
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public UntukPOS As Boolean = False

    Dim Obj As New clsMasterProduct
    Dim TBL As New dseMasterProduct.tblmaster_product_view0DataTable
    Dim TBLPOS As New dseMasterProduct.tblmaster_product_view01DataTable
    Dim KriteriaCari As ModGlobalVariable.EnumSearchCriteria = EnumSearchCriteria.ByCode
    Dim DTA As New dseMasterProductTableAdapters.tblmaster_product_view0TableAdapter
    Dim DTAPOS As New dseMasterProductTableAdapters.tblmaster_product_view01TableAdapter

    Private Sub Pilih()
        Try
            If UntukPOS = False Then
                If Me.dgvHasil.SelectedRows.Count > 0 Then
                    Me.PrimaryProduct = Me.dgvHasil.SelectedRows(0).Cells(0).Value
                Else
                    Me.PrimaryProduct = -1
                End If
            Else
                If Me.DGVHasil1.SelectedRows.Count > 0 Then
                    Me.PrimaryProduct = Me.DGVHasil1.SelectedRows(0).Cells(0).Value
                Else
                    Me.PrimaryProduct = -1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmCariMasterCustomer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'TODO: This line of code loads data into the 'DseMasterProduct.tblmaster_product_view0' table. You can move, or remove it, as needed.
        Try
            Me.txtCari.Focus()
            If Me.UntukPOS = False Then
                Me.cboCriteria.SelectedIndex = 0
                Me.KriteriaCari = EnumSearchCriteria.ByCode
            Else
                Me.cboCriteria.SelectedIndex = 3
                Me.KriteriaCari = EnumSearchCriteria.ByBarcode
            End If
            If Me.txtCari.Text <> "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.UntukPOS = False Then
                Me.dgvHasil.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Else
                Me.DGVHasil1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCari.KeyDown
        Try
            If Me.UntukPOS = False Then
                If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                    If Me.LocationControlledOnly = False Then
                        Me.IsiDatabase()
                    Else
                        Obj.IsiDGVLocationControlledOnly(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                        Me.KriteriaCari, Me.tslQuery)
                    End If
                    If Me.dgvHasil.Rows.Count > 0 Then
                        Me.dgvHasil.ClearSelection()
                    End If
                ElseIf e.KeyValue = 40 And Me.dgvHasil.RowCount > 0 Then
                    Me.dgvHasil.Focus()
                    If Me.dgvHasil.Rows.Count > 0 Then
                        Me.dgvHasil.Rows.Item(0).Selected = True
                    End If
                End If
            Else
                If e.KeyValue = 13 And Me.txtCari.Text <> "" Then
                    If Me.LocationControlledOnly = False Then
                        Me.IsiDatabase()
                    Else
                        Obj.IsiDGVLocationControlledOnly(Me.txtCari.Text, Me.DGVHasil1, Me.TBL, _
                        Me.KriteriaCari, Me.tslQuery)
                    End If
                    If Me.DGVHasil1.Rows.Count > 0 Then
                        Me.DGVHasil1.ClearSelection()
                    End If
                ElseIf e.KeyValue = 40 And Me.DGVHasil1.RowCount > 0 Then
                    Me.DGVHasil1.Focus()
                    If Me.DGVHasil1.Rows.Count > 0 Then
                        Me.DGVHasil1.Rows.Item(0).Selected = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCari_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Try
            If Me.LocationControlledOnly = False Then
                Me.IsiDatabase()
            Else
                Obj.IsiDGVLocationControlledOnly(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                Me.KriteriaCari, Me.tslQuery)
            End If
            If Me.UntukPOS = False Then
                If Me.dgvHasil.Rows.Count > 0 Then
                    Me.dgvHasil.ClearSelection()
                End If
            Else
                If Me.DGVHasil1.Rows.Count > 0 Then
                    Me.DGVHasil1.ClearSelection()
                End If
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
                If Me.UntukPOS = False Then
                    If Me.dgvHasil.SelectedRows.Count > 0 And Me.dgvHasil.SelectedRows(0).Index = 0 Then
                        Me.txtCari.Focus()
                    End If
                Else
                    If Me.DGVHasil1.SelectedRows.Count > 0 And Me.DGVHasil1.SelectedRows(0).Index = 0 Then
                        Me.txtCari.Focus()
                    End If
                End If
                ElseIf e.KeyValue = 9 Then
                    Me.txtCari.Focus()
                End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DGVHasil1.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.Pilih()
                Me.Close()
            ElseIf e.KeyValue = 38 Then
                If Me.UntukPOS = False Then
                    If Me.dgvHasil.SelectedRows.Count > 0 And Me.dgvHasil.SelectedRows(0).Index = 0 Then
                        Me.txtCari.Focus()
                    End If
                Else
                    If Me.DGVHasil1.SelectedRows.Count > 0 And Me.DGVHasil1.SelectedRows(0).Index = 0 Then
                        Me.txtCari.Focus()
                    End If
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
            Me.PrimaryProduct = -1
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCari.TextChanged
        Try
            If Me.Obj.GetJml(Me.txtCari.Text, Me.KriteriaCari) < 150 Then
                If Me.LocationControlledOnly = False Then
                    Me.IsiDatabase()
                Else
                    Obj.IsiDGVLocationControlledOnly(Me.txtCari.Text, Me.dgvHasil, Me.TBL, _
                    Me.KriteriaCari, Me.tslQuery)
                End If
                If Me.UntukPOS = False Then
                    If Me.dgvHasil.Rows.Count > 0 Then
                        Me.dgvHasil.ClearSelection()
                        '
                    End If
                Else
                    If Me.DGVHasil1.Rows.Count > 0 Then
                        Me.DGVHasil1.ClearSelection()
                        '
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
        txtCari.BackColor = Me.txtprimarygroup.BackColor
    End Sub

    Private Sub cboCriteria_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCriteria.SelectedIndexChanged
        Try
            If Me.cboCriteria.SelectedIndex = 0 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Upper
                Me.KriteriaCari = EnumSearchCriteria.ByCode
            ElseIf Me.cboCriteria.SelectedIndex = 1 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.KriteriaCari = EnumSearchCriteria.ByDescription
            ElseIf Me.cboCriteria.SelectedIndex = 2 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.KriteriaCari = EnumSearchCriteria.ByProductSpec
            ElseIf Me.cboCriteria.SelectedIndex = 3 Then
                Me.txtCari.CharacterCasing = CharacterCasing.Normal
                Me.KriteriaCari = EnumSearchCriteria.ByBarcode
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvHasil.MouseDoubleClick
        Try
            If UntukPOS = False Then
                If Me.dgvHasil.SelectedRows.Count > 0 Then
                    Me.Pilih()
                    Me.Close()
                End If
            Else
                If Me.DGVHasil1.SelectedRows.Count > 0 Then
                    Me.Pilih()
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHasil1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DGVHasil1.MouseDoubleClick
        Try
            If UntukPOS = False Then
                If Me.dgvHasil.SelectedRows.Count > 0 Then
                    Me.Pilih()
                    Me.Close()
                End If
            Else
                If Me.DGVHasil1.SelectedRows.Count > 0 Then
                    Me.Pilih()
                    Me.Close()
                End If
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

    Private Sub dgvHasil1_RowHeaderMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGVHasil1.RowHeaderMouseDoubleClick
        Try
            If Me.DGVHasil1.SelectedRows.Count > 0 Then
                Me.Pilih()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDatabase()
        Dim QueryResult As Double = 0
        Dim StartT As Date = Now
        Try
            Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If UntukPOS = False Then
                If Me.LocationControlledOnly = False Then
                    If Me.KriteriaCari = EnumSearchCriteria.ByCode Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_CodeLIKE(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByDescription Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_DESCLike(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_ProductSpecLIKE(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByBarcode Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_BarcodeLIKE(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    End If
                    Me.dgvHasil.BringToFront()
                Else
                    If Me.KriteriaCari = EnumSearchCriteria.ByCode Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_CodeLIKE(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByDescription Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_DESCLike(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_ProductSpecLIKE(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByBarcode Then
                        QueryResult = Me.Tblmaster_product_view0TableAdapter.FillBy_BarcodeLocationControl(Me.DseMasterProduct.tblmaster_product_view0, Me.txtCari.Text + "%")
                    End If
                    Me.dgvHasil.BringToFront()
                End If
            Else
                If Me.LocationControlledOnly = False Then
                    If Me.KriteriaCari = EnumSearchCriteria.ByCode Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_CodeLIKE(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByDescription Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_DESCLike(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_ProductSpecLIKE(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByBarcode Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_BarcodeLIKE(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    End If
                    Me.DGVHasil1.BringToFront()
                Else
                    If Me.KriteriaCari = EnumSearchCriteria.ByCode Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_CodeLIKE(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByDescription Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_DESCLike(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_ProductSpecLIKE(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    ElseIf Me.KriteriaCari = EnumSearchCriteria.ByBarcode Then
                        QueryResult = Me.Tblmaster_product_view01TableAdapter.FillBy_BarcodeLocationControl(Me.DseMasterProduct.tblmaster_product_view01, Me.txtCari.Text + "%")
                    End If
                    Me.DGVHasil1.BringToFront()
                End If
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    Me.tslQuery.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    Me.tslQuery.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                Me.tslQuery.Text = "No result found"
            End If

            If Me.UntukPOS = False Then
                Me.dgvHasil.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Else
                Me.DGVHasil1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class