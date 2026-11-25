Option Explicit On
Option Strict Off

Public Class FrmEditPriceList
    Public BackColorAll As Color = BackColor
    Dim TBL As New dseEditPriceList.TblEditPriceListDataTable
    Dim Buffer As Double = 0

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        txtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        txtCurrencyCode.BackColor = Me.txtPrimaryCurrency.BackColor
    End Sub

    Private Sub rboProductGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboProductGroup.CheckedChanged
        Me.SetelahClickRbo()
    End Sub

    Private Sub SetelahClickRbo()
        Try
            If rboProductGroup.Checked = True Then
                Me.txtProductGroup.ReadOnly = False
                Me.txtBrand.ReadOnly = True
                Me.txtCategory.ReadOnly = True
                Me.btnBrand.Enabled = False
                Me.btnCategory.Enabled = False
                Me.btnProductGroup.Enabled = True
                Me.txtProductGroup.Focus()
            ElseIf Me.rboCategory.Checked = True Then
                Me.txtProductGroup.ReadOnly = True
                Me.txtBrand.ReadOnly = True
                Me.txtCategory.ReadOnly = False
                Me.btnBrand.Enabled = False
                Me.btnCategory.Enabled = True
                Me.btnProductGroup.Enabled = False
                Me.txtCategory.Focus()
            ElseIf Me.rboBrand.Checked = True Then
                Me.txtProductGroup.ReadOnly = True
                Me.txtBrand.ReadOnly = False
                Me.txtCategory.ReadOnly = True
                Me.btnBrand.Enabled = True
                Me.btnCategory.Enabled = False
                Me.btnProductGroup.Enabled = False
                Me.txtBrand.Focus()
            ElseIf Me.rboAll.Checked = True Then
                Me.txtProductGroup.ReadOnly = True
                Me.txtBrand.ReadOnly = True
                Me.txtCategory.ReadOnly = True
                Me.btnBrand.Enabled = False
                Me.btnCategory.Enabled = False
                Me.btnProductGroup.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboBrand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboBrand.CheckedChanged
        Me.SetelahClickRbo()
    End Sub

    Private Sub rboCategory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboCategory.CheckedChanged
        Me.SetelahClickRbo()
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New Core_Module.FrmCariMasterCurrency
        Try
            i.BackColorAll = Me.BackColorAll
            i.txtCari.Text = Me.txtCurrencyCode.Text
            i.ShowDialog()
            If i.PrimaryCurrency <> -1 Then
                Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                Dim objC As New Core_Module.clsMasterCurrency
                Me.txtCurrencyCode.Text = objC.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                objC = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnProductGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProductGroup.Click
        Dim i As New Core_Module.FrmCariMasterProductGroup
        Try
            i.BackColorAll = Me.BackColorAll
            i.txtCari.Text = Me.txtProductGroup.Text
            i.ProductGroupLevel = 3
            i.ShowDialog()
            If i.PrimaryProductGroup <> -1 Then
                Me.txtPrimaryProductGroup.Text = i.PrimaryProductGroup
                Dim objProd As New Core_Module.ClsMasterProductGroup
                Me.txtProductGroup.Text = objProd.GetGroupCode(Me.txtPrimaryProductGroup.Text)
                objProd = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnBrand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrand.Click
        Dim i As New Core_Module.FrmCariMasterProductGroup
        Try
            i.BackColorAll = Me.BackColorAll
            i.txtCari.Text = Me.txtProductGroup.Text
            i.ProductGroupLevel = 2
            i.ShowDialog()
            If i.PrimaryProductGroup <> -1 Then
                Me.txtPrimaryBrand.Text = i.PrimaryProductGroup
                Dim objProd As New Core_Module.ClsMasterProductGroup
                Me.txtBrand.Text = objProd.GetGroupCode(Me.txtPrimaryBrand.Text)
                objProd = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnCategory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCategory.Click
        Dim i As New Core_Module.frmCariMasterCategory
        Try
            i.BackColorAll = Me.BackColorAll
            i.txtCari.Text = Me.txtCategory.Text
            i.ShowDialog()
            If i.PrimaryCategory <> -1 Then
                Me.txtPrimaryCategory.Text = i.PrimaryCategory
                Dim objC As New Core_Module.ClsMasterCategory
                Me.txtCategory.Text = objC.GetCategoryCode(Me.txtPrimaryCategory.Text)
                objC = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Dim objC As New Core_Module.clsMasterCurrency
                Dim Primary As Double = objC.GetPrimary(Me.txtCurrencyCode.Text)
                objC = Nothing
                If Primary <> -1 Then
                    Me.txtPrimaryCurrency.Text = Primary
                Else
                    Me.btnFindCurrency.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                Dim objC As New Core_Module.clsMasterCurrency
                Dim Primary As Double = objC.GetPrimary(Me.txtCurrencyCode.Text)
                objC = Nothing
                If Primary <> -1 Then
                    Me.txtPrimaryCurrency.Text = Primary
                Else
                    Me.btnFindCurrency.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductGroup_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductGroup.GotFocus
        If txtProductGroup.ReadOnly = False Then
            txtProductGroup.BackColor = Me.BackColorAll
        End If
    End Sub

    Private Sub txtProductGroup_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductGroup.KeyDown
        Try
            If Me.txtProductGroup.ReadOnly = False Then
                If e.KeyCode = Keys.Enter Then
                    Dim objProd As New Core_Module.ClsMasterProductGroup
                    Dim Primary As Double = objProd.GetPrimary(Me.txtProductGroup.Text, 3)
                    objProd = Nothing
                    If Primary <> -1 Then
                        Me.txtPrimaryProductGroup.Text = Primary
                    Else
                        Me.btnProductGroup.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductGroup_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductGroup.Leave
        Try
            If Me.txtProductGroup.ReadOnly = False Then
                If Me.txtProductGroup.Text <> "" Then
                    Dim objProd As New Core_Module.ClsMasterProductGroup
                    Dim Primary As Double = objProd.GetPrimary(Me.txtProductGroup.Text, 3)
                    objProd = Nothing
                    If Primary <> -1 Then
                        Me.txtPrimaryProductGroup.Text = Primary
                    Else
                        Me.btnProductGroup.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrand_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrand.GotFocus
        If txtBrand.ReadOnly = False Then
            txtBrand.BackColor = Me.BackColorAll
        End If
    End Sub

    Private Sub txtBrand_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBrand.KeyDown
        Try
            If Me.txtBrand.ReadOnly = False Then
                If e.KeyCode = Keys.Enter Then
                    Dim objProd As New Core_Module.ClsMasterProductGroup
                    Dim Primary As Double = objProd.GetPrimary(Me.txtBrand.Text, 2)
                    objProd = Nothing
                    If Primary <> -1 Then
                        Me.txtPrimaryProductGroup.Text = Primary
                    Else
                        Me.btnBrand.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrand_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrand.Leave
        Try
            If Me.txtProductGroup.ReadOnly = False Then
                If Me.txtBrand.Text <> "" Then
                    Dim objProd As New Core_Module.ClsMasterProductGroup
                    Dim Primary As Double = objProd.GetPrimary(Me.txtBrand.Text, 2)
                    objProd = Nothing
                    If Primary <> -1 Then
                        Me.txtPrimaryProductGroup.Text = Primary
                    Else
                        Me.btnBrand.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategory_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategory.GotFocus
        If txtCategory.ReadOnly = False Then
            txtCategory.BackColor = Me.BackColorAll
        End If
    End Sub

    Private Sub txtCategory_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCategory.KeyDown
        Try
            If Me.txtCategory.ReadOnly = False Then
                If e.KeyCode = Keys.Enter Then
                    Dim objC As New Core_Module.ClsMasterCategory
                    Dim Primary As Double = objC.GetPrimary(Me.txtCategory.Text)
                    objC = Nothing
                    If Primary <> -1 Then
                        Me.txtPrimaryCategory.Text = Primary
                    Else
                        Me.btnCategory.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategory_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategory.Leave
        Try
            If Me.txtCategory.ReadOnly = False Then
                If Me.txtCategory.Text <> "" Then
                    Dim objC As New Core_Module.ClsMasterCategory
                    Dim Primary As Double = objC.GetPrimary(Me.txtCategory.Text)
                    objC = Nothing
                    If Primary <> -1 Then
                        Me.txtPrimaryCategory.Text = Primary
                    Else
                        Me.btnCategory.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                Me.Filter()
            Else
                MsgBox("Please insert currency.", MsgBoxStyle.Information, "Information")
                Me.txtCurrencyCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Filter()
        Dim DTA As New dseEditPriceListTableAdapters.TblEditPriceListTableAdapter
        Try
            Me.TBL.Clear()
            Me.DataGridView2.Rows.Clear()
            If Me.rboAll.Checked = True Then
                DTA.FillBy_PrimaryCurrency(TBL, CDbl(Me.txtPrimaryCurrency.Text))
            ElseIf Me.rboCategory.Checked = True Then
                If Me.txtCategory.Text = "" Then
                    MsgBox("Please insert category.", MsgBoxStyle.Information, "Information")
                    Me.txtCategory.Focus()
                    Exit Try
                End If
                DTA.FillBy_PrimaryCurrencyNCategory(TBL, CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.txtPrimaryCategory.Text))
            ElseIf Me.rboProductGroup.Checked = True Then
                If Me.txtProductGroup.Text = "" Then
                    MsgBox("Please insert product group.", MsgBoxStyle.Information, "Information")
                    Me.txtProductGroup.Focus()
                    Exit Try
                End If
                DTA.FillBy_PrimaryCurrencyNPrimaryGroupCode(TBL, CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.txtPrimaryProductGroup.Text))
            ElseIf Me.rboBrand.Checked = True Then
                If Me.txtBrand.Text = "" Then
                    MsgBox("Please insert brand.", MsgBoxStyle.Information, "Information")
                    Me.txtBrand.Focus()
                    Exit Try
                End If
                DTA.FillBy_PrimaryCurrencyNBrand(TBL, CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.txtPrimaryBrand.Text))
            End If

            Me.DataGridView1.DataSource = TBL.DefaultView
            Me.DataGridView1.Refresh()
            If Me.DataGridView1.ColumnCount > 0 Then
                Me.DataGridView1.Columns(TBL.Primary_BrandColumn.ColumnName).Visible = False
                Me.DataGridView1.Columns(TBL.Primary_CategoryColumn.ColumnName).Visible = False
                Me.DataGridView1.Columns(TBL.Primary_CurrencyColumn.ColumnName).Visible = False
                Me.DataGridView1.Columns(TBL.Primary_Group_CodeColumn.ColumnName).Visible = False
                Me.DataGridView1.Columns(TBL.Primary_ProductColumn.ColumnName).Visible = False
                Me.DataGridView1.Columns(TBL.Product_CodeColumn.ColumnName).Frozen = True
                Me.DataGridView1.Columns(TBL.Product_CodeColumn.ColumnName).ReadOnly = True
                Me.DataGridView1.Columns(TBL.BrandColumn.ColumnName).ReadOnly = True
                Me.DataGridView1.Columns(TBL.Category_CodeColumn.ColumnName).ReadOnly = True
                Me.DataGridView1.Columns(TBL.Group_CodeColumn.ColumnName).ReadOnly = True
                Me.DataGridView1.Columns(TBL.Unit_PriceColumn.ColumnName).ReadOnly = False
                Me.DataGridView1.Columns(TBL.DiscountColumn.ColumnName).ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub KosongkanDetail()
        Try
            Me.txtPrimaryBrand.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryProductGroup.Text = "-1"
            Me.txtProductGroup.Text = ""
            Me.txtBrand.Text = ""
            Me.txtCategory.Text = ""
            Me.txtCurrencyCode.Text = ""
            Me.DataGridView1.Rows.Clear()
            Me.DataGridView2.Rows.Clear()
            Me.btnBrand.Enabled = False
            Me.btnCategory.Enabled = False
            Me.btnProductGroup.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmEditPriceList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.KosongkanDetail()
            Me.txtCurrencyCode.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Try
            If e.ColumnIndex = TBL.Columns.IndexOf(TBL.DiscountColumn) Or e.ColumnIndex = TBL.Columns.IndexOf(TBL.Unit_PriceColumn) Then
                If IsNumeric(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.Buffer = 0
                Else
                    Me.Buffer = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Try
            If IsNumeric(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                Exit Try
            End If
            If e.ColumnIndex = TBL.Columns.IndexOf(TBL.DiscountColumn) Then
                Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Math.Abs(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value)
                If CDbl(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = 100
                End If
            End If
            If CDbl(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) <> Me.Buffer Then
                If Me.DataGridView2.Rows.Count = 0 Then
                    Me.DataGridView2.Rows.Add(CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.Primary_ProductColumn), e.RowIndex).Value), _
                    CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.Unit_PriceColumn), e.RowIndex).Value), _
                    CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.DiscountColumn), e.RowIndex).Value))
                Else
                    Dim row As DataGridViewRow
                    For Each row In Me.DataGridView2.Rows
                        If row.Cells(Me.colPrimaryProduct.Index).Value <> _
                        CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.Primary_ProductColumn), e.RowIndex).Value) Then
                            'Tambahkan baru pada dgv 2
                            Me.DataGridView2.Rows.Add(CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.Primary_ProductColumn), e.RowIndex).Value), _
                            CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.Unit_PriceColumn), e.RowIndex).Value), _
                            CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.DiscountColumn), e.RowIndex).Value))
                        Else
                            row.Cells(Me.colUnitPrice.Index).Value = CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.Unit_PriceColumn), e.RowIndex).Value)
                            row.Cells(Me.colDiscount.Index).Value = CDbl(Me.DataGridView1.Item(TBL.Columns.IndexOf(TBL.DiscountColumn), e.RowIndex).Value)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If Me.DataGridView2.Rows.Count > 0 Then
                If Me.Save = True Then
                    MsgBox("Data updated", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Save() As Boolean
        Dim DTA As New dseEditPriceListTableAdapters.tblmaster_product_pricingTableAdapter
        Try
            Dim row As DataGridViewRow
            For Each row In Me.DataGridView2.Rows
                DTA.UpdateQuery_UbahHargaNDiskon(row.Cells(Me.colUnitPrice.Index).Value, row.Cells(Me.colDiscount.Index).Value, _
                row.Cells(Me.colPrimaryProduct.Index).Value, Me.txtPrimaryCurrency.Text)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub rboAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboAll.CheckedChanged
        Me.SetelahClickRbo()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        Try
            If e.ColumnIndex = Me.TBL.Columns.IndexOf(TBL.DiscountColumn) Then
                Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductGroup_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductGroup.LostFocus
        If txtProductGroup.ReadOnly = False Then
            txtProductGroup.BackColor = Me.txtPrimaryBrand.BackColor
        End If
    End Sub

    Private Sub txtBrand_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrand.LostFocus
        If txtBrand.ReadOnly = False Then
            txtBrand.BackColor = Me.txtPrimaryBrand.BackColor
        End If
    End Sub

    Private Sub txtCategory_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategory.LostFocus
        If txtCategory.ReadOnly = False Then
            txtCategory.BackColor = Me.txtPrimaryBrand.BackColor
        End If
    End Sub

End Class