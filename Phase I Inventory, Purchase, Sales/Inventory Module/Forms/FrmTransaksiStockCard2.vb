Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockCard2

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCD As New Core_Module.ClsCekData
    Dim RowDgv As DataGridViewRow
    Dim RowIndex As Integer
    Dim Buffer As String = ""
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim ObjWarehouse As New Core_Module.clsMasterWarehouse
    Dim ObjLocation As New Core_Module.clsMasterStorageArea

    Dim TblData As New dseTransaksiStockCard.Table_DataDataTable
    Dim RowData As dseTransaksiStockCard.Table_DataRow
    Dim ListOfProduct As New List(Of Double)
    Dim ListOfLocation As New List(Of Double)

    Public Enum EnumProductGroupLevel As Integer
        Category
        SubBrand
        Brand
        ProductGroup
    End Enum

    Private Sub txtProductCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.Enter
        Try
            Me.Buffer = Me.txtProductCode.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.GotFocus
        Try
            Me.txtProductCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductCode.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0
        Try
            If e.KeyValue = 13 Then
                Jml = ObjProduct.GetJml(Me.txtProductCode.Text)
                If Jml = 1 Then
                    Me.txtPrimaryProduct.Text = ObjProduct.GetPrimary(Me.txtProductCode.Text)
                    Me.txtDescription.Text = ObjProduct.GetDescription(Me.txtPrimaryProduct.Text)
                    Me.DtpFrom.Focus()
                    If Me.txtProductCode.Text <> Me.Buffer Then
                        Me.crvStockCard.ReportSource = Nothing
                        Me.crvStockCard.Refresh()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProductCode.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryProduct.Text = i.PrimaryProduct
                        Me.txtProductCode.Text = ObjProduct.GetProductCode(Me.txtPrimaryProduct.Text)
                        Me.DtpFrom.Focus()
                        If Me.txtProductCode.Text <> Me.Buffer Then
                            Me.crvStockCard.ReportSource = Nothing
                            Me.crvStockCard.Refresh()
                        End If
                    Else
                        Me.txtProductCode.Clear()
                        Me.txtPrimaryProduct.Text = "-1"
                        Me.txtDescription.Clear()
                        Me.txtProductCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.Leave
        Dim i As New frmCariMasterProduct
        Try
            If Me.txtProductCode.Text <> "" Then
                If ObjProduct.GetJml(Me.txtProductCode.Text) = 1 Then
                    Me.txtPrimaryProduct.Text = ObjProduct.GetPrimary(Me.txtProductCode.Text)
                    Me.txtDescription.Text = ObjProduct.GetDescription(Me.txtPrimaryProduct.Text)
                    Me.DtpFrom.Focus()
                    If Me.txtProductCode.Text <> Me.Buffer Then
                        Me.crvStockCard.ReportSource = Nothing
                        Me.crvStockCard.Refresh()
                    End If
                Else
                    If Me.txtProductCode.Text <> "" Then
                        i.txtCari.Text = Me.txtProductCode.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryProduct.Text = i.PrimaryProduct
                        Me.txtProductCode.Text = ObjProduct.GetProductCode(Me.txtPrimaryProduct.Text)
                        Me.DtpFrom.Focus()
                        If Me.txtProductCode.Text <> Me.Buffer Then
                            Me.crvStockCard.ReportSource = Nothing
                            Me.crvStockCard.Refresh()
                        End If
                    Else
                        Me.txtProductCode.Clear()
                        Me.txtPrimaryProduct.Text = "-1"
                        Me.txtDescription.Clear()
                        Me.txtProductCode.Focus()
                        Me.crvStockCard.ReportSource = Nothing
                        Me.crvStockCard.Refresh()
                    End If
                End If
            Else
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtDescription.Clear()
                Me.crvStockCard.ReportSource = Nothing
                Me.crvStockCard.Refresh()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim i As New frmCariMasterProduct
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtProductCode.Text <> "" Then
                i.txtCari.Text = Me.txtProductCode.Text
            End If

            i.BackColorAll = Me.BackColorAll
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.txtPrimaryProduct.Text = i.PrimaryProduct
                Me.txtProductCode.Text = ObjProduct.GetProductCode(Me.txtPrimaryProduct.Text)
                Me.txtDescription.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryProduct.Text)
                Me.DtpFrom.Focus()
            Else
                Me.txtProductCode.Clear()
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtDescription.Clear()
                Me.txtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error") 
        End Try
    End Sub

    Private Sub FrmTransaksiStockCard_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DtpFrom.Value = Now - TimeSpan.FromDays(30)
            Me.CboGroup.SelectedIndex = 0
            Me.CboScope.SelectedIndex = 0
            If Me.UserCode = "USER" Then
                LinkLabel1.Visible = True
                LinkLabel2.Visible = True
            End If
            Me.MinimizeBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboGroup.SelectedIndexChanged
        Dim objCd As New Core_Module.ClsCekData
        Try
            objCd.KosongkanControl(gboDetail.Controls, True, False, False)
            Me.txtPrimaryProduct.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.txtPrimarySubBrand.Text = "-1"
            Me.txtPrimaryBrand.Text = "-1"
            Me.txtPrimaryProductGroup.Text = "-1"
            Select Case CboGroup.Text
                Case "Product Code"
                    Me.lblProductCode.Visible = True
                    Me.txtProductCode.Visible = True
                    Me.btnFindProduct.Visible = True
                    Me.txtDescription.Visible = True

                    Me.lblCategory.Visible = False
                    Me.txtCategoryCode.Visible = False
                    Me.btnFindCategory.Visible = False
                    Me.txtCategoryDesc.Visible = False

                    Me.lblSubBrand.Visible = False
                    Me.txtSubBrandCode.Visible = False
                    Me.btnFindSubBrand.Visible = False
                    Me.txtSubBrandDesc.Visible = False

                    Me.lblBrand.Visible = False
                    Me.txtBrandCode.Visible = False
                    Me.btnFindBrand.Visible = False
                    Me.txtBrandDesc.Visible = False

                    Me.lblProductGroup.Visible = False
                    Me.txtProductGroupCode.Visible = False
                    Me.btnFindProductGroup.Visible = False
                    Me.txtProductGroupDesc.Visible = False

                    Me.CheckBox2.Visible = False
                    Exit Select
                Case "Category"
                    Me.lblProductCode.Visible = False
                    Me.txtProductCode.Visible = False
                    Me.btnFindProduct.Visible = False
                    Me.txtDescription.Visible = False

                    Me.lblCategory.Visible = True
                    Me.txtCategoryCode.Visible = True
                    Me.btnFindCategory.Visible = True
                    Me.txtCategoryDesc.Visible = True

                    Me.lblSubBrand.Visible = False
                    Me.txtSubBrandCode.Visible = False
                    Me.btnFindSubBrand.Visible = False
                    Me.txtSubBrandDesc.Visible = False

                    Me.lblBrand.Visible = False
                    Me.txtBrandCode.Visible = False
                    Me.btnFindBrand.Visible = False
                    Me.txtBrandDesc.Visible = False

                    Me.lblProductGroup.Visible = False
                    Me.txtProductGroupCode.Visible = False
                    Me.btnFindProductGroup.Visible = False
                    Me.txtProductGroupDesc.Visible = False

                    Me.CheckBox2.Visible = True
                    Exit Select
                Case "Sub Brand"
                    Me.lblProductCode.Visible = False
                    Me.txtProductCode.Visible = False
                    Me.btnFindProduct.Visible = False
                    Me.txtDescription.Visible = False

                    Me.lblCategory.Visible = False
                    Me.txtCategoryCode.Visible = False
                    Me.btnFindCategory.Visible = False
                    Me.txtCategoryDesc.Visible = False

                    Me.lblSubBrand.Visible = True
                    Me.txtSubBrandCode.Visible = True
                    Me.btnFindSubBrand.Visible = True
                    Me.txtSubBrandDesc.Visible = True

                    Me.lblBrand.Visible = False
                    Me.txtBrandCode.Visible = False
                    Me.btnFindBrand.Visible = False
                    Me.txtBrandDesc.Visible = False

                    Me.lblProductGroup.Visible = False
                    Me.txtProductGroupCode.Visible = False
                    Me.btnFindProductGroup.Visible = False
                    Me.txtProductGroupDesc.Visible = False

                    Me.CheckBox2.Visible = True
                    Exit Select
                Case "Brand"
                    Me.lblProductCode.Visible = False
                    Me.txtProductCode.Visible = False
                    Me.btnFindProduct.Visible = False
                    Me.txtDescription.Visible = False

                    Me.lblCategory.Visible = False
                    Me.txtCategoryCode.Visible = False
                    Me.btnFindCategory.Visible = False
                    Me.txtCategoryDesc.Visible = False

                    Me.lblSubBrand.Visible = False
                    Me.txtSubBrandCode.Visible = False
                    Me.btnFindSubBrand.Visible = False
                    Me.txtSubBrandDesc.Visible = False

                    Me.lblBrand.Visible = True
                    Me.txtBrandCode.Visible = True
                    Me.btnFindBrand.Visible = True
                    Me.txtBrandDesc.Visible = True

                    Me.lblProductGroup.Visible = False
                    Me.txtProductGroupCode.Visible = False
                    Me.btnFindProductGroup.Visible = False
                    Me.txtProductGroupDesc.Visible = False

                    Me.CheckBox2.Visible = True
                    Exit Select
                Case "Product Group"
                    Me.lblProductCode.Visible = False
                    Me.txtProductCode.Visible = False
                    Me.btnFindProduct.Visible = False
                    Me.txtDescription.Visible = False

                    Me.lblCategory.Visible = False
                    Me.txtCategoryCode.Visible = False
                    Me.btnFindCategory.Visible = False
                    Me.txtCategoryDesc.Visible = False

                    Me.lblSubBrand.Visible = False
                    Me.txtSubBrandCode.Visible = False
                    Me.btnFindSubBrand.Visible = False
                    Me.txtSubBrandDesc.Visible = False

                    Me.lblBrand.Visible = False
                    Me.txtBrandCode.Visible = False
                    Me.btnFindBrand.Visible = False
                    Me.txtBrandDesc.Visible = False

                    Me.lblProductGroup.Visible = True
                    Me.txtProductGroupCode.Visible = True
                    Me.btnFindProductGroup.Visible = True
                    Me.txtProductGroupDesc.Visible = True

                    Me.CheckBox2.Visible = True
                    Exit Select
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategoryCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.GotFocus
        Try
            Me.txtCategoryCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategoryCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCategoryCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtCategoryCode.Text <> "" Then
                    If Me.ObjCategory.ApakahSudahAda(Me.txtCategoryCode.Text) = True Then
                        Me.txtPrimaryCategory.Text = Me.ObjCategory.GetPrimary(Me.txtCategoryCode.Text)
                        Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text)
                    Else
                        Me.btnFindCategory.PerformClick()
                    End If
                Else
                    Me.txtPrimaryCategory.Text = "-1"
                    Me.txtCategoryDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategoryCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.Leave
        Try
            If Me.txtCategoryCode.Text <> "" Then
                If Me.ObjCategory.ApakahSudahAda(Me.txtCategoryCode.Text) = True Then
                    Me.txtPrimaryCategory.Text = Me.ObjCategory.GetPrimary(Me.txtCategoryCode.Text)
                    Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(Me.txtPrimaryCategory.Text)
                Else
                    Me.btnFindCategory.PerformClick()
                End If
            Else
                Me.txtPrimaryCategory.Text = "-1"
                Me.txtCategoryDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCategory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCategory.Click
        Dim frm As New frmCariMasterCategory
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtCategoryCode.Text <> "" Then
                frm.txtCari.Text = Me.txtCategoryCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimaryCategory > 0 Then
                Me.txtPrimaryCategory.Text = frm.PrimaryCategory
                Me.txtCategoryCode.Text = Me.ObjCategory.GetCategoryCode(frm.PrimaryCategory)
                Me.txtCategoryDesc.Text = Me.ObjCategory.GetDesc(frm.PrimaryCategory)
            Else
                Me.txtPrimaryCategory.Text = "-1"
                Me.txtCategoryCode.Clear()
                Me.txtCategoryDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtSubBrandCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubBrandCode.GotFocus
        Try
            Me.txtSubBrandCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubBrandCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSubBrandCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtSubBrandCode.Text <> "" Then
                    If Me.ObjProductGroup.ApakahSudahAda(Me.txtSubBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.SubBrand) = True Then
                        Me.txtPrimarySubBrand.Text = Me.ObjProductGroup.GetPrimary(Me.txtSubBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.SubBrand)
                        Me.txtSubBrandDesc.Text = Me.ObjProductGroup.GetGroupDesc(Me.txtPrimarySubBrand.Text)
                    Else
                        Me.btnFindSubBrand.PerformClick()
                    End If
                Else
                    Me.txtPrimarySubBrand.Text = "-1"
                    Me.txtSubBrandDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubBrandCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubBrandCode.Leave
        Try
            If Me.txtSubBrandCode.Text <> "" Then
                If Me.ObjProductGroup.ApakahSudahAda(Me.txtSubBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.SubBrand) = True Then
                    Me.txtPrimarySubBrand.Text = Me.ObjProductGroup.GetPrimary(Me.txtSubBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.SubBrand)
                    Me.txtSubBrandDesc.Text = Me.ObjProductGroup.GetGroupDesc(Me.txtPrimarySubBrand.Text)
                Else
                    Me.btnFindSubBrand.PerformClick()
                End If
            Else
                Me.txtPrimarySubBrand.Text = "-1"
                Me.txtSubBrandDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSubBrand_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSubBrand.Click
        Dim frm As New FrmCariMasterProductGroup
        frm.ProductGroupLevel = FrmTransaksiStockCard2.EnumProductGroupLevel.SubBrand
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtSubBrandCode.Text <> "" Then
                frm.txtCari.Text = Me.txtSubBrandCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimaryProductGroup > 0 Then
                Me.txtPrimarySubBrand.Text = frm.PrimaryProductGroup
                Me.txtSubBrandCode.Text = Me.ObjProductGroup.GetGroupCode(frm.PrimaryProductGroup)
                Me.txtSubBrandDesc.Text = Me.ObjProductGroup.GetGroupDesc(frm.PrimaryProductGroup)
            Else
                Me.txtPrimarySubBrand.Text = "-1"
                Me.txtSubBrandDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtBrandCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCode.GotFocus
        Try
            Me.txtBrandCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBrandCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtBrandCode.Text <> "" Then
                    If Me.ObjProductGroup.ApakahSudahAda(Me.txtBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.Brand) = True Then
                        Me.txtPrimaryBrand.Text = Me.ObjProductGroup.GetPrimary(Me.txtBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.Brand)
                        Me.txtBrandDesc.Text = Me.ObjProductGroup.GetGroupDesc(Me.txtPrimaryBrand.Text)
                    Else
                        Me.btnFindBrand.PerformClick()
                    End If
                Else
                    Me.txtPrimaryBrand.Text = "-1"
                    Me.txtBrandDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCode.Leave
        Try
            If Me.txtBrandCode.Text <> "" Then
                If Me.ObjProductGroup.ApakahSudahAda(Me.txtBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.Brand) = True Then
                    Me.txtPrimaryBrand.Text = Me.ObjProductGroup.GetPrimary(Me.txtBrandCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.Brand)
                    Me.txtBrandDesc.Text = Me.ObjProductGroup.GetGroupDesc(Me.txtPrimaryBrand.Text)
                Else
                    Me.btnFindBrand.PerformClick()
                End If
            Else
                Me.txtPrimaryBrand.Text = "-1"
                Me.txtBrandDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindBrand_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindBrand.Click
        Dim frm As New FrmCariMasterProductGroup
        frm.ProductGroupLevel = FrmTransaksiStockCard2.EnumProductGroupLevel.Brand
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtBrandCode.Text <> "" Then
                frm.txtCari.Text = Me.txtBrandCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimaryProductGroup > 0 Then
                Me.txtPrimaryBrand.Text = frm.PrimaryProductGroup
                Me.txtBrandCode.Text = Me.ObjProductGroup.GetGroupCode(frm.PrimaryProductGroup)
                Me.txtBrandDesc.Text = Me.ObjProductGroup.GetGroupDesc(frm.PrimaryProductGroup)
            Else
                Me.txtPrimaryBrand.Text = "-1"
                Me.txtBrandDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtProductGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductGroupCode.GotFocus
        Try
            Me.txtProductGroupCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductGroupCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtProductGroupCode.Text <> "" Then
                    If Me.ObjProductGroup.ApakahSudahAda(Me.txtProductGroupCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.ProductGroup) = True Then
                        Me.txtPrimaryProductGroup.Text = Me.ObjProductGroup.GetPrimary(Me.txtProductGroupCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.ProductGroup)
                        Me.txtProductGroupDesc.Text = Me.ObjProductGroup.GetGroupDesc(Me.txtPrimaryProductGroup.Text)
                    Else
                        Me.btnFindProductGroup.PerformClick()
                    End If
                Else
                    Me.txtPrimaryProductGroup.Text = "-1"
                    Me.txtProductGroupDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductGroupCode.Leave
        Try
            If Me.txtProductGroupCode.Text <> "" Then
                If Me.ObjProductGroup.ApakahSudahAda(Me.txtProductGroupCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.ProductGroup) = True Then
                    Me.txtPrimaryProductGroup.Text = Me.ObjProductGroup.GetPrimary(Me.txtProductGroupCode.Text, FrmTransaksiStockCard2.EnumProductGroupLevel.ProductGroup)
                    Me.txtProductGroupDesc.Text = Me.ObjProductGroup.GetGroupDesc(Me.txtPrimaryProductGroup.Text)
                Else
                    Me.btnFindProductGroup.PerformClick()
                End If
            Else
                Me.txtPrimaryProductGroup.Text = "-1"
                Me.txtProductGroupDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindProductGroup.Click
        Dim frm As New FrmCariMasterProductGroup
        frm.ProductGroupLevel = FrmTransaksiStockCard2.EnumProductGroupLevel.ProductGroup

        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtProductGroupCode.Text <> "" Then
                frm.txtCari.Text = Me.txtProductGroupCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimaryProductGroup > 0 Then
                Me.txtPrimaryProductGroup.Text = frm.PrimaryProductGroup
                Me.txtProductGroupCode.Text = Me.ObjProductGroup.GetGroupCode(frm.PrimaryProductGroup)
                Me.txtProductGroupDesc.Text = Me.ObjProductGroup.GetGroupDesc(frm.PrimaryProductGroup)
            Else
                Me.txtPrimaryProductGroup.Text = "-1"
                Me.txtProductGroupDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.LostFocus
        Try
            Me.txtProductCode.BackColor = Me.txtPrimaryProduct.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCategoryCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCategoryCode.LostFocus
        Try
            Me.txtCategoryCode.BackColor = Me.txtPrimaryProduct.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSubBrandCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSubBrandCode.LostFocus
        Try
            Me.txtSubBrandCode.BackColor = Me.txtPrimaryProduct.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBrandCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBrandCode.LostFocus
        Try
            Me.txtBrandCode.BackColor = Me.txtPrimaryProduct.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductGroupCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductGroupCode.LostFocus
        Try
            Me.txtProductGroupCode.BackColor = Me.txtPrimaryProduct.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvLocation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.Col1Level.Index Or e.ColumnIndex = Me.Col1Code.Index Then
                Me.Buffer = Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value
                Exit Sub
            End If
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvLocation.CellEndEdit
        Try
            If e.ColumnIndex = Me.Col1Level.Index Then
                If Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                    Exit Sub
                End If
                Me.DgvLocation(Me.Col1Code.Index, e.RowIndex).Value = ""
                Me.DgvLocation(Me.Col1Description.Index, e.RowIndex).Value = ""
                Me.DgvLocation(Me.Col1Primary.Index, e.RowIndex).Value = "-1"
            End If
            If e.ColumnIndex = Me.Col1Code.Index Then
                If Not IsNothing(Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value) Then
                    If Not String.IsNullOrEmpty(Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value) Then
                        Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                        If Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                            Exit Sub
                        End If
                        If Me.GetLocationRowIndex(Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value, Me.DgvLocation(Me.Col1Code.Index, e.RowIndex).Value, e.RowIndex) <> -1 Then
                            MsgBox(Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value & " is already exist.", MsgBoxStyle.Information, "Information")
                            If Not String.IsNullOrEmpty(Me.Buffer) Then
                                Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Else
                                Me.RowIndex = e.RowIndex
                                Me.BgwLocation.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                        If Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value = "Warehouse" Then
                            If Me.ObjWarehouse.ApakahSudahAda(Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value) Then
                                Me.DgvLocation(Me.Col1Primary.Index, e.RowIndex).Value = Me.ObjWarehouse.GetPrimary(Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value)
                                Me.DgvLocation(Me.Col1Description.Index, e.RowIndex).Value = Me.ObjWarehouse.GetDescription(Me.DgvLocation(Me.Col1Primary.Index, e.RowIndex).Value)
                            Else
                                Dim Frm As New Core_Module.frmCariMasterWarehouse
                                Frm.txtCari.Text = Me.DgvLocation(Me.Col1Code.Index, e.RowIndex).Value
                                Frm.ShowDialog()
                                If Frm.PrimaryWarehouse > -1 Then
                                    If Me.GetLocationRowIndex(Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value, Me.ObjWarehouse.GetWarehouseCode(Frm.PrimaryWarehouse), e.RowIndex) <> -1 Then
                                        MsgBox(Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value & " is already exist.", MsgBoxStyle.Information, "Information")
                                        If Not String.IsNullOrEmpty(Me.Buffer) Then
                                            Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Else
                                            Me.RowIndex = e.RowIndex
                                            Me.BgwLocation.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    End If
                                    Me.DgvLocation(Me.Col1Primary.Index, e.RowIndex).Value = Frm.PrimaryWarehouse
                                    Me.DgvLocation(Me.Col1Code.Index, e.RowIndex).Value = Me.ObjWarehouse.GetWarehouseCode(Frm.PrimaryWarehouse)
                                    Me.DgvLocation(Me.Col1Description.Index, e.RowIndex).Value = Me.ObjWarehouse.GetDescription(Frm.PrimaryWarehouse)
                                Else
                                    If Not String.IsNullOrEmpty(Me.Buffer) Then
                                        Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Else
                                        Me.RowIndex = e.RowIndex
                                        Me.BgwLocation.RunWorkerAsync()
                                    End If
                                End If
                            End If
                        End If
                        If Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value = "Location" Then
                            If Me.ObjLocation.ApakahSudahAda(Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value) Then
                                Me.DgvLocation(Me.Col1Primary.Index, e.RowIndex).Value = Me.ObjLocation.GetPrimary(Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value)
                                Me.DgvLocation(Me.Col1Description.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(Me.DgvLocation(Me.Col1Primary.Index, e.RowIndex).Value)
                            Else
                                Dim Frm As New Core_Module.FrmCariMasterStorageArea
                                Frm.txtCari.Text = Me.DgvLocation(Me.Col1Code.Index, e.RowIndex).Value
                                Frm.ShowDialog()
                                If Frm.PrimaryStorageArea > -1 Then
                                    If Me.GetLocationRowIndex(Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value, Me.ObjLocation.GetStorageAreaCode(Frm.PrimaryStorageArea), e.RowIndex) <> -1 Then
                                        MsgBox(Me.DgvLocation(Me.Col1Level.Index, e.RowIndex).Value & " is already exist.", MsgBoxStyle.Information, "Information")
                                        If Not String.IsNullOrEmpty(Me.Buffer) Then
                                            Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        Else
                                            Me.RowIndex = e.RowIndex
                                            Me.BgwLocation.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    End If
                                    Me.DgvLocation(Me.Col1Primary.Index, e.RowIndex).Value = Frm.PrimaryStorageArea
                                    Me.DgvLocation(Me.Col1Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(Frm.PrimaryStorageArea)
                                    Me.DgvLocation(Me.Col1Description.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(Frm.PrimaryStorageArea)
                                Else
                                    If Not String.IsNullOrEmpty(Me.Buffer) Then
                                        Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Else
                                        Me.RowIndex = e.RowIndex
                                        Me.BgwLocation.RunWorkerAsync()
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Not String.IsNullOrEmpty(Me.Buffer) Then
                            Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.RowIndex = e.RowIndex
                            Me.BgwLocation.RunWorkerAsync()
                        End If
                    End If
                Else
                    If Not String.IsNullOrEmpty(Me.Buffer) Then
                        Me.DgvLocation(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Else
                        Me.RowIndex = e.RowIndex
                        Me.BgwLocation.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetLocationRowIndex(ByVal vScope As String, ByVal vCode As String, ByVal vCurentRowIndex As Integer) As Integer
        Try
            Me.DgvLocation.AllowUserToAddRows = False
            For Each RowDgv In Me.DgvLocation.Rows
                If RowDgv.Index <> vCurentRowIndex Then
                    If vCode.Trim <> "" And RowDgv.Cells(Me.Col1Level.Index).Value = vScope And RowDgv.Cells(Me.Col1Code.Index).Value.ToString.Trim.ToLower = vCode.ToString.Trim.ToLower Then
                        Return RowDgv.Index
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvLocation.AllowUserToAddRows = True
        End Try
    End Function

    Private Sub DgvLocation_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvLocation.DataError
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvLocation.RowsAdded
        Try
            RowDgv = Me.DgvLocation.Rows(Me.DgvLocation.Rows.Count - 1)
            RowDgv.Cells(Me.Col1No.Index).Value = RowDgv.Index + 1
            RowDgv.Cells(Me.Col1Code.Index).Value = ""
            RowDgv.Cells(Me.Col1Description.Index).Value = ""
            RowDgv.Cells(Me.Col1Level.Index).Value = "Location"
            RowDgv.Cells(Me.Col1Primary.Index).Value = "-1"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvLocation.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvLocation_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvLocation.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvLocation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwLocation_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwLocation.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwLocation_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwLocation.RunWorkerCompleted
        Try
            If Me.RowIndex >= 0 And Me.RowIndex < Me.DgvLocation.Rows.Count - 1 Then
                Me.DgvLocation.Rows.RemoveAt(Me.RowIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try

            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("Invalid period filter.", MsgBoxStyle.Information, "Information")
                Return
            End If

            If Me.CboGroup.SelectedIndex = 0 Then
                If Me.txtPrimaryProduct.Text = "" Or Me.txtPrimaryProduct.Text = "-1" Then
                    MsgBox("Define a product first.", MsgBoxStyle.Information, "Information")
                    Return
                End If
            ElseIf Me.CboGroup.SelectedIndex = 1 Then
                If Me.txtPrimaryCategory.Text = "" Or Me.txtPrimaryCategory.Text = "-1" Then
                    MsgBox("Define a category first.", MsgBoxStyle.Information, "Information")
                    Return
                End If
            ElseIf Me.CboGroup.SelectedIndex = 2 Then
                If Me.txtPrimarySubBrand.Text = "" Or Me.txtSubBrandCode.Text = "-1" Then
                    MsgBox("Define a sub brand first.", MsgBoxStyle.Information, "Information")
                    Return
                End If
            ElseIf Me.CboGroup.SelectedIndex = 3 Then
                If Me.txtPrimaryBrand.Text = "" Or Me.txtPrimaryBrand.Text = "-1" Then
                    MsgBox("Define a brand first.", MsgBoxStyle.Information, "Information")
                    Return
                End If
            ElseIf Me.CboGroup.SelectedIndex = 4 Then
                If Me.txtPrimaryProductGroup.Text = "" Or Me.txtPrimaryProductGroup.Text = "-1" Then
                    MsgBox("Define a product group first.", MsgBoxStyle.Information, "Information")
                    Return
                End If
            End If

            Me.DgvLocation.AllowUserToAddRows = False
            For i As Integer = Me.DgvLocation.RowCount - 1 To 0 Step -1
                If Me.DgvLocation.Rows(i).Cells(Me.Col1Primary.Index).Value.ToString = "" Or Me.DgvLocation.Rows(i).Cells(Me.Col1Primary.Index).Value.ToString = "-1" Then
                    Me.DgvLocation.Rows.RemoveAt(i)
                End If
            Next
            Me.ObjCD.SetNumber(Me.DgvLocation)
            Me.DgvLocation.AllowUserToAddRows = True

            Me.LoadStockCard()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub FillCodeDesc()
        Dim dtaProduct As New dseTransaksiStockCardTableAdapters.Table_ProductTableAdapter
        Dim tblProduct As New dseTransaksiStockCard.Table_ProductDataTable
        Dim rowProduct As dseTransaksiStockCard.Table_ProductRow

        Dim dtaLocation As New dseTransaksiStockCardTableAdapters.Table_LocationTableAdapter
        Dim tblLocation As New dseTransaksiStockCard.Table_LocationDataTable
        Dim rowLocation As dseTransaksiStockCard.Table_LocationRow
        Dim index0 As Integer = 0
        Dim index1 As Integer = 0
        Try
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            dtaLocation.Connection.ConnectionString = Core_Module.GetConstring

            dtaProduct.Fill(tblProduct, Me.ListOfProduct(0), Me.ListOfProduct(Me.ListOfProduct.Count - 1))
            dtaLocation.Fill(tblLocation, -1, 999999999999)

            Me.TblData.DefaultView.RowFilter = "[Is Begining] = -1"
            Me.TblData.DefaultView.Sort = "[Primary Product]"
            tblProduct.DefaultView.RowFilter = ""
            tblProduct.DefaultView.Sort = "[Primary Product]"
            While index0 < Me.TblData.DefaultView.Count
                RowData = Me.TblData.DefaultView(index0).Row
                While index1 < tblProduct.DefaultView.Count
                    rowProduct = tblProduct.DefaultView(index1).Row
                    If rowProduct.Primary_Product < RowData.Primary_Product Then
                        index1 += 1
                    ElseIf rowProduct.Primary_Product = RowData.Primary_Product Then
                        RowData.Product_Code = rowProduct.Product_Code
                        RowData.Product_Description = rowProduct.Product_Description
                        RowData.Default_Precision = rowProduct._Default__Precision_
                        RowData.Primary_Unit = rowProduct.Primary_Unit
                        RowData.Unit_Code = rowProduct.Unit_Code
                        RowData.Unit_Description = rowProduct.Unit_Description
                        If rowProduct.Location_Controlled = 0 Then
                            RowData.Location_Controlled = False
                        Else
                            RowData.Location_Controlled = True
                        End If
                        Exit While
                    ElseIf rowProduct.Primary_Product > RowData.Primary_Product Then
                        'MsgBox("Logic Error: Class:FrmTransaksiStockCard2, Function: CombineProductLocation" & vbNewLine & "Desc: Product not found", MsgBoxStyle.Information, "Information")
                        Exit While
                    End If
                End While
                index0 += 1
            End While

            index0 = 0
            index1 = 0
            Me.TblData.DefaultView.RowFilter = "[Is Begining] = -1"
            Me.TblData.DefaultView.Sort = "[Primary Location]"
            tblLocation.DefaultView.RowFilter = ""
            tblLocation.DefaultView.Sort = "[Primary Location]"
            While index0 < Me.TblData.DefaultView.Count
                RowData = Me.TblData.DefaultView(index0).Row
                While index1 < tblLocation.DefaultView.Count
                    rowLocation = tblLocation.DefaultView(index1).Row
                    If rowLocation.Primary_Location < RowData.Primary_Location Then
                        index1 += 1
                    ElseIf rowLocation.Primary_Location = RowData.Primary_Location Then
                        RowData.Location_Code = rowLocation.Location_Code
                        RowData.Location_Description = rowLocation.Location_Description
                        RowData.Primary_Warehouse = rowLocation.Primary_Warehouse
                        RowData.Warehouse_Code = rowLocation.Warehouse_Code
                        RowData.Warehouse_Description = rowLocation.Warehouse_Description
                        Exit While
                    ElseIf rowLocation.Primary_Location > RowData.Primary_Location Then
                        'MsgBox("Logic Error: Class:FrmTransaksiStockCard2, Funct: CombineProductLocation" & vbNewLine & "Desc: Location not found", MsgBoxStyle.Information, "Information")
                        Exit While
                    End If
                End While
                index0 += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadStockCard()
        Dim MonthlyClosingPeriod As Date = #1/1/1950#
        Dim TblClosing As New dseTransaksiStockCard.tbltrans_purchase_closing_item_productqtyDataTable
        Dim TblLog As New dseTransaksiStockCard.Table_LogDataTable
        Dim DtaClosing As New dseTransaksiStockCardTableAdapters.tbltrans_purchase_closing_item_productqtyTableAdapter
        Dim index0 As Integer = 0
        Dim index1 As Integer = 0
        Dim PrimaryMonthlyClosing As Double = -1
        Try
            Me.Cursor = Cursors.WaitCursor
            TblData.Rows.Clear()
            DtaClosing.Connection.ConnectionString = Core_Module.GetConstring
            Me.SetProductList()
            Me.SetlocationList()
            PrimaryMonthlyClosing = Me.GetPrimaryMonthlyClosing(Me.DtpFrom.Value.Date)
            MonthlyClosingPeriod = Me.GetClosingPeriod(PrimaryMonthlyClosing)
            Me.FillLog(TblLog, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date, True)
            DtaClosing.Fill(TblClosing, PrimaryMonthlyClosing, Me.ListOfProduct(0), Me.ListOfProduct(Me.ListOfProduct.Count - 1))
            Me.CalculateBeginingBalance(MonthlyClosingPeriod, TblClosing)
            Me.FillCodeDesc()
            Me.WriteLogCard(TblLog)
            Me.GroupAndCalculateEndingBalance() 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub GroupAndCalculateEndingBalance()
        Dim DS As New dseTransaksiStockCard
        Dim TblPrint As dseTransaksiStockCard.Table_DataDataTable = DS.Table_Data
        Dim RowPrint As dseTransaksiStockCard.Table_DataRow
        Dim LastBalance As Decimal = 0
        Dim LastPrimaryProduct As Double = -1
        Dim LastPrimaryLocation As Double = -1
        Dim LastPrimaryWarehouse As Double = -1
        Dim reportGroup As New CrpTransaksiStockCard2_Group
        Dim report As New CrpTransaksiStockCard2
        Dim ListGroup2 As New List(Of String)
        Try

            If Me.CboScope.SelectedIndex = 0 Then
                Me.TblData.DefaultView.RowFilter = ""
                Me.TblData.DefaultView.Sort = "[Product Code], [Is Begining], [Date], [Moving In] DESC, [Moving Out]"
                For Each DV As Data.DataRowView In Me.TblData.DefaultView
                    RowData = DV.Row
                    If RowData.Primary_Product = LastPrimaryProduct Then
                        If RowData.Is_Begining = -1 Then
                            RowPrint.Balance += RowData.Moving_In
                            GoTo nextLoop0
                        Else
                            GoTo addRow0
                        End If
                    Else
                        GoTo addRow0
                    End If

addRow0:
                    RowPrint = TblPrint.NewTable_DataRow
                    For i As Integer = 0 To TblData.Columns.Count - 1
                        RowPrint.Item(i) = RowData.Item(i)
                    Next
                    RowPrint.Group_1 = RowPrint.Product_Code 
                    If RowPrint.Is_Begining = -1 Then
                        RowPrint.Balance = RowPrint.Moving_In
                        RowPrint.Moving_In = 0
                    Else
                        RowPrint.Balance = LastBalance + RowPrint.Moving_In - RowPrint.Moving_Out
                    End If
                    If Me.ChkShowBatch.Checked And RowPrint.Is_Begining = 0 Then
                        If Not String.IsNullOrEmpty(RowPrint.Lot_No) Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Batch: " & RowPrint.Lot_No
                        End If
                        If RowPrint.Expiry_Date > #1/1/1950# Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Expiry: " & Format(RowPrint.Expiry_Date, "dd MMM yyyy")
                        End If
                        If Not String.IsNullOrEmpty(RowPrint.Location_Code) Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Location: " & RowPrint.Location_Code
                        End If
                    End If
                    If Me.ChkShowItemRemarks.Checked Then
                        If Not String.IsNullOrEmpty(RowPrint.Remarks) Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Remarks: " & RowPrint.Remarks
                        End If
                    End If
                    TblPrint.Rows.Add(RowPrint)

nextLoop0:
                    LastBalance = RowPrint.Balance
                    LastPrimaryProduct = RowPrint.Primary_Product
                Next
                If TblPrint.Count > 0 Then
                    RowPrint = TblPrint.Rows(TblPrint.Count - 1)
                    RowPrint.Last_Group = 1
                End If
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_Period.ParameterFieldName, Format(Me.DtpFrom.Value.Date, "dd-MMM-yyyy") & " to " & Format(Me.DtpTo.Value.Date, "dd-MMM-yyyy"))
                Me.crvStockCard.ReportSource = report
                Me.crvStockCard.Zoom(90)
                Me.crvStockCard.Refresh()
                Me.crvStockCard.Show()
            ElseIf Me.CboScope.SelectedIndex = 1 Then
                Me.TblData.DefaultView.RowFilter = ""
                Me.TblData.DefaultView.Sort = "[Product Code], [Warehouse Code], [Is Begining], [Date], [Moving In] DESC, [Moving Out]"
                For Each DV As Data.DataRowView In Me.TblData.DefaultView
                    RowData = DV.Row
                   
                    If RowData.Primary_Product = LastPrimaryProduct And RowData.Primary_Warehouse = LastPrimaryWarehouse Then
                        If RowData.Is_Begining = -1 Then
                            RowPrint.Balance += RowData.Moving_In
                            GoTo nextLoop
                        Else
                            GoTo addRow
                        End If
                    Else
                        GoTo addRow
                    End If

addRow:
                    RowPrint = TblPrint.NewTable_DataRow
                    For i As Integer = 0 To TblData.Columns.Count - 1
                        RowPrint.Item(i) = RowData.Item(i)
                    Next
                    RowPrint.Group_1 = RowPrint.Product_Code 
                    RowPrint.Group_2 = RowPrint.Warehouse_Code
                    If RowPrint.Is_Begining = -1 Then
                        RowPrint.Balance = RowPrint.Moving_In
                        RowPrint.Moving_In = 0
                    Else
                        RowPrint.Balance = LastBalance + RowPrint.Moving_In - RowPrint.Moving_Out
                    End If
                    If Me.ChkShowBatch.Checked Then
                        If Not String.IsNullOrEmpty(RowPrint.Lot_No) Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Batch: " & RowPrint.Lot_No
                        End If
                        If RowPrint.Expiry_Date > #1/1/1950# Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Expiry: " & Format(RowPrint.Expiry_Date, "dd MMM yyyy")
                        End If
                    End If
                    If Me.ChkShowItemRemarks.Checked Then
                        If Not String.IsNullOrEmpty(RowPrint.Remarks) Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Remarks: " & RowPrint.Remarks
                        End If
                    End If
                    TblPrint.Rows.Add(RowPrint)

nextLoop:
                    LastBalance = RowPrint.Balance
                    LastPrimaryProduct = RowPrint.Primary_Product
                    LastPrimaryWarehouse = RowPrint.Primary_Warehouse

                    If ListGroup2.Contains(RowPrint.Group_2) = False Then
                        ListGroup2.Add(RowPrint.Group_2)
                    End If
                Next
                If TblPrint.Count > 0 Then
                    RowPrint = TblPrint.Rows(TblPrint.Count - 1)
                    RowPrint.Last_Group = 1
                End If
                reportGroup.SetDataSource(DS)
                reportGroup.SetParameterValue(reportGroup.Parameter_Period.ParameterFieldName, Format(Me.DtpFrom.Value.Date, "dd-MMM-yyyy") & " to " & Format(Me.DtpTo.Value.Date, "dd-MMM-yyyy"))
                reportGroup.SetParameterValue(reportGroup.Parameter_Scope.ParameterFieldName, "Warehouse")
                reportGroup.SetParameterValue(reportGroup.Parameter_Total_Group.ParameterFieldName, ListGroup2.Count)

                Me.crvStockCard.ReportSource = reportGroup
                Me.crvStockCard.Zoom(90)
                Me.crvStockCard.Refresh()
                Me.crvStockCard.Show()
            ElseIf Me.CboScope.SelectedIndex = 2 Then
                Me.TblData.DefaultView.RowFilter = ""
                Me.TblData.DefaultView.Sort = "[Product Code], [Location Code], [Is Begining], [Date], [Moving In] DESC, [Moving Out]"
                For Each DV As Data.DataRowView In Me.TblData.DefaultView
                    RowData = DV.Row
                    RowPrint = TblPrint.NewTable_DataRow
                    For i As Integer = 0 To TblData.Columns.Count - 1
                        RowPrint.Item(i) = RowData.Item(i)
                    Next
                    RowPrint.Group_1 = RowPrint.Product_Code 
                    RowPrint.Group_2 = RowPrint.Location_Code
                    If RowPrint.Primary_Product = LastPrimaryProduct And RowPrint.Primary_Location = LastPrimaryLocation Then
                        RowPrint.Balance = LastBalance + RowPrint.Moving_In - RowPrint.Moving_Out
                    Else
                        RowPrint.Balance = RowPrint.Moving_In
                        RowPrint.Moving_In = 0
                    End If
                    If Me.ChkShowBatch.Checked Then
                        If Not String.IsNullOrEmpty(RowPrint.Lot_No) Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Batch: " & RowPrint.Lot_No
                        End If
                        If RowPrint.Expiry_Date > #1/1/1950# Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Expiry: " & Format(RowPrint.Expiry_Date, "dd MMM yyyy")
                        End If
                    End If
                    If Me.ChkShowItemRemarks.Checked Then
                        If Not String.IsNullOrEmpty(RowPrint.Remarks) Then
                            RowPrint.Description = RowPrint.Description & vbNewLine & "Remarks: " & RowPrint.Remarks
                        End If
                    End If
                    TblPrint.Rows.Add(RowPrint)
                    LastBalance = RowPrint.Balance
                    LastPrimaryProduct = RowPrint.Primary_Product
                    LastPrimaryLocation = RowPrint.Primary_Location

                    If ListGroup2.Contains(RowPrint.Group_2) = False Then
                        ListGroup2.Add(RowPrint.Group_2)
                    End If
                Next
                If TblPrint.Count > 0 Then
                    RowPrint = TblPrint.Rows(TblPrint.Count - 1)
                    RowPrint.Last_Group = 1
                End If
                reportGroup.SetDataSource(DS)
                reportGroup.SetParameterValue(reportGroup.Parameter_Period.ParameterFieldName, Format(Me.DtpFrom.Value.Date, "dd-MMM-yyyy") & " to " & Format(Me.DtpTo.Value.Date, "dd-MMM-yyyy"))
                reportGroup.SetParameterValue(reportGroup.Parameter_Scope.ParameterFieldName, "Location")
                reportGroup.SetParameterValue(reportGroup.Parameter_Total_Group.ParameterFieldName, ListGroup2.Count)
                Me.crvStockCard.ReportSource = reportGroup
                Me.crvStockCard.Zoom(90)
                Me.crvStockCard.Refresh()
                Me.crvStockCard.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub WriteLogCard(ByVal vTblLog As dseTransaksiStockCard.Table_LogDataTable)
        Dim rowLog As dseTransaksiStockCard.Table_LogRow
        Dim index0 As Integer = 0
        Dim index1 As Integer = 0
        Try
            TblData.DefaultView.RowFilter = "[Is Begining] = -1"
            TblData.DefaultView.Sort = "[Primary Product], [Primary Location]"
            vTblLog.DefaultView.RowFilter = ""
            vTblLog.DefaultView.Sort = "[Primary Product], [Primary Location]"
            While index0 < vTblLog.DefaultView.Count
                rowLog = vTblLog.DefaultView(index0).Row
                If ListOfProduct.Contains(rowLog.Primary_Product) And ListOfLocation.Contains(rowLog.Primary_Location) Then
                    Dim Found As Boolean = False
                    While index1 < TblData.DefaultView.Count
                        RowData = TblData.DefaultView(index1).Row
                        If RowData.Primary_Product < rowLog.Primary_Product Then
                            index1 += 1
                        ElseIf RowData.Primary_Product = rowLog.Primary_Product Then
                            If RowData.Primary_Location < rowLog.Primary_Location Then
                                index1 += 1
                            ElseIf RowData.Primary_Location = rowLog.Primary_Location Then
                                Found = True
                                Exit While
                            ElseIf RowData.Primary_Location > rowLog.Primary_Location Then
                                Exit While
                            End If
                        ElseIf RowData.Primary_Product > rowLog.Primary_Product Then
                            Exit While
                        End If
                    End While
                    If Not Found Then
                        RowData = TblData.NewTable_DataRow
                        RowData.Primary_Product = rowLog.Primary_Product
                        RowData.Product_Code = rowLog.Product_Code
                        RowData.Product_Description = rowLog.Product_Description
                        RowData.Default_Precision = rowLog._Product__Precision_
                        RowData.Primary_Unit = rowLog.Primary_Product
                        RowData.Unit_Code = rowLog.Unit_Cude
                        RowData.Unit_Description = rowLog.Unit_Description
                        RowData.Primary_Location = rowLog.Primary_Location
                        If rowLog.Primary_Location = -1 Then
                            RowData.Location_Code = "No Location Controlled"
                            RowData.Location_Description = "No Location Controlled"
                        Else
                            RowData.Location_Code = rowLog.Location_Code
                            RowData.Location_Description = rowLog.Location_Description
                        End If
                        RowData.Primary_Warehouse = rowLog.Primary_Warehouse
                        If rowLog.Primary_Warehouse = -1 Then
                            RowData.Warehouse_Code = "No Location Controlled"
                            RowData.Warehouse_Description = "No Location Controlled"
                        Else
                            RowData.Warehouse_Code = rowLog.Warehouse_Code
                            RowData.Warehouse_Description = rowLog.Warehouse_Description
                        End If
                        RowData.Is_Begining = -1
                        RowData.Description = "Starting Balance"
                        RowData._Date = Me.DtpFrom.Value.Date
                        RowData.Moving_In = 0
                        RowData.Moving_Out = 0
                        RowData.Balance = 0
                        TblData.Rows.Add(RowData) 
                    End If
                    RowData = Me.TblData.NewTable_DataRow
                    RowData.Primary_Product = rowLog.Primary_Product
                    RowData.Product_Code = rowLog.Product_Code
                    RowData.Product_Description = rowLog.Product_Description
                    RowData.Default_Precision = rowLog._Product__Precision_
                    RowData.Primary_Unit = rowLog.Primary_Product
                    RowData.Unit_Code = rowLog.Unit_Cude
                    RowData.Unit_Description = rowLog.Unit_Description
                    RowData.Primary_Location = rowLog.Primary_Location
                    If rowLog.Primary_Location = -1 Then
                        RowData.Location_Code = "No Location Controlled"
                        RowData.Location_Description = "No Location Controlled"
                    Else
                        RowData.Location_Code = rowLog.Location_Code
                        RowData.Location_Description = rowLog.Location_Description
                    End If
                    RowData.Primary_Warehouse = rowLog.Primary_Warehouse
                    If rowLog.Primary_Warehouse = -1 Then
                        RowData.Warehouse_Code = "No Location Controlled"
                        RowData.Warehouse_Description = "No Location Controlled"
                    Else
                        RowData.Warehouse_Code = rowLog.Warehouse_Code
                        RowData.Warehouse_Description = rowLog.Warehouse_Description
                    End If
                    RowData.Lot_No = rowLog.Lot_No
                    RowData.Expiry_Date = rowLog.Expiry
                    RowData.Remarks = rowLog.Remarks
                    If rowLog.Qty_In > 0 Then
                        RowData.Moving_In = rowLog.Qty_In
                        RowData.Moving_Out = 0
                    Else
                        RowData.Moving_In = 0
                        RowData.Moving_Out = rowLog._Qty__Out_
                    End If
                    RowData.Description = rowLog.Description
                    RowData._Date = rowLog._Date
                    RowData.Document_No = rowLog.Document_No
                    Me.TblData.Rows.Add(RowData)
                End If
                index0 += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub CalculateBeginingBalance(ByVal vMonthlyClosingPeriod As Date, ByVal vTblClosing As dseTransaksiStockCard.tbltrans_purchase_closing_item_productqtyDataTable)
        Dim index0 As Integer = 0
        Dim index1 As Integer = 0
        Dim TblLog As New dseTransaksiStockCard.Table_LogDataTable
        Dim RowLog As dseTransaksiStockCard.Table_LogRow
        Dim RowClosing As dseTransaksiStockCard.tbltrans_purchase_closing_item_productqtyRow
        Try
            For Each RowClosing In vTblClosing
                If ListOfProduct.Contains(RowClosing.Primary_Product) And ListOfLocation.Contains(RowClosing.Primary_Location) Then
                    RowData = TblData.NewTable_DataRow
                    RowData.Primary_Product = RowClosing.Primary_Product
                    RowData.Primary_Location = RowClosing.Primary_Location
                    RowData.Is_Begining = -1
                    RowData.Description = "Starting Balance"
                    RowData._Date = Me.DtpFrom.Value.Date
                    RowData.Moving_In = RowClosing.Qty
                    RowData.Balance = RowData.Moving_In
                    TblData.Rows.Add(RowData)
                End If
            Next

            Me.FillLog(TblLog, Me.ObjCD.GetLasDateOfTheMonth(vMonthlyClosingPeriod).AddDays(1), Me.DtpFrom.Value.Date.AddDays(-1), True)
            Me.TblData.DefaultView.RowFilter = ""
            Me.TblData.DefaultView.Sort = "[Primary Product], [Primary Location]"
            TblLog.DefaultView.RowFilter = ""
            TblLog.DefaultView.Sort = "[Primary Product], [Primary Location], [Date], [Qty `Out`]"
            index0 = 0
            index1 = 0
            While index0 < TblLog.DefaultView.Count
                RowLog = TblLog.DefaultView(index0).Row
                If ListOfProduct.Contains(RowLog.Primary_Product) And ListOfLocation.Contains(RowLog.Primary_Location) Then
                    Dim Found As Boolean = False
                    While index1 < TblData.DefaultView.Count
                        RowData = TblData.DefaultView(index1).Row
                        If RowData.Primary_Product < RowLog.Primary_Product Then
                            index1 += 1
                        ElseIf RowData.Primary_Product = RowLog.Primary_Product Then
                            If RowData.Primary_Location < RowLog.Primary_Location Then
                                index1 += 1
                            ElseIf RowData.Primary_Location = RowLog.Primary_Location Then
                                RowData.Moving_In = RowData.Moving_In + RowLog.Qty_In - RowLog._Qty__Out_
                                RowData.Balance = RowData.Moving_In
                                'index1 += 1
                                Found = True
                                Exit While
                            ElseIf RowData.Primary_Location > RowLog.Primary_Location Then
                                Exit While
                            End If
                        ElseIf RowData.Primary_Product > RowLog.Primary_Product Then
                            Exit While
                        End If
                    End While
                    If Not Found Then
                        RowData = TblData.NewTable_DataRow
                        RowData.Primary_Product = RowLog.Primary_Product
                        RowData.Primary_Location = RowLog.Primary_Location
                        RowData.Is_Begining = -1
                        RowData.Description = "Starting Balance"
                        RowData._Date = Me.DtpFrom.Value.Date
                        RowData.Moving_In = RowLog.Qty_In - RowLog._Qty__Out_
                        RowData.Balance = RowData.Moving_In
                        TblData.Rows.Add(RowData)
                    End If
                End If
                index0 += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function FillLog(ByVal vTblLog As dseTransaksiStockCard.Table_LogDataTable, ByVal vFrom As Date, ByVal vTo As Date, Optional ByVal vLoadLiteData As Boolean = False) As Integer
        Dim dtaLog As New dseTransaksiStockCardTableAdapters.Table_LogTableAdapter
        Dim rowLog As dseTransaksiStockCard.Table_LogRow
        Dim dtaLogTransfer As New dseTransaksiStockCardTableAdapters.Table_Log_TransferTableAdapter
        Dim tblLogTransfer As New dseTransaksiStockCard.Table_Log_TransferDataTable
        Dim rowLogTransfer As dseTransaksiStockCard.Table_Log_TransferRow
        Try
            dtaLog.Connection.ConnectionString = Core_Module.GetConstring
            dtaLogTransfer.Connection.ConnectionString = Core_Module.GetConstring

            If Not vLoadLiteData Then
                dtaLog.Fill(vTblLog, Me.ListOfProduct(0), Me.ListOfProduct(Me.ListOfProduct.Count - 1), vFrom, vTo)
            Else
                Dim index0 As Integer = 0
                Dim index1 As Integer = 0
                Dim dtaProduct As New dseTransaksiStockCardTableAdapters.Table_ProductTableAdapter
                Dim tblProduct As New dseTransaksiStockCard.Table_ProductDataTable
                Dim rowProduct As dseTransaksiStockCard.Table_ProductRow
                Dim dtaLocation As New dseTransaksiStockCardTableAdapters.Table_LocationTableAdapter
                Dim tblLocation As New dseTransaksiStockCard.Table_LocationDataTable
                Dim rowLocation As dseTransaksiStockCard.Table_LocationRow
                Dim PrimaryLocationMin As Double = 9999999999
                Dim PrimaryLocationMax As Double = 1

                Dim dtaLogLite As New dseTransaksiStockCardTableAdapters.Table_Log_LiteTableAdapter
                Dim tblLogLite As New dseTransaksiStockCard.Table_Log_LiteDataTable
                Dim rowLogLite As dseTransaksiStockCard.Table_Log_LiteRow

                dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
                dtaLocation.Connection.ConnectionString = Core_Module.GetConstring
                dtaLogLite.Connection.ConnectionString = Core_Module.GetConstring
                dtaLogLite.ClearBeforeFill = False
                If Me.CboGroup.SelectedIndex = 0 Then
                    dtaLogLite.FillBy_PrimaryProduct(tblLogLite, vFrom, vTo, Me.txtPrimaryProduct.Text)
                ElseIf Me.CboGroup.SelectedIndex = 1 Then
                    dtaLogLite.FillBy_PrimaryCategory(tblLogLite, vFrom, vTo, Me.txtPrimaryCategory.Text)
                ElseIf Me.CboGroup.SelectedIndex = 2 Then
                    dtaLogLite.FillBy_PrimarySubBrand(tblLogLite, vFrom, vTo, Me.txtPrimarySubBrand.Text)
                ElseIf Me.CboGroup.SelectedIndex = 3 Then
                    dtaLogLite.FillBy_PrimaryBrand(tblLogLite, vFrom, vTo, Me.txtPrimaryBrand.Text)
                ElseIf Me.CboGroup.SelectedIndex = 4 Then
                    dtaLogLite.FillBy_PrimaryProductGroup(tblLogLite, vFrom, vTo, Me.txtPrimaryProductGroup.Text)
                End If
                'For Each primaryProduct_ As Double In Me.ListOfProduct
                '    dtaLogLite.FillBy_PrimaryProduct(tblLogLite, vFrom, vTo, primaryProduct_)
                'Next
                If tblLogLite.Count > 0 Then
                    For Each rowLogLite In tblLogLite
                        If Me.ListOfProduct.Contains(rowLogLite.Primary_Product) Then
                            rowLog = vTblLog.NewTable_LogRow
                            rowLog.Transaction_Type = rowLogLite.Transaction_Type
                            rowLog.Document_No = rowLogLite.Document_No
                            rowLog.Qty_In = rowLogLite.Qty_In
                            rowLog._Qty__Out_ = rowLogLite._Qty__Out_
                            rowLog.Description = rowLogLite.Description
                            rowLog._Date = rowLogLite._Date
                            rowLog.Remarks = rowLogLite.Remarks
                            rowLog.Lot_No = rowLogLite.Lot_No
                            rowLog.Expiry = rowLogLite.Expiry
                            rowLog.Primary_Product = rowLogLite.Primary_Product
                            rowLog.Primary_Location = rowLogLite.primary_location
                            vTblLog.Rows.Add(rowLog)
                            If PrimaryLocationMin > rowLogLite.primary_location Then PrimaryLocationMin = rowLogLite.primary_location
                            If PrimaryLocationMax < rowLogLite.primary_location Then PrimaryLocationMax = rowLogLite.primary_location
                        End If
                    Next
                    dtaProduct.Fill(tblProduct, Me.ListOfProduct(0), Me.ListOfProduct(Me.ListOfProduct.Count - 1))
                    dtaLocation.Fill(tblLocation, PrimaryLocationMin, PrimaryLocationMax)

                    vTblLog.DefaultView.RowFilter = ""
                    vTblLog.DefaultView.Sort = "[Primary Product]"
                    tblProduct.DefaultView.RowFilter = ""
                    tblProduct.DefaultView.Sort = "[Primary Product]"
                    While index0 < vTblLog.DefaultView.Count
                        rowLog = vTblLog.DefaultView(index0).Row
                        While index1 < tblProduct.DefaultView.Count
                            rowProduct = tblProduct.DefaultView(index1).Row
                            If rowProduct.Primary_Product < rowLog.Primary_Product Then
                                index1 += 1
                            ElseIf rowProduct.Primary_Product = rowLog.Primary_Product Then
                                rowLog.Product_Code = rowProduct.Product_Code
                                rowLog.Product_Description = rowProduct.Product_Description
                                rowLog._Product__Precision_ = rowProduct._Default__Precision_
                                rowLog.Primary_Unit = rowProduct.Primary_Unit
                                rowLog.Unit_Cude = rowProduct.Unit_Code
                                rowLog.Unit_Description = rowProduct.Unit_Description
                                Exit While
                            ElseIf rowProduct.Primary_Product > rowLog.Primary_Product Then
                                MsgBox("Logic Error: Class:FrmTransaksiStockCard2, Function: FillLog" & vbNewLine & "Desc: Product not found", MsgBoxStyle.Information, "Information")
                                Exit While
                            End If
                        End While
                        index0 += 1
                    End While
                    index0 = 0
                    index1 = 0
                    vTblLog.DefaultView.RowFilter = ""
                    vTblLog.DefaultView.Sort = "[Primary Location]"
                    tblLocation.DefaultView.RowFilter = ""
                    tblLocation.DefaultView.Sort = "[Primary Location]"
                    While index0 < vTblLog.DefaultView.Count
                        rowLog = vTblLog.DefaultView(index0).Row
                        If rowLog.Primary_Location <> -1 Then
                            While index1 < tblLocation.DefaultView.Count
                                rowLocation = tblLocation.DefaultView(index1).Row
                                If rowLocation.Primary_Location < rowLog.Primary_Location Then
                                    index1 += 1
                                ElseIf rowLocation.Primary_Location = rowLog.Primary_Location Then
                                    rowLog.Location_Code = rowLocation.Location_Code
                                    rowLog.Location_Description = rowLocation.Location_Description
                                    rowLog.Primary_Warehouse = rowLocation.Primary_Warehouse
                                    rowLog.Warehouse_Code = rowLocation.Warehouse_Code
                                    rowLog.Warehouse_Description = rowLocation.Warehouse_Description
                                    Exit While
                                ElseIf rowLocation.Primary_Location > rowLog.Primary_Location Then
                                    MsgBox("Logic Error: Class:FrmTransaksiStockCard2, Funct: FillLog" & vbNewLine & "Desc: Location not found", MsgBoxStyle.Information, "Information")
                                    Exit While
                                End If
                            End While
                        End If
                        index0 += 1
                    End While

                End If
            End If


            dtaLogTransfer.Fill(tblLogTransfer, Me.ListOfProduct(0), Me.ListOfProduct(Me.ListOfProduct.Count - 1), vFrom, vTo)
            For Each rowLogTransfer In tblLogTransfer
                If Me.ListOfProduct.Contains(rowLogTransfer.Primary_Product) Then
                    rowLog = vTblLog.NewTable_LogRow
                    rowLog.Primary_Product = rowLogTransfer.Primary_Product
                    rowLog.Product_Code = rowLogTransfer.Product_Code
                    rowLog.Product_Description = rowLogTransfer.Product_Description
                    rowLog._Product__Precision_ = rowLogTransfer._Product__Precision_
                    rowLog.Primary_Unit = rowLogTransfer.Primary_Unit
                    rowLog.Unit_Cude = rowLogTransfer.Unit_Code
                    rowLog.Unit_Description = rowLogTransfer.Unit_Description
                    rowLog.Primary_Location = rowLogTransfer.Primary_Location_Destination
                    rowLog.Location_Code = rowLogTransfer.Location_Code_Desctination
                    rowLog.Location_Description = rowLogTransfer.Location_Description_Destination
                    rowLog.Primary_Warehouse = rowLogTransfer.Primary_Warehouse_Destination
                    rowLog.Warehouse_Code = rowLogTransfer.Warehouse_Code_Destination
                    rowLog.Warehouse_Description = rowLogTransfer.Warehouse_Description_Destination
                    rowLog.Transaction_Type = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Transfer
                    rowLog.Document_No = rowLogTransfer.Document_No
                    rowLog.Qty_In = rowLogTransfer.Qty
                    rowLog._Qty__Out_ = 0
                    rowLog.Description = "Stock transfer from " & rowLogTransfer.Location_Code
                    rowLog._Date = rowLogTransfer._Date
                    rowLog.Remarks = rowLogTransfer.Remarks
                    rowLog.Lot_No = rowLogTransfer.Lot_No
                    rowLog.Expiry = rowLogTransfer.Expiry
                    vTblLog.Rows.Add(rowLog)
                    '-
                    rowLog = vTblLog.NewTable_LogRow
                    rowLog.Primary_Product = rowLogTransfer.Primary_Product
                    rowLog.Product_Code = rowLogTransfer.Product_Code
                    rowLog.Product_Description = rowLogTransfer.Product_Description
                    rowLog._Product__Precision_ = rowLogTransfer._Product__Precision_
                    rowLog.Primary_Unit = rowLogTransfer.Primary_Unit
                    rowLog.Unit_Cude = rowLogTransfer.Unit_Code
                    rowLog.Unit_Description = rowLogTransfer.Unit_Description
                    rowLog.Primary_Location = rowLogTransfer.Primary_Location
                    rowLog.Location_Code = rowLogTransfer.Location_Code
                    rowLog.Location_Description = rowLogTransfer.Location_Description
                    rowLog.Primary_Warehouse = rowLogTransfer.Primary_Warehouse
                    rowLog.Warehouse_Code = rowLogTransfer.Warehouse_Code
                    rowLog.Warehouse_Description = rowLogTransfer.Warehouse_Description
                    rowLog.Transaction_Type = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Transfer
                    rowLog.Document_No = rowLogTransfer.Document_No
                    rowLog.Qty_In = 0
                    rowLog._Qty__Out_ = rowLogTransfer.Qty
                    rowLog.Description = "Stock transfer to " & rowLogTransfer.Location_Code_Desctination
                    rowLog._Date = rowLogTransfer._Date
                    rowLog.Remarks = rowLogTransfer.Remarks
                    rowLog.Lot_No = rowLogTransfer.Lot_No
                    rowLog.Expiry = rowLogTransfer.Expiry
                    vTblLog.Rows.Add(rowLog)

                End If
            Next
            Return vTblLog.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function 

    Private Function GetClosingPeriod(ByVal vPrimaryClosingPeriod As Double) As Date
        Dim Dta As New dseTransaksiStockCardTableAdapters.DataTable1TableAdapter
        Dim val As New Nullable(Of Date)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            val = Dta.ScalarQuery_GetClosingPeriod(vPrimaryClosingPeriod)
            If val.HasValue Then
                Return val.Value
            End If
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Private Function GetPrimaryMonthlyClosing(ByVal vRefDate As Date) As Double
        Dim Dta As New dseTransaksiStockCardTableAdapters.DataTable1TableAdapter
        Dim Tbl As New dseTransaksiStockCard.DataTable1DataTable
        Dim Row As dseTransaksiStockCard.DataTable1Row
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring 
            If Dta.Fill(Tbl, Me.ObjCD.GetFirstDateofTheMonth(vRefDate.Date), 1) > 0 Then
                Row = Tbl.Rows(0)
                Return Row.primary
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub SetlocationList()
        Dim Dta As New dseTransaksiStockCardTableAdapters.tblmaster_storage_areaTableAdapter
        Dim Tbl As New dseTransaksiStockCard.tblmaster_storage_areaDataTable
        Dim Row As dseTransaksiStockCard.tblmaster_storage_areaRow
        Try
            Me.ListOfLocation.Clear()
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Me.DgvLocation.AllowUserToAddRows = False
            For Each RowDgv In Me.DgvLocation.Rows
                If RowDgv.Cells(Me.Col1Level.Index).Value = "Location" Then
                    If Me.ListOfLocation.Contains(RowDgv.Cells(Me.Col1Primary.Index).Value) = False Then
                        Me.ListOfLocation.Add(RowDgv.Cells(Me.Col1Primary.Index).Value)
                    End If
                ElseIf RowDgv.Cells(Me.Col1Level.Index).Value = "Warehouse" Then
                    Dta.Fill(Tbl, RowDgv.Cells(Me.Col1Primary.Index).Value)
                    For Each Row In Tbl
                        If Me.ListOfLocation.Contains(Row.Primary) = False Then
                            Me.ListOfLocation.Add(Row.Primary)
                        End If
                    Next
                End If
            Next
            If Me.ListOfLocation.Count = 0 Then
                Dta.FillAll(Tbl)
                For Each Row In Tbl
                    Me.ListOfLocation.Add(Row.Primary)
                Next
            End If
            Me.ListOfLocation.Add(-1)
            Me.ListOfLocation.Sort()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvLocation.AllowUserToAddRows = True
        End Try
    End Sub

    Private Sub SetProductList()
        Dim dtaProduct As New dseTransaksiStockCardTableAdapters.tblmaster_productTableAdapter
        Dim tblProduct As New dseTransaksiStockCard.tblmaster_productDataTable
        Dim rowProduct As dseTransaksiStockCard.tblmaster_productRow
        Try
            Me.ListOfProduct.Clear()
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            If Me.CboGroup.SelectedIndex = 0 Then
                If Me.ListOfProduct.Contains(Me.txtPrimaryProduct.Text) = False Then
                    Me.ListOfProduct.Add(Me.txtPrimaryProduct.Text)
                End If
            ElseIf Me.CboGroup.SelectedIndex = 1 Then
                dtaProduct.FillBy_PrimaryCategory(tblProduct, Me.txtPrimaryCategory.Text)
            ElseIf Me.CboGroup.SelectedIndex = 2 Then
                dtaProduct.FillBy_PrimarySubBrand(tblProduct, Me.txtPrimarySubBrand.Text)
            ElseIf Me.CboGroup.SelectedIndex = 3 Then
                dtaProduct.FillBy_PrimaryBrand(tblProduct, Me.txtPrimaryBrand.Text)
            ElseIf Me.CboGroup.SelectedIndex = 4 Then
                dtaProduct.FillBy_ProductGroup(tblProduct, Me.txtPrimaryProductGroup.Text)
            End If
            For Each rowProduct In tblProduct
                If Me.ListOfProduct.Contains(rowProduct.Primary) = False Then
                    Me.ListOfProduct.Add(rowProduct.Primary)
                End If
            Next
            Me.ListOfProduct.Sort()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim _FrmTransaksiStockCard2_UpdateMonthlyClosing As New FrmTransaksiStockCard2_UpdateMonthlyClosing
        Try
            _FrmTransaksiStockCard2_UpdateMonthlyClosing.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        'Dim _FrmTransaksiStockCard2_GetStock As New FrmTransaksiStockCard2_GetStock
        'Try
        '    _FrmTransaksiStockCard2_GetStock.ShowDialog()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub
End Class