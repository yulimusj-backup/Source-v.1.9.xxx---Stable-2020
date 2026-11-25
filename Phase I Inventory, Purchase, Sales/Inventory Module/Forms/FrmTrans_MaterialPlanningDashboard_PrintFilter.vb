Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_PrintFilter
    Dim objPG As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory

    Private Sub chkMP_RequiredOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMP_RequiredOnly.CheckedChanged
        Try
            If Me.chkMP_RequiredOnly.Checked = True Then
                Me.dtpMP_Filter.Enabled = True
            Else
                Me.dtpMP_Filter.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.chkFilter_Goods.Checked = False And Me.chkFilter_Service.Checked = False Then
                Me.chkFilter_Goods.Checked = True
            End If
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtFilter_ProductGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_ProductGroupCode.KeyDown
        Dim enumAktif As FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.cboFilter_ProductGroupLevel.SelectedIndex <> 3 Then
                    If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                        enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.ProductGroup
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                        enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Brand
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                        enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.SubBrand
                    End If
                    If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                        If Me.objPG.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text, enumAktif) = True Then
                            Me.TxtFilter_PrimaryProductGroup.Text = Me.objPG.GetPrimary(Me.TxtFilter_ProductGroupCode.Text, enumAktif)
                            Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objPG.GetGroupDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                        Else
                            Me.btnFilter_FindProductGroupLevel.PerformClick()
                        End If
                    Else
                        Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                        Me.TxtFilter_ProductGroupLevelDesc.Clear()
                    End If
                Else
                    If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                        If Me.objCategory.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text) = True Then
                            Me.TxtFilter_PrimaryProductGroup.Text = Me.objCategory.GetPrimary(Me.TxtFilter_ProductGroupCode.Text)
                            Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objCategory.GetDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                        Else
                            Me.btnFilter_FindProductGroupLevel.PerformClick()
                        End If
                    Else
                        Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                        Me.TxtFilter_ProductGroupLevelDesc.Clear()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFilter_FindProductGroupLevel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindProductGroupLevel.Click
        Dim frm As New Core_Module.FrmCariMasterProductGroup
        If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.ProductGroup
        ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Brand
        ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.SubBrand
        ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 3 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Category
        End If
        Dim frmCat As New Core_Module.frmCariMasterCategory

        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.cboFilter_ProductGroupLevel.SelectedIndex <> 3 Then
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    frm.txtCari.Text = Me.TxtFilter_ProductGroupCode.Text
                End If
                frm.ShowDialog()
                If frm.PrimaryProductGroup > 0 Then
                    Me.TxtFilter_PrimaryProductGroup.Text = frm.PrimaryProductGroup
                    Me.TxtFilter_ProductGroupCode.Text = Me.objPG.GetGroupCode(frm.PrimaryProductGroup)
                    Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objPG.GetGroupDesc(frm.PrimaryProductGroup)
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupCode.Clear()
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                End If
            Else
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    frmCat.txtCari.Text = Me.TxtFilter_ProductGroupCode.Text
                End If
                frmCat.ShowDialog()
                If frmCat.PrimaryCategory > 0 Then
                    Me.TxtFilter_PrimaryProductGroup.Text = frmCat.PrimaryCategory
                    Me.TxtFilter_ProductGroupCode.Text = Me.objCategory.GetCategoryCode(frmCat.PrimaryCategory)
                    Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objCategory.GetDesc(frmCat.PrimaryCategory)
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupCode.Clear()
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
            frmCat.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboFilter_ProductGroupLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFilter_ProductGroupLevel.SelectedIndexChanged
        Try
            Me.TxtFilter_PrimaryProductGroup.Text = "-1"
            Me.TxtFilter_ProductGroupCode.Clear()
            Me.TxtFilter_ProductGroupLevelDesc.Clear()

            If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                Me.LblFilter_ProductGroup.Text = "Group Code"
                Me.LblFilter_PoductGroupDesc.Text = "Group Description"
            ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                Me.LblFilter_ProductGroup.Text = "Brand Code"
                Me.LblFilter_PoductGroupDesc.Text = "Brand Description"
            ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                Me.LblFilter_ProductGroup.Text = "Sub Brand Code"
                Me.LblFilter_PoductGroupDesc.Text = "Sub Brand Description"
            ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 3 Then
                Me.LblFilter_ProductGroup.Text = "Category Code"
                Me.LblFilter_PoductGroupDesc.Text = "Catgeory Description"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub TxtFilter_ProductGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_ProductGroupCode.Leave
        Dim enumAktif As FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.cboFilter_ProductGroupLevel.SelectedIndex <> 3 Then
                If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                    enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.ProductGroup
                ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                    enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Brand
                ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                    enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.SubBrand
                End If
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    If Me.objPG.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text, enumAktif) = True Then
                        Me.TxtFilter_PrimaryProductGroup.Text = Me.objPG.GetPrimary(Me.TxtFilter_ProductGroupCode.Text, enumAktif)
                        Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objPG.GetGroupDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                    Else
                        Me.btnFilter_FindProductGroupLevel.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                End If
            Else
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    If Me.objCategory.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text) = True Then
                        Me.TxtFilter_PrimaryProductGroup.Text = Me.objCategory.GetPrimary(Me.TxtFilter_ProductGroupCode.Text)
                        Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objCategory.GetDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                    Else
                        Me.btnFilter_FindProductGroupLevel.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_PrintFilter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.cboFilter_ProductGroupLevel.SelectedIndex = 0
        Me.cboProductDetail.SelectedIndex = 0
        Me.cboGrouping.SelectedIndex = 0
    End Sub
End Class