'''==by @rosadinnn=='''

Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_Data
    Dim ObjUser As New Core_Module.ClsUser

#Region "init"

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim buffer As String = ""
    Dim buffer1 As String = ""

    Dim objSupp As New Core_Module.clsMasterSupplier
    Dim objPG As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim obj As New ClsTransaksiMaterialPlanning_Data
    Dim UserCode As String = "USER"
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim CheckForRecalculateAll As Boolean = True

    Public Enum EnumProductGroupLevel As Integer
        Category
        SubBrand
        Brand
        ProductGroup
    End Enum
    Dim indexRowBegin As Integer = -1

#End Region

#Region "common"

    Private Sub FrmTransaksiReq_Purchase_Data_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Dummy As String = ""
        Try
            'jika tidak memiliki hak akses maka tutup link label material grouping
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_Data1") = True Then
                Me.llbMaterialPlanningGroup.Enabled = True
            Else
                Me.llbMaterialPlanningGroup.Enabled = False
            End If

            Me.Cursor = Cursors.WaitCursor
            Me.dtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Me.dtpTo.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpTo.Value)
            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan)
            If Dummy = "" Then
                Me.txtWorkDays.Text = "0"
            Else
                Me.txtWorkDays.Text = Dummy
            End If

            ''Init
            Me.cboParameter.SelectedIndex = 1
            Me.cboProductGroup.SelectedIndex = 0
            Me.CboFilterProcurementStrategy.SelectedIndex = 0
            Me.ObjCD.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, False, False, True, False, False, False)
            Me.ObjCD.KendalikanControl(Me.GroupBox2.Controls, True, False, True, False, False, False, True, False, False, False)
            Me.ObjCD.KendalikanControl(Me.GroupBox3.Controls, True, False, True, False, False, False, True, False, False, False)
            Me.cboSuppMode.SelectedIndex = 0
            Me.cboOrderQtyMode.SelectedIndex = 0
            Me.cboProcurementStrategy.SelectedIndex = 0
            Me.cboProcurementType.SelectedIndex = 0
            Me.cboAllIncludeAdjustment.SelectedIndex = 0
            Me.cboAllIncludeBundling.SelectedIndex = 0
            Me.cboAllIncludeSales.SelectedIndex = 0
            Me.cboAllSafetyStock.SelectedIndex = 0
            Me.cboAllLeadTime.SelectedIndex = 0
            Me.cboAllSafetyStockParams.SelectedIndex = 0
            Me.cboAllLeadTimeParams.SelectedIndex = 0
            Me.cboAllOptimumStockParams.SelectedIndex = 0
            Me.cboAllOptimumStock.SelectedIndex = 0

            Me.Prg1.Visible = False
            Me.LblProgress.Visible = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboParameter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboParameter.SelectedIndexChanged
        Try
            Me.Cursor = Cursors.WaitCursor

            ''blok param 1
            Me.ColProcurementType.Visible = False
            Me.ColProcurementStrategy.Visible = False
            Me.chkProcurementType.Visible = False
            Me.chkProcurementType.TabStop = False
            Me.cboProcurementType.Visible = False
            Me.cboProcurementType.TabStop = False
            Me.ColSupplierSelectionMode.Visible = False
            Me.ColPreferedSupplier.Visible = False
            Me.ColRecomendedOrderQtyType.Visible = False
            Me.GroupBox1.Visible = False
            Me.GroupBox5.Visible = False

            ''blok param 2
            Me.ColAverageMonthlyDemand.Visible = False
            Me.ColStartPeriod.Visible = False
            Me.ColEndPeriod.Visible = False
            Me.ColIncludeSales.Visible = False
            Me.ColIncludeBundling.Visible = False
            Me.ColIncludeAdjustment.Visible = False

            Me.ColSafetyStockStrategy.Visible = False
            Me.ColSafetyStockDays.Visible = False
            Me.ColSafetyStockQty.Visible = False

            Me.ColLeadTimeStrategy.Visible = False
            Me.ColLeadTimeDays.Visible = False
            Me.ColLeadTimeQty.Visible = False

            Me.ColOptStockStrategy.Visible = False
            Me.ColOptStockDays.Visible = False
            Me.ColOptStockQty.Visible = False

            Me.ColExplainAverageQty.Visible = False
            Me.ColExplainSafetyStockQty.Visible = False
            Me.ColExplainLeadTimeQty.Visible = False
            Me.GroupBox2.Visible = False

            ''blok param 3
            Me.ColExclude.Visible = False

            ''blok param 4
            Me.ColMinStockQty.Visible = False
            Me.ColMaxStockQty.Visible = False
            Me.ColMinOrderQty.Visible = False
            Me.ColOptOrderQty.Visible = False
            Me.ColMaxOrderQty.Visible = False
            Me.GroupBox3.Visible = False

            Me.SplitContainer2.SplitterDistance = 0
            Me.btnApply.Visible = False

            ''case
            If Me.cboParameter.SelectedIndex = 0 Then
                Me.ColProcurementType.Visible = True
                Me.ColProcurementStrategy.Visible = True
                Me.ColSupplierSelectionMode.Visible = True
                Me.ColPreferedSupplier.Visible = True
                'Me.ColRecomendedOrderQtyType.Visible = True
                Me.GroupBox1.Visible = True
                Console.WriteLine(Me.GroupBox1.Visible.ToString)
                Me.SplitContainer2.Panel1Collapsed = False
                Me.SplitContainer2.SplitterDistance = 100
                Me.btnApply.Visible = True
                Me.cboOrderQtyMode.Visible = False
                Me.chkOrderQtyTo.Visible = False
                Me.cboOrderQtyMode.TabStop = False
                Me.chkOrderQtyTo.TabStop = False
                Me.ColDef_LowerTolerancePerct.Visible = True
                Me.ColDef_UpperTolerancePerct.Visible = True
                Me.chkProcurementType.Visible = True
                Me.chkProcurementType.TabStop = True
                Me.cboProcurementType.Visible = True
                Me.cboProcurementType.TabStop = True
            ElseIf Me.cboParameter.SelectedIndex = 1 Then
                Me.ColAverageMonthlyDemand.Visible = True
                Me.ColStartPeriod.Visible = True
                Me.ColEndPeriod.Visible = True
                Me.ColSafetyStockDays.Visible = True
                Me.ColSafetyStockQty.Visible = True
                Me.ColLeadTimeDays.Visible = True
                Me.ColLeadTimeQty.Visible = True

                Me.ColOptStockDays.Visible = True
                Me.ColOptStockQty.Visible = True

                Me.GroupBox2.Visible = True
                Me.SplitContainer2.Panel1Collapsed = False
                Me.SplitContainer2.SplitterDistance = 160
                Me.btnApply.Visible = True
                Me.GroupBox5.Visible = True
                Me.ColDef_LowerTolerancePerct.Visible = False
                Me.ColDef_UpperTolerancePerct.Visible = False
            ElseIf Me.cboParameter.SelectedIndex = 2 Then
                Me.SplitContainer2.Panel1Collapsed = True
                Me.ColExclude.Visible = True
                Me.ColDef_LowerTolerancePerct.Visible = False
                Me.ColDef_UpperTolerancePerct.Visible = False
            ElseIf Me.cboParameter.SelectedIndex = 3 Then
                'Me.ColMinStockQty.Visible = True
                Me.ColOptStockQty.Visible = True
                'Me.ColMaxStockQty.Visible = True
                'Me.ColMinOrderQty.Visible = True
                Me.ColOptOrderQty.Visible = True
                'Me.ColMaxOrderQty.Visible = True
                Me.ColAverageMonthlyDemand.Visible = True
                Me.GroupBox3.Visible = True
                Me.SplitContainer2.Panel1Collapsed = False
                Me.SplitContainer2.SplitterDistance = 100
                Me.btnApply.Visible = True
                Me.ColRecomendedOrderQtyType.Visible = True
                Me.ColSafetyStockQty.Visible = True
                Me.ColLeadTimeQty.Visible = True
                Me.ColDef_LowerTolerancePerct.Visible = False
                Me.ColDef_UpperTolerancePerct.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

#End Region

#Region "data filtering"

    Private Sub txtSuppCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSuppCode.KeyDown
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.txtSuppCode.Text <> "" Then
                    If Me.objSupp.ApakahSudahAda(Me.txtSuppCode.Text) = True Then
                        Me.txtPrimarySupp.Text = Me.objSupp.GetPrimary(Me.txtSuppCode.Text)
                        Me.txtSuppCode.Text = Me.objSupp.GetSupplierCode(Me.txtPrimarySupp.Text)
                        Me.txtSuppDesc.Text = Me.objSupp.GetSupplierName(Me.txtPrimarySupp.Text)
                    Else
                        Me.btnFindSupp.PerformClick()
                    End If
                Else
                    Me.txtPrimarySupp.Text = "-1"
                    Me.txtSuppCode.Clear()
                    Me.txtSuppDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFindSupp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupp.Click
        Dim frm As New Core_Module.FrmCariMasterSupplier

        Try
            If Me.txtSuppCode.Text <> "" Then
                frm.txtCari.Text = Me.txtSuppCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimarySupplier > 0 Then
                Me.txtPrimarySupp.Text = frm.PrimarySupplier
                Me.txtSuppCode.Text = Me.objSupp.GetSupplierCode(frm.PrimarySupplier)
                Me.txtSuppDesc.Text = Me.objSupp.GetSupplierName(frm.PrimarySupplier)
            Else
                Me.txtPrimarySupp.Text = "-1"
                Me.txtSuppCode.Clear()
                Me.txtSuppDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub txtSuppCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuppCode.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtSuppCode.Text <> "" Then
                If Me.objSupp.ApakahSudahAda(Me.txtSuppCode.Text) = True Then
                    Me.txtPrimarySupp.Text = Me.objSupp.GetPrimary(Me.txtSuppCode.Text)
                    Me.txtSuppCode.Text = Me.objSupp.GetSupplierCode(Me.txtPrimarySupp.Text)
                    Me.txtSuppDesc.Text = Me.objSupp.GetSupplierName(Me.txtPrimarySupp.Text)
                Else
                    Me.btnFindSupp.PerformClick()
                End If
            Else
                Me.txtPrimarySupp.Text = "-1"
                Me.txtSuppCode.Clear()
                Me.txtSuppDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGroupCode.KeyDown
        Dim enumAktif As FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.cboProductGroup.SelectedIndex <> 3 Then
                    If Me.cboProductGroup.SelectedIndex = 0 Then
                        enumAktif = EnumProductGroupLevel.ProductGroup
                    ElseIf Me.cboProductGroup.SelectedIndex = 1 Then
                        enumAktif = EnumProductGroupLevel.Brand
                    ElseIf Me.cboProductGroup.SelectedIndex = 2 Then
                        enumAktif = EnumProductGroupLevel.SubBrand
                    End If
                    If Me.txtGroupCode.Text <> "" Then
                        If Me.objPG.ApakahSudahAda(Me.txtGroupCode.Text, enumAktif) = True Then
                            Me.TxtPrimaryGrop.Text = Me.objPG.GetPrimary(Me.txtGroupCode.Text, enumAktif)
                            Me.txtGroupDesc.Text = Me.objPG.GetGroupDesc(Me.TxtPrimaryGrop.Text)
                        Else
                            Me.btnFindGroupLevel.PerformClick()
                        End If
                    Else
                        Me.TxtPrimaryGrop.Text = "-1"
                        Me.txtGroupDesc.Clear()
                    End If
                Else
                    If Me.txtGroupCode.Text <> "" Then
                        If Me.objCategory.ApakahSudahAda(Me.txtGroupCode.Text) = True Then
                            Me.TxtPrimaryGrop.Text = Me.objCategory.GetPrimary(Me.txtGroupCode.Text)
                            Me.txtGroupDesc.Text = Me.objCategory.GetDesc(Me.TxtPrimaryGrop.Text)
                        Else
                            Me.btnFindGroupLevel.PerformClick()
                        End If
                    Else
                        Me.TxtPrimaryGrop.Text = "-1"
                        Me.txtGroupDesc.Clear()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFindGroupLevel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindGroupLevel.Click
        Dim frm As New Core_Module.FrmCariMasterProductGroup
        If Me.cboProductGroup.SelectedIndex = 0 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.ProductGroup
        ElseIf Me.cboProductGroup.SelectedIndex = 1 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Brand
        ElseIf Me.cboProductGroup.SelectedIndex = 2 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.SubBrand
        ElseIf Me.cboProductGroup.SelectedIndex = 3 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Category
        End If
        Dim frmCat As New Core_Module.frmCariMasterCategory
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.cboProductGroup.SelectedIndex <> 3 Then
                If Me.txtGroupCode.Text <> "" Then
                    frm.txtCari.Text = Me.txtGroupCode.Text
                End If
                frm.ShowDialog()
                If frm.PrimaryProductGroup > 0 Then
                    Me.TxtPrimaryGrop.Text = frm.PrimaryProductGroup
                    Me.txtGroupCode.Text = Me.objPG.GetGroupCode(frm.PrimaryProductGroup)
                    Me.txtGroupDesc.Text = Me.objPG.GetGroupDesc(frm.PrimaryProductGroup)
                Else
                    Me.TxtPrimaryGrop.Text = "-1"
                    Me.txtGroupDesc.Clear()
                    Me.txtGroupCode.Clear()
                End If
            Else
                If Me.txtGroupCode.Text <> "" Then
                    frmCat.txtCari.Text = Me.txtGroupCode.Text
                End If
                frmCat.ShowDialog()
                If frmCat.PrimaryCategory > 0 Then
                    Me.TxtPrimaryGrop.Text = frmCat.PrimaryCategory
                    Me.txtGroupCode.Text = Me.objCategory.GetCategoryCode(frmCat.PrimaryCategory)
                    Me.txtGroupDesc.Text = Me.objCategory.GetDesc(frmCat.PrimaryCategory)
                Else
                    Me.TxtPrimaryGrop.Text = "-1"
                    Me.txtGroupCode.Clear()
                    Me.txtGroupDesc.Clear()
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

    Private Sub cboProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProductGroup.SelectedIndexChanged
        Try
            Me.TxtPrimaryGrop.Text = "-1"
            Me.txtGroupCode.Clear()
            Me.txtGroupDesc.Clear()

            If Me.cboProductGroup.SelectedIndex = 0 Then
                Me.LblProductGroup.Text = "Group Code"
                Me.LblPoductGroupDesc.Text = "Group Description"
            ElseIf Me.cboProductGroup.SelectedIndex = 1 Then
                Me.LblProductGroup.Text = "Brand Code"
                Me.LblPoductGroupDesc.Text = "Brand Description"
            ElseIf Me.cboProductGroup.SelectedIndex = 2 Then
                Me.LblProductGroup.Text = "Sub Brand Code"
                Me.LblPoductGroupDesc.Text = "Sub Brand Description"
            ElseIf Me.cboProductGroup.SelectedIndex = 3 Then
                Me.LblProductGroup.Text = "Category Code"
                Me.LblPoductGroupDesc.Text = "Catgeory Description"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.Leave
        Dim enumAktif As FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.cboProductGroup.SelectedIndex <> 3 Then
                If Me.cboProductGroup.SelectedIndex = 0 Then
                    enumAktif = EnumProductGroupLevel.ProductGroup
                ElseIf Me.cboProductGroup.SelectedIndex = 1 Then
                    enumAktif = EnumProductGroupLevel.Brand
                ElseIf Me.cboProductGroup.SelectedIndex = 2 Then
                    enumAktif = EnumProductGroupLevel.SubBrand
                End If
                If Me.txtGroupCode.Text <> "" Then
                    If Me.objPG.ApakahSudahAda(Me.txtGroupCode.Text, enumAktif) = True Then
                        Me.TxtPrimaryGrop.Text = Me.objPG.GetPrimary(Me.txtGroupCode.Text, enumAktif)
                        Me.txtGroupDesc.Text = Me.objPG.GetGroupDesc(Me.TxtPrimaryGrop.Text)
                    Else
                        Me.btnFindGroupLevel.PerformClick()
                    End If
                Else
                    Me.TxtPrimaryGrop.Text = "-1"
                    Me.txtGroupDesc.Clear()
                End If
            Else
                If Me.txtGroupCode.Text <> "" Then
                    If Me.objCategory.ApakahSudahAda(Me.txtGroupCode.Text) = True Then
                        Me.TxtPrimaryGrop.Text = Me.objCategory.GetPrimary(Me.txtGroupCode.Text)
                        Me.txtGroupDesc.Text = Me.objCategory.GetDesc(Me.TxtPrimaryGrop.Text)
                    Else
                        Me.btnFindGroupLevel.PerformClick()
                    End If
                Else
                    Me.TxtPrimaryGrop.Text = "-1"
                    Me.txtGroupDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim filterType As New ClsTransaksiMaterialPlanning_Data.EnumPencarian
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtPrimarySupp.Text = "-1" And Me.TxtPrimaryGrop.Text = "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDesc
            ElseIf Me.txtPrimarySupp.Text <> "-1" And Me.TxtPrimaryGrop.Text = "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescSupp
            ElseIf Me.txtPrimarySupp.Text = "-1" And Me.cboProductGroup.SelectedIndex = 3 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescCat
            ElseIf Me.txtPrimarySupp.Text = "-1" And Me.cboProductGroup.SelectedIndex = 2 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescSubbrand
            ElseIf Me.txtPrimarySupp.Text = "-1" And Me.cboProductGroup.SelectedIndex = 1 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescBrand
            ElseIf Me.txtPrimarySupp.Text = "-1" And Me.cboProductGroup.SelectedIndex = 0 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescProductgroup
            ElseIf Me.txtPrimarySupp.Text <> "-1" And Me.cboProductGroup.SelectedIndex = 3 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescSuppCat
            ElseIf Me.txtPrimarySupp.Text <> "-1" And Me.cboProductGroup.SelectedIndex = 2 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescSuppSubbrand
            ElseIf Me.txtPrimarySupp.Text <> "-1" And Me.cboProductGroup.SelectedIndex = 1 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescSuppBrand
            ElseIf Me.txtPrimarySupp.Text <> "-1" And Me.cboProductGroup.SelectedIndex = 0 And Me.TxtPrimaryGrop.Text <> "-1" Then
                filterType = ClsTransaksiMaterialPlanning_Data.EnumPencarian.CodeDescSuppProductGroup
            End If
            Me.obj.filterData(filterType, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

#End Region

    Private Sub btnFindPreferedSupp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPreferedSupp.Click
        Dim frm As New Core_Module.FrmCariMasterSupplier
        Try
            If Me.txtPreferedSuppCode.Text <> "" Then
                frm.txtCari.Text = Me.txtPreferedSuppCode.Text
            End If
            frm.ShowDialog()
            If frm.PrimarySupplier > 0 Then
                Me.txtPrimaryPreferedSupp.Text = frm.PrimarySupplier
                Me.txtPreferedSuppCode.Text = Me.objSupp.GetSupplierCode(frm.PrimarySupplier)
            Else
                Me.txtPrimaryPreferedSupp.Text = "-1"
                Me.txtPreferedSuppCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtPreferedSuppCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPreferedSuppCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                If Me.txtPreferedSuppCode.Text <> "" Then
                    If Me.objSupp.ApakahSudahAda(Me.txtPreferedSuppCode.Text) = True Then
                        Me.txtPrimaryPreferedSupp.Text = Me.objSupp.GetPrimary(Me.txtPreferedSuppCode.Text)
                        Me.txtPreferedSuppCode.Text = Me.objSupp.GetSupplierCode(Me.txtPrimaryPreferedSupp.Text)
                    Else
                        Me.btnFindPreferedSupp.PerformClick()
                    End If
                Else
                    Me.txtPrimaryPreferedSupp.Text = "-1"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub txtPreferedSuppCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPreferedSuppCode.Leave
        Try
            If Me.txtPreferedSuppCode.Text <> "" Then
                If Me.objSupp.ApakahSudahAda(Me.txtPreferedSuppCode.Text) = True Then
                    Me.txtPrimaryPreferedSupp.Text = Me.objSupp.GetPrimary(Me.txtPreferedSuppCode.Text)
                    Me.txtPreferedSuppCode.Text = Me.objSupp.GetSupplierCode(Me.txtPrimaryPreferedSupp.Text)
                Else
                    Me.btnFindPreferedSupp.PerformClick()
                End If
            Else
                Me.txtPrimaryPreferedSupp.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub ReCalculateAvgSafetystockAndLeadTime(ByVal vRowIndex As Integer)
        Dim rowdgv As DataGridViewRow
        Dim tblHistorical As New DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable
        Dim dtaB As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            dtaB.Connection.ConnectionString = Core_Module.GetConstring
            dtaD.Connection.ConnectionString = dtaB.Connection.ConnectionString

            rowdgv = Me.dgvItem.Rows(vRowIndex)
            ''hitung avg dkk::
            ''get dulu historicalnya::
            Me.obj.GetHistoricalByPrimary(tblHistorical, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing))
            ''hitung avg monthly
            Me.obj.CalculateAverageMonthlyDemandByPrimary(rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                    Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColIncludeSales.Index).Value, rowdgv.Cells(Me.ColIncludeBundling.Index).Value, _
                    rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value, rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, _
                    rowdgv.Cells(Me.ColExplainAverageQty.Index).Value, tblHistorical, rowdgv.Cells(Me.ColDecimal.Index).Value)
            If Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                ''hitung safety stock qty
                Me.obj.CalculateSafetyStock(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
            ElseIf Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined Then
                ''hitung safety stock days
                Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
            End If
            If Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage Then
                ''hitung lead time stock
                Me.obj.CalculateLeadTimeQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
            ElseIf Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined Then
                ''hitung lead time days
                Me.obj.CalculatedLeadTimeDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
            End If

            'rnew::
            If Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                ''hitung optimum stock qty
                Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
            ElseIf Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined Then
                ''hitung optimum stock days
                Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
            End If

            ''save to DB........D :D
            Me.SaveParamsB(Me.dgvItem.SelectedCells(0).RowIndex, dtaB)

            '''tambahan batasan qty:
            'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
            '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
            '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
            '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
            '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
            '    End If
            '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
            '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
            '    End If
            '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
            '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
            '    End If
            '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
            '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
            '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaB.Dispose()
            tblHistorical.Dispose()
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub cboSafetyStockStrategy_Enter_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSafetyStockStrategy.Enter
        Try
            Me.buffer = cboSafetyStockStrategy.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboSafetyStockStrategy_Leave_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSafetyStockStrategy.Leave
        Dim rowdgv As DataGridViewRow
        Dim dtaB As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            dtaB.Connection.ConnectionString = Core_Module.GetConstring
            dtaD.Connection.ConnectionString = dtaB.Connection.ConnectionString

            If Me.cboSafetyStockStrategy.Text = Me.buffer Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.cboSafetyStockStrategy.SelectedIndex >= 0 Then
                If Me.dgvItem.SelectedCells.Count > 0 Then
                    rowdgv = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                    If Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                        rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = "Based On Usage"
                        ''hitung ulang safety stock qty
                        Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined Then
                        rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = "User Defined"
                        ''hitung safety stock days
                        Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                    ''save to DB........D :D
                    Me.SaveParamsB(Me.dgvItem.SelectedCells(0).RowIndex, dtaB)

                    '''tambahan batasan qty:
                    'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                    '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                    '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                    'End If

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            dtaB.Dispose()
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub cboLeadTimeStrategy_Enter_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLeadTimeStrategy.Enter
        Try
            Me.buffer = cboLeadTimeStrategy.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboLeadTimeStrategy_Leave_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLeadTimeStrategy.Leave
        Dim rowdgv As DataGridViewRow
        Dim dtaB As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            If Me.cboLeadTimeStrategy.Text = Me.buffer Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            dtaB.Connection.ConnectionString = Core_Module.GetConstring
            dtaD.Connection.ConnectionString = dtaB.Connection.ConnectionString

            If Me.cboLeadTimeStrategy.SelectedIndex >= 0 Then
                If Me.dgvItem.SelectedCells.Count > 0 Then
                    rowdgv = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                    If Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage Then
                        rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = "Based On Usage"
                        ''hitung ulang lead time qty
                        Me.obj.CalculateLeadTimeQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined Then
                        rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = "User Defined"
                        ''hitung ulang lead time days
                        Me.obj.CalculatedLeadTimeDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                    ''save to DB........D :D
                    Me.SaveParamsB(Me.dgvItem.SelectedCells(0).RowIndex, dtaB)

                    '''tambahan batasan qty:
                    'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                    '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                    '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                    'End If

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            dtaB.Dispose()
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub chkIncludeSales_Enter_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeSales.Enter
        Try
            Me.buffer = Me.chkIncludeSales.CheckState.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkIncludeSales_Leave_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeSales.Leave
        Dim rowdgv As DataGridViewRow
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            dtaD.Connection.ConnectionString = Core_Module.GetConstring

            If Me.chkIncludeSales.CheckState.ToString = Me.buffer Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.dgvItem.SelectedCells.Count > 0 Then
                rowdgv = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                If Me.chkIncludeSales.Checked = True Then
                    rowdgv.Cells(Me.ColIncludeSales.Index).Value = True
                Else
                    rowdgv.Cells(Me.ColIncludeSales.Index).Value = False
                End If
                Me.ReCalculateAvgSafetystockAndLeadTime(Me.dgvItem.SelectedCells(0).RowIndex)

                '''tambahan batasan qty:
                'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                '    End If
                '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                '    End If
                '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                '    End If
                '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                'End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub chkIncludeBundling_Enter_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeBundling.Enter
        Try
            Me.buffer = Me.chkIncludeBundling.CheckState.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkIncludeBundling_Leave_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeBundling.Leave
        Dim rowdgv As DataGridViewRow
        Try
            If Me.chkIncludeBundling.CheckState.ToString = Me.buffer Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.dgvItem.SelectedCells.Count > 0 Then
                rowdgv = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                If Me.chkIncludeBundling.Checked = True Then
                    rowdgv.Cells(Me.ColIncludeBundling.Index).Value = True
                Else
                    rowdgv.Cells(Me.ColIncludeBundling.Index).Value = False
                End If
                Me.ReCalculateAvgSafetystockAndLeadTime(Me.dgvItem.SelectedCells(0).RowIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub chkIncludeAdjustment_Enter_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeAdjustment.Enter
        Try
            Me.buffer = Me.chkIncludeAdjustment.CheckState.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkIncludeAdjustment_Leave_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeAdjustment.Leave
        Dim rowdgv As DataGridViewRow
        Try
            If Me.chkIncludeAdjustment.CheckState.ToString = Me.buffer Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.dgvItem.SelectedCells.Count > 0 Then
                rowdgv = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                If Me.chkIncludeAdjustment.Checked = True Then
                    rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value = True
                Else
                    rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value = False
                End If
                Me.ReCalculateAvgSafetystockAndLeadTime(Me.dgvItem.SelectedCells(0).RowIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSupplierSelectionMode.Index Or _
                    e.ColumnIndex = Me.ColPreferedSupplier.Index Or _
                    e.ColumnIndex = Me.ColRecomendedOrderQtyType.Index Or _
                    e.ColumnIndex = Me.ColMinStockQty.Index Or _
                    e.ColumnIndex = Me.ColOptStockQty.Index Or _
                    e.ColumnIndex = Me.ColMaxStockQty.Index Or _
                    e.ColumnIndex = Me.ColMinOrderQty.Index Or _
                    e.ColumnIndex = Me.ColOptOrderQty.Index Or _
                    e.ColumnIndex = Me.ColMaxOrderQty.Index Or _
                    e.ColumnIndex = Me.ColStartPeriod.Index Or _
                    e.ColumnIndex = Me.ColEndPeriod.Index Or _
                    e.ColumnIndex = Me.ColSafetyStockDays.Index Or _
                    e.ColumnIndex = Me.ColSafetyStockQty.Index Or _
                    e.ColumnIndex = Me.ColLeadTimeDays.Index Or _
                    e.ColumnIndex = Me.ColLeadTimeQty.Index Or _
                    e.ColumnIndex = Me.ColExclude.Index Or _
                    e.ColumnIndex = Me.ColProcurementStrategy.Index Or _
                    e.ColumnIndex = Me.ColProcurementType.Index Or _
                    e.ColumnIndex = Me.ColOptStockQty.Index Or _
                    e.ColumnIndex = Me.ColOptStockDays.Index Or _
                    e.ColumnIndex = Me.ColDef_LowerTolerancePerct.Index Or _
                    e.ColumnIndex = Me.ColDef_UpperTolerancePerct.Index Then
                If e.ColumnIndex = Me.ColSafetyStockQty.Index And Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                    e.Cancel = True
                End If
                If e.ColumnIndex = Me.ColSafetyStockDays.Index And Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined Then
                    e.Cancel = True
                End If
                If e.ColumnIndex = Me.ColLeadTimeQty.Index And Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                    e.Cancel = True
                End If
                If e.ColumnIndex = Me.ColLeadTimeDays.Index And Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined Then
                    e.Cancel = True
                End If

                If e.ColumnIndex = Me.ColOptStockQty.Index And Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                    e.Cancel = True
                End If
                If e.ColumnIndex = Me.ColOptStockDays.Index And Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined Then
                    e.Cancel = True
                End If

                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Me.indexRowBegin = e.RowIndex
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Try
            If e.ColumnIndex = Me.ColAverageMonthlyDemand.Index And e.RowIndex <> -1 Then
                Me.Cursor = Cursors.WaitCursor
                Dim frmExplain As New FrmTrans_MaterialPlanningDashboard_Explain
                Dim arrayBarisAvg As String()
                arrayBarisAvg = Me.dgvItem.Item(Me.ColExplainAverageQty.Index, e.RowIndex).Value.ToString.Split(vbNewLine)
                If arrayBarisAvg.Length < 9 Then
                    MsgBox("The Explanation of Average Monthly Demand is not correct, it will not be displayed.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                arrayBarisAvg = Me.dgvItem(Me.ColExplainSafetyStockQty.Index, e.RowIndex).Value.ToString.Split(vbNewLine)
                If arrayBarisAvg.Length < 2 Then
                    MsgBox("The Explanation of Safety Stock is not correct, it will not be displayed.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                arrayBarisAvg = Me.dgvItem(Me.ColExplainLeadTimeQty.Index, e.RowIndex).Value.ToString.Split(vbNewLine)
                If arrayBarisAvg.Length < 2 Then
                    MsgBox("The Explanation of Lead Stock is not correct, it will not be displayed.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                arrayBarisAvg = Me.dgvItem(Me.ColExplainOptStockQty.Index, e.RowIndex).Value.ToString.Split(vbNewLine)
                If arrayBarisAvg.Length < 2 Then
                    MsgBox("The Explanation of Optimum Stock is not correct, it will not be displayed.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                frmExplain.ExplainAvg = Me.dgvItem.Item(Me.ColExplainAverageQty.Index, e.RowIndex).Value.ToString
                frmExplain.ExplainSafetyStock = Me.dgvItem.Item(Me.ColExplainSafetyStockQty.Index, e.RowIndex).Value.ToString
                frmExplain.ExplainLeadTime = Me.dgvItem.Item(Me.ColExplainLeadTimeQty.Index, e.RowIndex).Value.ToString

                frmExplain.ExplainOptimumStock = Me.dgvItem.Item(Me.ColExplainOptStockQty.Index, e.RowIndex).Value.ToString

                If Me.dgvItem.Item(Me.ColSafetyStockStrategy.Index, e.RowIndex).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage) Then
                    frmExplain.pSafateyStock = True
                End If
                If Me.dgvItem.Item(Me.ColLeadTimeStrategy.Index, e.RowIndex).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage) Then
                    frmExplain.pLeadTime = True
                End If

                If Me.dgvItem.Item(Me.ColOptStockStrategy.Index, e.RowIndex).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage) Then
                    frmExplain.pOptimumStock = True
                End If

                frmExplain.vPrimaryProduct = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value
                frmExplain.ShowDialog()
                frmExplain.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim rowdgv As DataGridViewRow
        Dim frmSupp As New Core_Module.FrmCariMasterSupplier
        Dim selectSuppBy As Core_Module.clsMasterProduct.EnumSupplierSelectionMode
        Dim orderQtyMode As Core_Module.clsMasterProduct.EnumOrderStrategy
        Dim isExclude As Integer = -1
        Dim dtaB As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            dtaB.Connection.ConnectionString = Core_Module.GetConstring
            dtaD.Connection.ConnectionString = dtaB.Connection.ConnectionString

            rowdgv = Me.dgvItem.Rows(Me.indexRowBegin)

            ''PROCUREMENT::
            If e.ColumnIndex = Me.ColPreferedSupplier.Index Then
                If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    If Me.objSupp.ApakahSudahAda(Me.dgvItem.Item(Me.ColPreferedSupplier.Index, e.RowIndex).Value) = True Then
                        Me.dgvItem.Item(Me.ColPrimaryPreferedSupplier.Index, e.RowIndex).Value = Me.objSupp.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvItem.Item(Me.ColSupplierName.Index, e.RowIndex).Value = Me.objSupp.GetSupplierName(Me.dgvItem.Item(Me.ColPrimaryPreferedSupplier.Index, e.RowIndex).Value)
                    Else
                        frmSupp.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        frmSupp.ShowDialog()
                        If frmSupp.PrimarySupplier > 0 Then
                            Me.dgvItem.Item(Me.ColPrimaryPreferedSupplier.Index, e.RowIndex).Value = frmSupp.PrimarySupplier
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.objSupp.GetSupplierCode(frmSupp.PrimarySupplier)
                            Me.dgvItem.Item(Me.ColSupplierName.Index, e.RowIndex).Value = Me.objSupp.GetSupplierName(frmSupp.PrimarySupplier)
                        Else
                            Me.dgvItem.Item(Me.ColPrimaryPreferedSupplier.Index, e.RowIndex).Value = "-1"
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Me.dgvItem.Item(Me.ColSupplierName.Index, e.RowIndex).Value = ""
                        End If
                    End If
                Else
                    Me.dgvItem.Item(Me.ColPrimaryPreferedSupplier.Index, e.RowIndex).Value = "-1"
                    Me.dgvItem.Item(Me.ColSupplierName.Index, e.RowIndex).Value = ""
                End If
                If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                    Exit Sub
                End If
                Me.obj.UpdateParams_A_PreferedSupplier(CDbl(Me.dgvItem.Item(Me.ColPrimaryPreferedSupplier.Index, e.RowIndex).Value), Me.UserCode, CDbl(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColSupplierSelectionMode.Index Then
                If rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order) Then
                    selectSuppBy = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order
                ElseIf rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier) Then
                    selectSuppBy = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier
                ElseIf rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO) Then
                    selectSuppBy = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO
                End If
                If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                    Exit Sub
                End If
                Me.obj.UpdateParams_A_SupplierSelectionMode(selectSuppBy, Me.UserCode, CDbl(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColRecomendedOrderQtyType.Index Then
                If rowdgv.Cells(Me.ColRecomendedOrderQtyType.Index).Value = Me.cboOrderQtyMode.Items.Item(0) Then
                    orderQtyMode = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty
                Else
                    orderQtyMode = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty
                End If
                If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                    Exit Sub
                End If
                Me.obj.UpdateParams_A_RecomendOrderType(orderQtyMode, Me.UserCode, CDbl(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColDef_UpperTolerancePerct.Index Then
                If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'pastikan angka non negatif
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'nilainya max = 50%
                    If CDec(Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value) > 50 Then
                        MsgBox("Maximum tolerance allowed is 50%", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit gouping
                    Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value, 2)
                    'update ke dalam db
                    Me.obj.UpdateParams_A_DefaultUpperTolerance(Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value, _
                    Me.UserCode, CDbl(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value), dtaB)
                Else
                    Me.dgvItem.Item(Me.ColDef_UpperTolerancePerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColDef_LowerTolerancePerct.Index Then
                If String.IsNullOrEmpty(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'pastikan angka non negatif
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'nilainya max = 50%
                    If CDec(Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value) > 50 Then
                        MsgBox("Maximum tolerance allowed is 50%", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit gouping
                    Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value, 2)
                    'update ke dalam db
                    Me.obj.UpdateParams_A_DefaultLowerTolerance(Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value, _
                    Me.UserCode, CDbl(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value), dtaB)
                Else
                    Me.dgvItem.Item(Me.ColDef_LowerTolerancePerct.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Exit Sub
            End If

            ''QTY::
            If e.ColumnIndex = Me.ColMinOrderQty.Index Or _
                    e.ColumnIndex = Me.ColOptOrderQty.Index Or _
                    e.ColumnIndex = Me.ColMaxOrderQty.Index Or _
                    e.ColumnIndex = Me.ColMinStockQty.Index Or _
                    e.ColumnIndex = Me.ColMaxStockQty.Index Then
                If Me.ObjCD.CekNumerik(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    MsgBox("Please enter number", MsgBoxStyle.Information, "Information")
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If Me.ObjCD.CekNonNegatif(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If e.ColumnIndex = Me.ColMinStockQty.Index And CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    MsgBox("Minimum Stock Qty value must bigger than Safety Stock Qty value: (" & rowdgv.Cells(Me.ColSafetyStockQty.Index).Value.ToString & ")", MsgBoxStyle.Information, "Information")
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If e.ColumnIndex = Me.ColOptStockQty.Index And CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    MsgBox("Optimum Stock Qty value must bigger than Safety Stock Qty value: (" & rowdgv.Cells(Me.ColSafetyStockQty.Index).Value.ToString & ")", MsgBoxStyle.Information, "Information")
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If e.ColumnIndex = Me.ColMaxStockQty.Index And CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value < rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    MsgBox("Maximum Stock Qty value must bigger than Safety Stock Qty value: (" & rowdgv.Cells(Me.ColSafetyStockQty.Index).Value.ToString & ")", MsgBoxStyle.Information, "Information")
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowdgv.Cells(e.ColumnIndex).Value, rowdgv.Cells(Me.ColDecimal.Index).Value)
                If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                    Exit Sub
                End If
                Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                        rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                        Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                Exit Sub
            End If

            ''EXCLUSION::
            If e.ColumnIndex = Me.ColExclude.Index Then
                If rowdgv.Cells(e.ColumnIndex).Value = Boolean.FalseString Then
                    isExclude = 0
                ElseIf rowdgv.Cells(e.ColumnIndex).Value = Boolean.TrueString Then
                    isExclude = 1
                End If
                If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                    Exit Sub
                End If
                Me.obj.Update_C_Exlude(isExclude, Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value)
                Exit Sub
            End If

            ''AVG, SAFETY STOCK, LEAD TIME
            If e.ColumnIndex = Me.ColEndPeriod.Index Or e.ColumnIndex = Me.ColStartPeriod.Index Then
                If String.IsNullOrEmpty(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    If IsDate(rowdgv.Cells(e.ColumnIndex).Value) = True Then
                        Dim tblHistorical As New DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable
                        rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.GetFirstDateofTheMonth(rowdgv.Cells(e.ColumnIndex).Value)
                        rowdgv.Cells(e.ColumnIndex).Value = Format(rowdgv.Cells(e.ColumnIndex).Value, "MMMM yyyy")
                        If Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing) > Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing) Then
                            MsgBox("Start Period can not bigger than End Period", MsgBoxStyle.Information, "Information")
                            rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                            Exit Sub
                        End If

                        ''rnew::
                        'max 36 bulan
                        Dim d1 As DateTime = Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing)
                        Dim d2 As DateTime = Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing)
                        Dim M As Integer = Math.Abs((d1.Year - d2.Year))
                        Dim months As Integer = ((M * 12) + Math.Abs((d1.Month - d2.Month)))
                        If months > 36 Then
                            MsgBox("Maximum different value of period allowed is 3 years (36 months).", MsgBoxStyle.Information, "Information")
                            rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                            Exit Sub
                        End If

                        ''hitung avg dkk::
                        ''get dulu historicalnya::
                        Me.obj.GetHistoricalByPrimary(tblHistorical, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                            Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                            Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing))
                        ''hitung avg monthly
                        Me.obj.CalculateAverageMonthlyDemandByPrimary(rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                                Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColIncludeSales.Index).Value, rowdgv.Cells(Me.ColIncludeBundling.Index).Value, _
                                rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value, rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, _
                                rowdgv.Cells(Me.ColExplainAverageQty.Index).Value, tblHistorical, rowdgv.Cells(Me.ColDecimal.Index).Value)
                        ''hitung safety stock atau safety stock time
                        If rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage) Then
                            Me.obj.CalculateSafetyStock(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        ElseIf rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined) Then
                            Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        End If
                        ''hitung lead time stock atau lead time days
                        If rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage) Then
                            Me.obj.CalculateLeadTimeQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        ElseIf rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined) Then
                            Me.obj.CalculatedLeadTimeDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                        End If

                        ''rnew::
                        If rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage) Then
                            Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        ElseIf rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined) Then
                            Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                        End If

                        ''save to DB........D :D
                        Me.SaveParamsB(e.RowIndex, dtaB)

                        '''tambahan batasan qty:
                        'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                        '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                        '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                        '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                        '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                        '    End If
                        '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                        '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                        '    End If
                        '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                        '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                        '    End If
                        '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                        '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                        '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                        'End If

                        Exit Sub
                    Else
                        MsgBox("You must input in date format ('January 2000')", MsgBoxStyle.Information, "Exlamation")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                Else
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
            End If
            If e.ColumnIndex = Me.ColSafetyStockDays.Index Then
                If String.IsNullOrEmpty(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    If Me.ObjCD.CekNumerik(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowdgv.Cells(e.ColumnIndex).Value, 0)
                    ''hitung safety stock qty
                    Me.obj.CalculateSafetyStock(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, _
                            Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                            rowdgv.Cells(Me.ColDecimal.Index).Value)

                    'cek jika safety stock menjadi > optimum stock maka perubahan dicancel
                    Me.buffer1 = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    If CDec(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) > CDec(rowdgv.Cells(Me.ColOptStockQty.Index).Value) Then
                        MsgBox("Please make sure optimum stock qty >= safety stock qty.", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(Me.ColSafetyStockQty.Index).Value = Me.buffer1
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Me.obj.CalculateSafetyStock(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)

                        Exit Sub
                    End If
                    If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    ''save to DB........D :D
                    Me.SaveParamsB(e.RowIndex, dtaB)

                    '''tambahan batasan qty:
                    'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                    '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                    '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                    'End If
                    Exit Sub
                Else
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                End If
            End If
            If e.ColumnIndex = Me.ColSafetyStockQty.Index Then
                If String.IsNullOrEmpty(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    If Me.ObjCD.CekNumerik(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowdgv.Cells(e.ColumnIndex).Value, 0)
                    ''hitung safety stock days
                    Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                            Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                            rowdgv.Cells(Me.ColDecimal.Index).Value)
                    If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    'pastikan optimum stock > safety stock
                    Me.buffer1 = rowdgv.Cells(Me.ColSafetyStockDays.Index).Value
                    If CDec(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) > CDec(rowdgv.Cells(Me.ColOptStockQty.Index).Value) Then
                        MsgBox("Please make sure optimum stock qty >= safety stock qty.", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(Me.ColSafetyStockDays.Index).Value = Me.buffer1
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)

                        Exit Sub
                    End If

                    ''save to DB........D :D
                    Me.SaveParamsB(e.RowIndex, dtaB)

                    '''tambahan batasan qty:
                    'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                    '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                    '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                    'End If

                    Exit Sub
                Else
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                End If
            End If
            If e.ColumnIndex = Me.ColLeadTimeDays.Index Then
                If String.IsNullOrEmpty(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    If Me.ObjCD.CekNumerik(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowdgv.Cells(e.ColumnIndex).Value, 0)
                    ''hitung lead time stock
                    Me.obj.CalculateLeadTimeQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                            Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                            rowdgv.Cells(Me.ColDecimal.Index).Value)
                    If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    ''save to DB........D :D
                    Me.SaveParamsB(e.RowIndex, dtaB)
                    Exit Sub
                Else
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                End If
            End If
            If e.ColumnIndex = Me.ColLeadTimeQty.Index Then
                If String.IsNullOrEmpty(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    If Me.ObjCD.CekNumerik(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowdgv.Cells(e.ColumnIndex).Value, 0)
                    ''hitung lead time days
                    Me.obj.CalculatedLeadTimeDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                            Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                            rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ''save to DB........D :D
                    If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    Me.SaveParamsB(e.RowIndex, dtaB)
                    Exit Sub
                Else
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                End If
            End If
            If e.ColumnIndex = Me.ColOptStockDays.Index Then
                If String.IsNullOrEmpty(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    If Me.ObjCD.CekNumerik(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowdgv.Cells(e.ColumnIndex).Value, 0)
                    ''hitung opt stock qty
                    Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                            Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                            rowdgv.Cells(Me.ColDecimal.Index).Value)
                    If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    Me.buffer1 = rowdgv.Cells(Me.ColOptStockQty.Index).Value
                    If CDec(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) > CDec(rowdgv.Cells(Me.ColOptStockQty.Index).Value) Then
                        MsgBox("Please make sure optimum stock qty >= safety stock qty.", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(Me.ColOptStockQty.Index).Value = Me.buffer1
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                             Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                             rowdgv.Cells(Me.ColDecimal.Index).Value)

                        Exit Sub
                    End If
                    ''save to DB........D :D
                    Me.SaveParamsB(e.RowIndex, dtaB)
                    Exit Sub
                Else
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                End If
            End If

            If e.ColumnIndex = Me.ColOptStockQty.Index Then
                If String.IsNullOrEmpty(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                    If Me.ObjCD.CekNumerik(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(rowdgv.Cells(e.ColumnIndex).Value) = False Then
                        MsgBox("Please enter non-negative number", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    rowdgv.Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(rowdgv.Cells(e.ColumnIndex).Value, 0)
                    ''hitung opt stock days
                    Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                            Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                            rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ''save to DB........D :D
                    Me.buffer1 = rowdgv.Cells(Me.ColOptStockDays.Index).Value
                    If CDec(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) > CDec(rowdgv.Cells(Me.ColOptStockQty.Index).Value) Then
                        MsgBox("Please make sure optimum stock qty >= safety stock qty.", MsgBoxStyle.Information, "Information")
                        rowdgv.Cells(Me.ColOptStockDays.Index).Value = Me.buffer1
                        rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                        Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)

                        Exit Sub
                    End If

                    If rowdgv.Cells(e.ColumnIndex).Value = Me.buffer Then
                        Exit Sub
                    End If
                    Me.SaveParamsB(e.RowIndex, dtaB)
                    Exit Sub
                Else
                    rowdgv.Cells(e.ColumnIndex).Value = Me.buffer
                End If
            End If

            If e.ColumnIndex = Me.ColProcurementStrategy.Index Then
                If Me.dgvItem.Item(Me.ColProcurementStrategy.Index, e.RowIndex).Value = "Build To Order" Then
                    Me.obj.UpdateParams_ProcurementStrategy(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order)
                Else
                    Me.obj.UpdateParams_ProcurementStrategy(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock)
                End If
            End If
            If e.ColumnIndex = Me.ColProcurementType.Index Then
                If Me.dgvItem.Item(Me.ColProcurementType.Index, e.RowIndex).Value = "Buy" Then
                    Me.obj.UpdateParams_ProcurementType(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Core_Module.clsMasterProduct.EnumProcurementType.Buy)
                Else
                    Me.obj.UpdateParams_ProcurementType(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Core_Module.clsMasterProduct.EnumProcurementType.Produce)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frmSupp.Dispose()
            dtaB.Dispose()
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub SaveParamsB(ByVal vRowIndex As Integer, ByVal dtaB As DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter)
        Dim rowdgv As DataGridViewRow
        Dim pSafetyStockStrategy As Integer = 0
        Dim pLeadTimeStrategy As Integer = 0
        Dim pOptimumStrategy As Integer = 0
        Dim pIncludeSales As Integer = 0
        Dim pIncludeBundling As Integer = 0
        Dim pIncludeAdjustment As Integer = 0

        Try
            rowdgv = Me.dgvItem.Rows(vRowIndex)
            If rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage) Then
                pSafetyStockStrategy = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage
            ElseIf rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined) Then
                pSafetyStockStrategy = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined
            End If
            If rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage) Then
                pLeadTimeStrategy = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage
            ElseIf rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined) Then
                pLeadTimeStrategy = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined
            End If
            If rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage) Then
                pOptimumStrategy = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage
            ElseIf rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined) Then
                pOptimumStrategy = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined
            End If
            If rowdgv.Cells(Me.ColIncludeSales.Index).Value = False Then
                pIncludeSales = 0
            ElseIf rowdgv.Cells(Me.ColIncludeSales.Index).Value = True Then
                pIncludeSales = 1
            End If
            If rowdgv.Cells(Me.ColIncludeBundling.Index).Value = False Then
                pIncludeBundling = 0
            ElseIf rowdgv.Cells(Me.ColIncludeBundling.Index).Value = True Then
                pIncludeBundling = 1
            End If
            If rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value = False Then
                pIncludeAdjustment = 0
            ElseIf rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value = True Then
                pIncludeAdjustment = 1
            End If
            Me.obj.Update_B_All(Me.UserCode, rowdgv.Cells(Me.ColStartPeriod.Index).Value, rowdgv.Cells(Me.ColEndPeriod.Index).Value, rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, _
                rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, pSafetyStockStrategy, pLeadTimeStrategy, pIncludeSales, pIncludeBundling, pIncludeAdjustment, rowdgv.Cells(Me.ColExplainAverageQty.Index).Value, _
                rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                pOptimumStrategy, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaB)
            'MsgBox("Parameters Saved", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.dgvItem.SelectedCells.Count <= 0 Then
                Exit Sub
            End If
            If Me.dgvItem.Item(Me.ColSafetyStockStrategy.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage) Then
                Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage
            ElseIf Me.dgvItem.Item(Me.ColSafetyStockStrategy.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined) Then
                Me.cboSafetyStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined
            End If
            If Me.dgvItem.Item(Me.ColLeadTimeStrategy.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage) Then
                Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage
            ElseIf Me.dgvItem.Item(Me.ColLeadTimeStrategy.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined) Then
                Me.cboLeadTimeStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined
            End If

            If Me.dgvItem.Item(Me.ColOptStockStrategy.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage) Then
                Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage
            ElseIf Me.dgvItem.Item(Me.ColOptStockStrategy.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined) Then
                Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined
            End If

            If Me.dgvItem.Item(Me.ColIncludeSales.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = False Then
                Me.chkIncludeSales.Checked = False
            Else
                Me.chkIncludeSales.Checked = True
            End If
            If Me.dgvItem.Item(Me.ColIncludeBundling.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = False Then
                Me.chkIncludeBundling.Checked = False
            Else
                Me.chkIncludeBundling.Checked = True
            End If
            If Me.dgvItem.Item(Me.ColIncludeAdjustment.Index, Me.dgvItem.SelectedCells.Item(0).RowIndex).Value = False Then
                Me.chkIncludeAdjustment.Checked = False
            Else
                Me.chkIncludeAdjustment.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click

        Try
            Control.CheckForIllegalCrossThreadCalls = False
            If Me.dgvItem.Rows.Count <= 0 Then
                MsgBox("Load one product or more", MsgBoxStyle.Information, "Information")
                Me.txtProductCode.Focus()
                Exit Sub
            End If

            Select Case Me.cboParameter.SelectedIndex
                ''PROCUREMENT::
                Case 0
                    If Me.chkPreferedSupp.Checked = False And Me.chkSelectSuppBy.Checked = False And _
                    Me.chkOrderQtyTo.Checked = False And Me.ChkProcurementStrategy.Checked = False And _
                    Me.chkUpperTolerance.Checked = False And Me.chkLowerTolerance.Checked = False And _
                    Me.chkProcurementType.Checked = False Then
                        MsgBox("Check one parameter or more", MsgBoxStyle.Exclamation, "Exlamation")
                        Exit Sub
                    End If
                    'If Me.chkPreferedSupp.CheckState = CheckState.Checked And (Me.txtPrimaryPreferedSupp.Text = "" Or Me.txtPrimaryPreferedSupp.Text = "-1") Then
                    '    MsgBox("Choose a Prefered Supplier Code", MsgBoxStyle.Exclamation, "Exlamation")
                    '    Me.txtPreferedSuppCode.Focus()
                    '    Exit Select
                    'End If
                    If Me.chkSelectSuppBy.CheckState = CheckState.Checked And Me.cboSuppMode.SelectedIndex < 0 Then
                        MsgBox("Choose Supplier Selection Mode", MsgBoxStyle.Exclamation, "Exlamation")
                        Me.cboSuppMode.Focus()
                        Exit Sub
                    End If
                    If Me.chkOrderQtyTo.CheckState = CheckState.Checked And Me.cboOrderQtyMode.SelectedIndex < 0 Then
                        MsgBox("Choose Order Qty Mode", MsgBoxStyle.Exclamation, "Exlamation")
                        Me.cboOrderQtyMode.Focus()
                        Exit Sub
                    End If

                    ''rnew
                    If Me.ChkProcurementStrategy.CheckState = CheckState.Checked And Me.cboProcurementStrategy.SelectedIndex < 0 Then
                        MsgBox("Choose Procurement Strategy", MsgBoxStyle.Information, "Information")
                        Me.cboProcurementStrategy.Focus()
                        Exit Sub
                    End If
                    If Me.chkProcurementType.CheckState = CheckState.Checked And Me.cboProcurementType.SelectedIndex < 0 Then
                        MsgBox("Choose Procurement Type", MsgBoxStyle.Information, "Information")
                        Me.cboProcurementType.Focus()
                        Exit Sub
                    End If

                    If MsgBox("Are you sure to update all data?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Update") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.LblProgress.Text = ""
                    Me.LblProgress.Visible = True
                    Me.NonAktifkanForm()
                    Me.BackgroundWorker2.RunWorkerAsync()

                    ''AVG,SAFETY STOCK, LEAD TIME::
                Case 1
                    If Me.chkAvgDemand.Checked = False And Me.chkAllIncludeSales.Checked = False And Me.chkAllIncludeBundling.Checked = False And Me.chkAllIncludeAdjustment.Checked = False And _
                           Me.chkAllSafetyStockStrategy.Checked = False And Me.chkAllSafetyStockParams.Checked = False And _
                           Me.chkAllLeadTimeStrategy.Checked = False And Me.chkAllLeadTimeParams.Checked = False And _
                           Me.chkAllOptimumStockStrategy.Checked = False And Me.chkAllOptimumStockParams.Checked = False Then
                        MsgBox("Check one parameter or more", MsgBoxStyle.Exclamation, "Exlamation")
                        Exit Sub
                    End If
                    If Me.chkAvgDemand.Checked = True And (Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value) > Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value)) Then
                        MsgBox("Start Period value can't bigger than End Period", MsgBoxStyle.Information, "Information")
                        Me.dtpTo.Focus()
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(Me.txtAllSafetyStock.Text) = False Then
                        MsgBox("Please enter non-negative number for Safety Stock Parameter", MsgBoxStyle.Information, "Information")
                        Me.txtAllSafetyStock.Focus()
                        Exit Sub
                    End If
                    If Me.ObjCD.CekNonNegatif(Me.txtAllLeadTime.Text) = False Then
                        MsgBox("Please enter non-negative number for Lead Time Parameter", MsgBoxStyle.Information, "Information")
                        Me.txtAllLeadTime.Focus()
                        Exit Sub
                    End If

                    If Me.ObjCD.CekNonNegatif(Me.txtAllOptimumStock.Text) = False Then
                        MsgBox("Please enter non-negative number for Optimum Stock Parameter", MsgBoxStyle.Information, "Information")
                        Me.txtAllOptimumStock.Focus()
                        Exit Sub
                    End If

                    If MsgBox("Are you sure to update all data?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Update") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                    ''update ke dgv dulu
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.LblProgress.Text = ""
                    Me.LblProgress.Visible = True
                    Me.NonAktifkanForm()
                    Me.BackgroundWorker3.RunWorkerAsync()

                    ''QTY::
                Case 3
                    If Me.chkAllMinStock.Checked = False And Me.chkAllOptStock.Checked = False And Me.chkAllMaxStock.Checked = False And _
                        Me.chkAllMinOrder.Checked = False And Me.chkAllOptOrder.Checked = False And Me.chkAllMaximumOrder.Checked = False Then
                        MsgBox("Check one parameter or more", MsgBoxStyle.Exclamation, "Exlamation")
                        Exit Sub
                    End If
                    If Me.chkAllMinStock.Checked = True And Me.txtAllMinStock.Text = "" Then
                        MsgBox("Input Minimum Stock Qty", MsgBoxStyle.Information, "Information")
                        Me.txtAllMinStock.Focus()
                        Exit Sub
                    End If
                    If Me.chkAllOptStock.Checked = True And Me.txtAllOptStock.Text = "" Then
                        MsgBox("Input Optimum Stock Qty", MsgBoxStyle.Information, "Information")
                        Me.txtAllOptStock.Focus()
                        Exit Sub
                    End If
                    If Me.chkAllMaxStock.Checked = True And Me.txtAllMaxStock.Text = "" Then
                        MsgBox("Input Maximum Stock Qty", MsgBoxStyle.Information, "Information")
                        Me.txtAllMaxStock.Focus()
                        Exit Sub
                    End If
                    If Me.chkAllMinOrder.Checked = True And Me.txtAllMinOrder.Text = "" Then
                        MsgBox("Input Minimum Order Qty", MsgBoxStyle.Information, "Information")
                        Me.txtAllMinOrder.Focus()
                        Exit Sub
                    End If
                    If Me.chkAllOptOrder.Checked = True And Me.txtAllOptOrder.Text = "" Then
                        MsgBox("Input Optimum Order Qty", MsgBoxStyle.Information, "Information")
                        Me.txtAllOptOrder.Focus()
                        Exit Sub
                    End If
                    If Me.chkAllMaximumOrder.Checked = True And Me.txtAllMaxOrder.Text = "" Then
                        MsgBox("Input Maximum Order Qty", MsgBoxStyle.Information, "Information")
                        Me.txtAllMaxOrder.Focus()
                        Exit Sub
                    End If
                    If MsgBox("Are you sure to update all data?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Update") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    Me.NonAktifkanForm()
                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.LblProgress.Text = ""
                    Me.LblProgress.Visible = True
                    Me.BackgroundWorker4.RunWorkerAsync()
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub


#Region "make up"

    Private Sub chkPreferedSupp_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPreferedSupp.CheckedChanged
        Try
            If Me.chkPreferedSupp.CheckState = CheckState.Checked Then
                Me.txtPreferedSuppCode.Enabled = True
                Me.btnFindPreferedSupp.Enabled = True
            ElseIf Me.chkPreferedSupp.CheckState = CheckState.Unchecked Then
                Me.chkPreferedSupp.Checked = False
                Me.btnFindPreferedSupp.Enabled = False
                Me.txtPreferedSuppCode.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkSelectSuppBy_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSelectSuppBy.CheckedChanged
        Try
            If Me.chkSelectSuppBy.CheckState = CheckState.Checked Then
                Me.cboSuppMode.Enabled = True
            ElseIf Me.chkSelectSuppBy.CheckState = CheckState.Unchecked Then
                Me.cboSuppMode.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkOrderQtyTo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOrderQtyTo.CheckedChanged
        Try
            If Me.chkOrderQtyTo.CheckState = CheckState.Checked Then
                Me.cboOrderQtyMode.Enabled = True
            ElseIf Me.chkOrderQtyTo.CheckState = CheckState.Unchecked Then
                Me.cboOrderQtyMode.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        Try
            'dibuat supaya tidak error pada event rows added. kalo ini dihapuskan nanti error
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkFilter_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles chkFilter.LinkClicked
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.SplitContainer1.Panel1Collapsed = True Then
                Me.chkFilter.Text = "<<< Hide Filter"
                Me.SplitContainer1.Panel1Collapsed = False
            ElseIf Me.SplitContainer1.Panel1Collapsed = False Then
                Me.chkFilter.Text = ">>> Show Filter"
                Me.SplitContainer1.Panel1Collapsed = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.GotFocus
        Try
            If Me.txtProductCode.Enabled = True Then
                txtProductCode.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.LostFocus
        Try
            Me.txtProductCode.BackColor = Me.TxtPrimaryGrop.BackColor
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TxtProductDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductDesc.GotFocus
        Try
            Me.TxtProductDesc.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TxtProductDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductDesc.LostFocus
        Try
            Me.TxtProductDesc.BackColor = Me.TxtPrimaryGrop.BackColor
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtSuppCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuppCode.GotFocus
        Try
            txtSuppCode.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtSuppCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuppCode.LostFocus
        Try
            Me.txtSuppCode.BackColor = Me.TxtPrimaryGrop.BackColor
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkAvgDemand_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAvgDemand.CheckedChanged
        Try
            If Me.chkAvgDemand.Checked = True Then
                Me.dtpFrom.Enabled = True
                Me.LblTo.Enabled = True
                Me.dtpTo.Enabled = True
            Else
                Me.dtpFrom.Enabled = False
                Me.LblTo.Enabled = False
                Me.dtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllIncludeAdjustment_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeAdjustment.CheckedChanged
        Try
            If Me.chkAllIncludeAdjustment.Checked = True Then
                Me.cboAllIncludeAdjustment.Enabled = True
            Else
                Me.cboAllIncludeAdjustment.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllIncludeSales_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeSales.CheckedChanged
        Try
            If Me.chkAllIncludeSales.Checked = True Then
                Me.cboAllIncludeSales.Enabled = True
            Else
                Me.cboAllIncludeSales.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllIncludeBundling_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeBundling.CheckedChanged
        Try
            If Me.chkAllIncludeBundling.Checked = True Then
                Me.cboAllIncludeBundling.Enabled = True
            Else
                Me.cboAllIncludeBundling.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllSafetyStockStrategy_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllSafetyStockStrategy.CheckedChanged
        Try
            If chkAllSafetyStockStrategy.Checked = True Then
                cboAllSafetyStock.Enabled = True
                If Me.cboAllSafetyStock.SelectedIndex = 0 Then
                    Me.cboAllSafetyStockParams.SelectedIndex = 0
                ElseIf Me.cboAllSafetyStock.SelectedIndex = 1 Then
                    Me.cboAllSafetyStockParams.SelectedIndex = 1
                End If
                cboAllSafetyStockParams.Enabled = False
            Else
                cboAllSafetyStock.Enabled = False
                If Me.cboAllSafetyStock.SelectedIndex = 0 Then
                    Me.cboAllSafetyStockParams.SelectedIndex = 0
                ElseIf Me.cboAllSafetyStock.SelectedIndex = 1 Then
                    Me.cboAllSafetyStockParams.SelectedIndex = 1
                End If
                If Me.chkAllSafetyStockParams.Checked = True Then
                    cboAllSafetyStockParams.Enabled = True
                Else
                    cboAllSafetyStockParams.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllSafetyStockParams_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllSafetyStockParams.CheckedChanged
        Try
            If Me.chkAllSafetyStockParams.Checked = True Then
                If Me.chkAllSafetyStockStrategy.Checked = False Then
                    Me.cboAllSafetyStockParams.Enabled = True
                Else
                    Me.cboAllSafetyStockParams.Enabled = False
                End If
                Me.txtAllSafetyStock.Enabled = True
                'Me.lblAllSafetyStockUnit.Enabled = True
            Else
                Me.cboAllSafetyStockParams.Enabled = False
                Me.txtAllSafetyStock.Enabled = False
                'Me.lblAllSafetyStockUnit.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllLeadTimeParams_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllLeadTimeParams.CheckedChanged
        Try
            If Me.chkAllLeadTimeParams.Checked = True Then
                If Me.chkAllLeadTimeStrategy.Checked = False Then
                    Me.cboAllLeadTimeParams.Enabled = True
                Else
                    Me.cboAllLeadTimeParams.Enabled = False
                End If
                Me.txtAllLeadTime.Enabled = True
                'Me.lblAllSafetyStockUnit.Enabled = True
            Else
                Me.cboAllLeadTimeParams.Enabled = False
                Me.txtAllLeadTime.Enabled = False
                'Me.lblAllSafetyStockUnit.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllLeadTimeStrategy_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllLeadTimeStrategy.CheckedChanged
        Try
            If chkAllLeadTimeStrategy.Checked = True Then
                cboAllLeadTime.Enabled = True
                If Me.cboAllLeadTime.SelectedIndex = 0 Then
                    Me.cboAllLeadTimeParams.SelectedIndex = 0
                ElseIf Me.cboAllLeadTime.SelectedIndex = 1 Then
                    Me.cboAllLeadTimeParams.SelectedIndex = 1
                End If
                cboAllLeadTimeParams.Enabled = False
            Else
                cboAllLeadTime.Enabled = False
                If Me.cboAllLeadTime.SelectedIndex = 0 Then
                    Me.cboAllLeadTimeParams.SelectedIndex = 0
                ElseIf Me.cboAllLeadTime.SelectedIndex = 1 Then
                    Me.cboAllLeadTimeParams.SelectedIndex = 1
                End If
                If Me.chkAllLeadTimeParams.Checked = True Then
                    cboAllLeadTimeParams.Enabled = True
                Else
                    cboAllLeadTimeParams.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub cboAllSafetyStockParams_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAllSafetyStockParams.SelectedIndexChanged
        Try
            If Me.cboAllSafetyStockParams.Text = "Safety Time" Then
                Me.lblAllSafetyStockUnit.Text = "Days"
            Else
                Me.lblAllSafetyStockUnit.Text = "Unit"
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboAllLeadTimeParams_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAllLeadTimeParams.SelectedIndexChanged
        Try
            If Me.cboAllLeadTimeParams.Text = "Lead Time" Then
                Me.lblAllLeadTimeUnit.Text = "Days"
            Else
                Me.lblAllLeadTimeUnit.Text = "Unit"
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkAllMinOrder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllMinOrder.CheckedChanged
        Try
            If Me.chkAllMinOrder.Checked = True Then
                Me.txtAllMinOrder.Enabled = True
            ElseIf Me.chkAllMinOrder.Checked = False Then
                Me.txtAllMinOrder.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllOptOrder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllOptOrder.CheckedChanged
        Try
            If Me.chkAllOptOrder.Checked = True Then
                Me.txtAllOptOrder.Enabled = True
            ElseIf Me.chkAllOptOrder.Checked = False Then
                Me.txtAllOptOrder.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllMaximumOrder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllMaximumOrder.CheckedChanged
        Try
            If Me.chkAllMaximumOrder.Checked = True Then
                Me.txtAllMaxOrder.Enabled = True
            ElseIf Me.chkAllMaximumOrder.Checked = False Then
                Me.txtAllMaxOrder.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllMinStock_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllMinStock.CheckedChanged
        Try
            If Me.chkAllMinStock.Checked = True Then
                Me.txtAllMinStock.Enabled = True
            ElseIf Me.chkAllMinStock.Checked = False Then
                Me.txtAllMinStock.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllOptStock_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllOptStock.CheckedChanged
        Try
            If Me.chkAllOptStock.Checked = True Then
                Me.txtAllOptStock.Enabled = True
            ElseIf Me.chkAllOptStock.Checked = False Then
                Me.txtAllOptStock.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllMaxStock_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllMaxStock.CheckedChanged
        Try
            If Me.chkAllMaxStock.Checked = True Then
                Me.txtAllMaxStock.Enabled = True
            ElseIf Me.chkAllMaxStock.Checked = False Then
                Me.txtAllMaxStock.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMinOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMinOrder.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllMinOrder, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMinOrder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMinOrder.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllMinOrder, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllOptOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllOptOrder.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllOptOrder, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllOptOrder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllOptOrder.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllOptOrder, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMaxOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMaxOrder.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllMaxOrder, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMaxOrder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMaxOrder.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllMaxOrder, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMinStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMinStock.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllMinStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMinStock_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMinStock.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllMinStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllOptStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllOptStock.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllOptStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllOptStock_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllOptStock.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllOptStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMaxStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMaxStock.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllMaxStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtAllMaxStock_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllMaxStock.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllMaxStock, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGroupCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.GotFocus
        Try
            Me.txtGroupCode.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtGroupCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGroupCode.LostFocus
        Try
            Me.txtGroupCode.BackColor = Me.TxtPrimaryGrop.BackColor
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtPreferedSuppCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPreferedSuppCode.GotFocus
        Try
            Me.txtPreferedSuppCode.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtPreferedSuppCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPreferedSuppCode.LostFocus
        Try
            If Me.txtPreferedSuppCode.Enabled = True Then
                Me.txtPreferedSuppCode.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtPreferedSuppCode.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllSafetyStock_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllSafetyStock.GotFocus
        Try
            Me.txtAllSafetyStock.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllSafetyStock_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllSafetyStock.LostFocus
        Try
            If Me.txtAllSafetyStock.Enabled = True Then
                Me.txtAllSafetyStock.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllSafetyStock.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllLeadTime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllLeadTime.GotFocus
        Try
            Me.txtAllLeadTime.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllLeadTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllLeadTime.LostFocus
        Try
            If Me.txtAllLeadTime.Enabled = True Then
                Me.txtAllLeadTime.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllLeadTime.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtWorkDays_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWorkDays.GotFocus
        Try
            Me.txtWorkDays.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtWorkDays_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWorkDays.LostFocus
        Try
            Me.txtWorkDays.BackColor = Me.TxtPrimaryGrop.BackColor
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMinStock_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMinStock.GotFocus
        Try
            Me.txtAllMinStock.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMinStock_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMinStock.LostFocus
        Try
            If Me.txtAllMinStock.Enabled = True Then
                Me.txtAllMinStock.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllMinStock.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllOptStock_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllOptStock.GotFocus
        Try
            Me.txtAllOptStock.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllOptStock_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllOptStock.LostFocus
        Try
            If Me.txtAllOptStock.Enabled = True Then
                Me.txtAllOptStock.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllOptStock.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMaxStock_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMaxStock.LostFocus
        Try
            If Me.txtAllMaxStock.Enabled = True Then
                Me.txtAllMaxStock.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllMaxStock.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMaxStock_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMaxStock.GotFocus
        Try
            Me.txtAllMaxStock.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMinOrder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMinOrder.GotFocus
        Try
            Me.txtAllMinOrder.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMinOrder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMinOrder.LostFocus
        Try
            If Me.txtAllMinOrder.Enabled = True Then
                Me.txtAllMinOrder.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllMinOrder.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllOptOrder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllOptOrder.GotFocus
        Try
            Me.txtAllOptOrder.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllOptOrder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllOptOrder.LostFocus
        Try
            If Me.txtAllOptOrder.Enabled = True Then
                Me.txtAllOptOrder.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllOptOrder.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMaxOrder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMaxOrder.GotFocus
        Try
            Me.txtAllMaxOrder.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllMaxOrder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllMaxOrder.LostFocus
        Try
            If Me.txtAllMaxOrder.Enabled = True Then
                Me.txtAllMaxOrder.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllMaxOrder.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllSafetyStock_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllSafetyStock.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllSafetyStock, sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllSafetyStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllSafetyStock.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllSafetyStock, sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllLeadTime_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllLeadTime.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtAllLeadTime, sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllLeadTime_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAllLeadTime.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtAllLeadTime, sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboAllSafetyStock_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAllSafetyStock.SelectedIndexChanged
        Try
            If Me.cboAllSafetyStock.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                Me.cboAllSafetyStockParams.SelectedIndex = 0
            Else
                Me.cboAllSafetyStockParams.SelectedIndex = 1
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboAllLeadTime_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAllLeadTime.SelectedIndexChanged
        Try
            If Me.cboAllLeadTime.SelectedIndex = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage Then
                Me.cboAllLeadTimeParams.SelectedIndex = 0
            Else
                Me.cboAllLeadTimeParams.SelectedIndex = 1
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkIncludeSales_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.buffer = Me.chkIncludeSales.CheckState.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkIncludeBundling_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.buffer = Me.chkIncludeBundling.CheckState.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkIncludeAdjustment_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.buffer = Me.chkIncludeAdjustment.CheckState.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboSafetyStockStrategy_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.buffer = cboSafetyStockStrategy.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboLeadTimeStrategy_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.buffer = cboLeadTimeStrategy.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Data_ResizeEnd(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.ResizeEnd
        Try
            Me.SplitContainer2.Panel1Collapsed = False
            Me.SplitContainer2.SplitterDistance = 130
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtWorkDays_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWorkDays.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtWorkDays, sender, e)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtWorkDays_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWorkDays.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtWorkDays, sender, e)
        Catch ex As Exception

        End Try
    End Sub

#End Region

    Private Sub txtWorkDays_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWorkDays.Enter
        Try
            Me.buffer = Me.txtWorkDays.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtWorkDays_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWorkDays.Leave
        Try
            If CDbl(Me.txtWorkDays.Text) <> CDbl(Me.buffer) Then
                If Me.CheckForRecalculateAll = True Then
                    If MsgBox("Update work days per months?" & vbNewLine & "System will recalculate all product", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Me.txtWorkDays.Text = Me.buffer
                        Exit Sub
                    End If

                    Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan, Me.txtWorkDays.Text)
                    Me.CheckForRecalculateAll = False

                    If Me.BackgroundWorker5.IsBusy = False Then
                        Me.NonAktifkanForm()
                        Control.CheckForIllegalCrossThreadCalls = False
                        Me.BackgroundWorker5.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim selectSuppBy As Core_Module.clsMasterProduct.EnumSupplierSelectionMode
        Dim orderQtyMode As Core_Module.clsMasterProduct.EnumOrderStrategy
        Dim procurementStrategy As Core_Module.clsMasterProduct.EnumProcurementStrategy
        Dim ProcurementType As Core_Module.clsMasterProduct.EnumProcurementType
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            Me.ObjCD.UpdateThreading()
            Me.Cursor = Cursors.WaitCursor
            Me.Prg1.Maximum = Me.dgvItem.Rows.Count
            Me.Prg1.Visible = True

            dta.Connection.ConnectionString = Core_Module.GetConstring
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                Me.Prg1.Value = Me.dgvItem.Rows.IndexOf(rowdgv) + 1
                Me.LblProgress.Text = "Updating data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " & rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString + " - " + rowdgv.Cells(Me.ColDescription.Index).Value
                If Me.chkPreferedSupp.Checked = True Then
                    rowdgv.Cells(Me.ColPrimaryPreferedSupplier.Index).Value = Me.txtPrimaryPreferedSupp.Text
                    rowdgv.Cells(Me.ColPreferedSupplier.Index).Value = Me.objSupp.GetSupplierCode(CDbl(Me.txtPrimaryPreferedSupp.Text))
                End If
                If Me.chkSelectSuppBy.Checked = True Then
                    If Me.cboSuppMode.SelectedIndex = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order Then
                        rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order)
                    ElseIf Me.cboSuppMode.SelectedIndex = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier Then
                        rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier)
                    ElseIf Me.cboSuppMode.SelectedIndex = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO Then
                        rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO)
                    End If
                End If
                If Me.chkOrderQtyTo.Checked = True Then
                    If Me.cboOrderQtyMode.SelectedIndex = 0 Then
                        rowdgv.Cells(Me.ColRecomendedOrderQtyType.Index).Value = Me.cboOrderQtyMode.Items.Item(0)
                    Else
                        rowdgv.Cells(Me.ColRecomendedOrderQtyType.Index).Value = Me.cboOrderQtyMode.Items.Item(1)
                    End If
                End If

                ''rnew
                If Me.ChkProcurementStrategy.Checked = True Then
                    If Me.cboProcurementStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order Then
                        rowdgv.Cells(Me.ColProcurementStrategy.Index).Value = Me.cboProcurementStrategy.Items.Item(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order)
                    ElseIf Me.cboProcurementStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                        rowdgv.Cells(Me.ColProcurementStrategy.Index).Value = Me.cboProcurementStrategy.Items.Item(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock)
                    End If
                End If
                If Me.chkProcurementType.Checked = True Then
                    If Me.cboProcurementType.SelectedIndex = Core_Module.clsMasterProduct.EnumProcurementType.Buy Then
                        rowdgv.Cells(Me.ColProcurementType.Index).Value = Me.cboProcurementType.Items.Item(Core_Module.clsMasterProduct.EnumProcurementType.Buy)
                    ElseIf Me.cboProcurementType.SelectedIndex = Core_Module.clsMasterProduct.EnumProcurementType.Produce Then
                        rowdgv.Cells(Me.ColProcurementType.Index).Value = Me.cboProcurementType.Items.Item(Core_Module.clsMasterProduct.EnumProcurementType.Produce)
                    End If
                End If

                If rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order) Then
                    selectSuppBy = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order
                ElseIf rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier) Then
                    selectSuppBy = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier
                ElseIf rowdgv.Cells(Me.ColSupplierSelectionMode.Index).Value = Me.cboSuppMode.Items.Item(Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO) Then
                    selectSuppBy = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO
                End If
                If rowdgv.Cells(Me.ColRecomendedOrderQtyType.Index).Value = Me.cboOrderQtyMode.Items.Item(0) Then
                    orderQtyMode = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty
                Else
                    orderQtyMode = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty
                End If

                ''rnew
                If rowdgv.Cells(Me.ColProcurementStrategy.Index).Value = Me.cboProcurementStrategy.Items.Item(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order) Then
                    procurementStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order
                ElseIf rowdgv.Cells(Me.ColProcurementStrategy.Index).Value = Me.cboProcurementStrategy.Items.Item(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock) Then
                    procurementStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock
                End If
                If rowdgv.Cells(Me.ColProcurementType.Index).Value = Me.cboProcurementType.Items.Item(Core_Module.clsMasterProduct.EnumProcurementType.Buy) Then
                    ProcurementType = Core_Module.clsMasterProduct.EnumProcurementType.Buy
                ElseIf rowdgv.Cells(Me.ColProcurementType.Index).Value = Me.cboProcurementType.Items.Item(Core_Module.clsMasterProduct.EnumProcurementType.Produce) Then
                    ProcurementType = Core_Module.clsMasterProduct.EnumProcurementType.Produce
                End If

                If Me.chkUpperTolerance.Checked = True Then
                    rowdgv.Cells(Me.ColDef_UpperTolerancePerct.Index).Value = Me.TxtUpperTolerance.Text
                End If
                If Me.chkLowerTolerance.Checked = True Then
                    rowdgv.Cells(Me.ColDef_LowerTolerancePerct.Index).Value = Me.TxtLowerTolerance.Text
                End If

                Me.obj.UpdateParams_A(CDbl(rowdgv.Cells(Me.ColPrimaryPreferedSupplier.Index).Value), selectSuppBy, _
                        orderQtyMode, Me.UserCode, procurementStrategy, CDbl(rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value), dta, _
                        CDec(rowdgv.Cells(Me.ColDef_UpperTolerancePerct.Index).Value), _
                        CDec(rowdgv.Cells(Me.ColDef_LowerTolerancePerct.Index).Value), ProcurementType)
            Next
            If Me.dgvItem.Rows.Count > 0 Then
                Dim rowDGV As DataGridViewRow
                Me.dgvItem.ClearSelection()
                rowDGV = Me.dgvItem.Rows(0)
                rowDGV.Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Private Sub AktifkanForm()
        Try
            Me.ControlBox = True
            Me.GroupBox4.Enabled = True
            Me.GroupBox2.Enabled = True
            Me.GroupBox1.Enabled = True
            Me.GroupBox3.Enabled = True
            Me.dgvItem.Enabled = True
            Me.GroupBox5.Enabled = True
            chkFilter.Enabled = True
            btnApply.Enabled = True
            btnClose.Enabled = True
            cboParameter.Enabled = True
            Me.Cursor = Cursors.Arrow

            Me.btnApply.Enabled = True
            Me.txtWorkDays.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanForm()
        Try
            Me.ControlBox = False
            Me.GroupBox4.Enabled = False
            Me.GroupBox2.Enabled = False
            Me.GroupBox1.Enabled = False
            Me.GroupBox3.Enabled = False
            Me.dgvItem.Enabled = False
            Me.GroupBox5.Enabled = False
            chkFilter.Enabled = False
            btnApply.Enabled = False
            btnClose.Enabled = False
            cboParameter.Enabled = False
            Me.Cursor = Cursors.Arrow

            Me.btnApply.Enabled = False
            Me.txtWorkDays.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.AktifkanForm()
            Me.LblProgress.Visible = False
            Me.Prg1.Visible = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            MsgBox("Procurement Parameters Updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Dim LSales As New List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable)
        Dim LBundling As New List(Of DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable)
        Dim LAdj As New List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable)
        Dim tblIndexSales As New DseTrans_MaterialPlanningDashboard_Data.indexSalesDataTable
        Dim tblIndexBundling As New DseTrans_MaterialPlanningDashboard_Data.indexBundlingDataTable
        Dim tblIndexAdjustment As New DseTrans_MaterialPlanningDashboard_Data.indexAdjustmentDataTable
        Dim tblHistorical As New DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable
        Dim startPeriodMin As Date = Me.ObjCD.GetFirstDateofTheMonth(Now)
        Dim finishPeriodMax As Date = Me.ObjCD.GetLasDateOfTheMonth(Now)
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            Me.ObjCD.UpdateThreading()
            Me.Cursor = Cursors.WaitCursor
            dtaD.Connection.ConnectionString = Core_Module.GetConstring

            Me.Prg1.Visible = True
            Me.Prg1.Maximum = Me.dgvItem.Rows.Count
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                Me.Prg1.Value = Me.dgvItem.Rows.IndexOf(rowdgv) + 1
                Me.LblProgress.Text = "Updating datagridview - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " & rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString + " - " + rowdgv.Cells(Me.ColDescription.Index).Value
                If chkAvgDemand.Checked = True Then
                    Me.dtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
                    Me.dtpTo.Value = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value)
                    rowdgv.Cells(Me.ColStartPeriod.Index).Value = Format(Me.dtpFrom.Value, "MMMM yyyy")
                    rowdgv.Cells(Me.ColEndPeriod.Index).Value = Format(Me.dtpTo.Value, "MMMM yyyy")
                End If
                If Me.chkAllIncludeSales.Checked = True Then
                    If Me.cboAllIncludeSales.Text = "Yes" Then
                        rowdgv.Cells(Me.ColIncludeSales.Index).Value = True
                    Else
                        rowdgv.Cells(Me.ColIncludeSales.Index).Value = False
                    End If
                End If
                If Me.chkAllIncludeBundling.Checked = True Then
                    If Me.cboAllIncludeBundling.Text = "Yes" Then
                        rowdgv.Cells(Me.ColIncludeBundling.Index).Value = True
                    Else
                        rowdgv.Cells(Me.ColIncludeBundling.Index).Value = False
                    End If
                End If
                If Me.chkAllIncludeAdjustment.Checked = True Then
                    If Me.cboAllIncludeAdjustment.Text = "Yes" Then
                        rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value = True
                    Else
                        rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value = False
                    End If
                End If
                If Me.chkAllSafetyStockStrategy.Checked = True Then
                    If Me.cboAllSafetyStock.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                        rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = "Based On Usage"
                    ElseIf Me.cboAllSafetyStock.SelectedIndex = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined Then
                        rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = "User Defined"
                    End If
                End If
                If Me.chkAllLeadTimeStrategy.Checked = True Then
                    If Me.cboAllLeadTime.SelectedIndex = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage Then
                        rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = "Based On Usage"
                    ElseIf Me.cboAllLeadTime.SelectedIndex = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined Then
                        rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = "User Defined"
                    End If
                End If

                ''rnew::
                If Me.chkAllOptimumStockStrategy.Checked = True Then
                    If Me.cboAllOptimumStock.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                        rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = "Based On Usage"
                    ElseIf Me.cboAllOptimumStock.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined Then
                        rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = "User Defined"
                    End If
                End If

                If Me.chkAllSafetyStockParams.Checked = True Then
                    If Me.cboAllSafetyStockParams.Text = "Safety Time" Then
                        rowdgv.Cells(Me.ColSafetyStockDays.Index).Value = Me.txtAllSafetyStock.Text
                    Else
                        rowdgv.Cells(Me.ColSafetyStockQty.Index).Value = Me.ObjCD.DigitGrouping(Me.txtAllSafetyStock.Text, rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                End If
                If Me.chkAllLeadTimeParams.Checked = True Then
                    If Me.cboAllLeadTimeParams.Text = "Lead Time" Then
                        rowdgv.Cells(Me.ColLeadTimeDays.Index).Value = Me.txtAllLeadTime.Text
                    Else
                        rowdgv.Cells(Me.ColLeadTimeQty.Index).Value = Me.ObjCD.DigitGrouping(Me.txtAllLeadTime.Text, rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                End If

                ''rnew::
                If Me.chkAllOptimumStockParams.Checked = True Then
                    If Me.cboAllOptimumStockParams.Text = "Optimum Time" Then
                        rowdgv.Cells(Me.ColOptStockDays.Index).Value = Me.txtAllOptimumStock.Text
                    Else
                        rowdgv.Cells(Me.ColOptStockQty.Index).Value = Me.ObjCD.DigitGrouping(Me.txtAllOptimumStock.Text, rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                End If

            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.chkAvgDemand.Checked = True Then
                ''##FILL ALL DAN HITUNG SEMUA KOMPONEN##
                ''get historical::
                Me.obj.GetHistoricalByPrimary(tblHistorical, -1, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value))
                ''load data sales, bundling, adj::
                Me.obj.GetListSales(LSales, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value), tblIndexSales)
                Me.obj.GetListBundling(LBundling, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value), tblIndexBundling)
                Me.obj.GetListAdj(LAdj, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value), tblIndexAdjustment)
                ''hitung avg, dkk
                For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                    Me.Prg1.Value = Me.dgvItem.Rows.IndexOf(rowdgv) + 1
                    Me.LblProgress.Text = "Recalculating data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " & rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString + " - " + rowdgv.Cells(Me.ColDescription.Index).Value

                    Me.obj.CalculateAverageMonthlyDemandAll(rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                        Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColIncludeSales.Index).Value, rowdgv.Cells(Me.ColIncludeBundling.Index).Value, _
                        rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value, rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, _
                        rowdgv.Cells(Me.ColExplainAverageQty.Index).Value, tblHistorical, LSales, LBundling, LAdj, tblIndexSales, tblIndexBundling, tblIndexAdjustment, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
                    If rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage) Then
                        Me.obj.CalculateSafetyStock(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined) Then
                        Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                    If rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage) Then
                        Me.obj.CalculateLeadTimeQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined) Then
                        Me.obj.CalculatedLeadTimeDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If

                    ''rnew::
                    If rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage) Then
                        Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined) Then
                        Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                Next
            ElseIf Me.chkAvgDemand.Checked = False And _
                    (Me.chkAllIncludeSales.Checked = True Or Me.chkAllIncludeBundling.Checked = True Or Me.chkAllIncludeAdjustment.Checked = True) Then
                ''##FILL ALL DAN HITUNG SEMUA KOMPONEN##
                ''##TAPI SUMBER PERIODE DARI DGV MAX AND MIN##
                ''get historical::
                For Each rowdgv0 As DataGridViewRow In Me.dgvItem.Rows
                    If Date.ParseExact("01 " & rowdgv0.Cells(Me.ColStartPeriod.Index).Value.ToString, "dd MMMM yyyy", Nothing) < startPeriodMin Then
                        startPeriodMin = Me.ObjCD.GetFirstDateofTheMonth(Date.ParseExact("01 " & rowdgv0.Cells(Me.ColStartPeriod.Index).Value.ToString, "dd MMMM yyyy", Nothing))
                    End If
                    If Me.ObjCD.GetLasDateOfTheMonth(Date.ParseExact("01 " & rowdgv0.Cells(Me.ColEndPeriod.Index).Value.ToString, "dd MMMM yyyy", Nothing)) > finishPeriodMax Then
                        finishPeriodMax = Me.ObjCD.GetLasDateOfTheMonth(Date.ParseExact("01 " & rowdgv0.Cells(Me.ColEndPeriod.Index).Value.ToString, "dd MMMM yyyy", Nothing))
                    End If
                Next
                Me.obj.GetHistoricalByPrimary(tblHistorical, -1, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax))
                ''load data sales, bundling, adj::
                Me.obj.GetListSales(LSales, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexSales)
                Me.obj.GetListBundling(LBundling, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexBundling)
                Me.obj.GetListAdj(LAdj, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexAdjustment)
                ''hitung avg, dkk
                For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                    Me.Prg1.Value = Me.dgvItem.Rows.IndexOf(rowdgv) + 1
                    Me.LblProgress.Text = "Recalculating data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " & rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString + " - " + rowdgv.Cells(Me.ColDescription.Index).Value
                    Me.obj.CalculateAverageMonthlyDemandAll(rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        Date.ParseExact("01 " & rowdgv.Cells(Me.ColStartPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                        Date.ParseExact("01 " & rowdgv.Cells(Me.ColEndPeriod.Index).Value, "dd MMMM yyyy", Nothing), _
                        Me.txtWorkDays.Text, rowdgv.Cells(Me.ColIncludeSales.Index).Value, rowdgv.Cells(Me.ColIncludeBundling.Index).Value, _
                        rowdgv.Cells(Me.ColIncludeAdjustment.Index).Value, rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, _
                        rowdgv.Cells(Me.ColExplainAverageQty.Index).Value, tblHistorical, LSales, LBundling, LAdj, tblIndexSales, tblIndexBundling, tblIndexAdjustment, _
                        rowdgv.Cells(Me.ColDecimal.Index).Value)
                    If rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage) Then
                        Me.obj.CalculateSafetyStock(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined) Then
                        Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                    If rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage) Then
                        Me.obj.CalculateLeadTimeQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined) Then
                        Me.obj.CalculatedLeadTimeDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If

                    ''rnew::
                    If rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage) Then
                        Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined) Then
                        Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If

                Next
            ElseIf (Me.chkAvgDemand.Checked = False And Me.chkAllIncludeSales.Checked = False And Me.chkAllIncludeBundling.Checked = False And Me.chkAllIncludeAdjustment.Checked = False) And _
                    (Me.chkAllSafetyStockStrategy.Checked = True Or Me.chkAllSafetyStockParams.Checked = True Or Me.chkAllLeadTimeStrategy.Checked = True Or Me.chkAllLeadTimeParams.Checked = True Or _
                            Me.chkAllOptimumStockStrategy.Checked = True Or Me.chkAllOptimumStockParams.Checked = True) Then
                For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                    ''##FILL 1 by 1 DAN HANYA HITUNG KOMPONEN YANG DIPERLUKAN##
                    Me.Prg1.Value = Me.dgvItem.Rows.IndexOf(rowdgv) + 1
                    Me.LblProgress.Text = "Recalculating data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " & rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString + " - " + rowdgv.Cells(Me.ColDescription.Index).Value
                    If Me.chkAllSafetyStockStrategy.Checked = True Or Me.chkAllSafetyStockParams.Checked = True Then
                        ''hitung semua safety stock
                        If rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage) Then
                            Me.obj.CalculateSafetyStock(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        ElseIf rowdgv.Cells(Me.ColSafetyStockStrategy.Index).Value = Me.cboSafetyStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined) Then
                            Me.obj.CalculateSafetyStockTime(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColSafetyStockQty.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColSafetyStockDays.Index).Value, rowdgv.Cells(Me.ColExplainSafetyStockQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        End If
                    End If
                    If Me.chkAllLeadTimeStrategy.Checked = True Or Me.chkAllLeadTimeParams.Checked = True Then
                        ''hitung lead time
                        If rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage) Then
                            Me.obj.CalculateLeadTimeQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        ElseIf rowdgv.Cells(Me.ColLeadTimeStrategy.Index).Value = Me.cboLeadTimeStrategy.Items.Item(Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined) Then
                            Me.obj.CalculatedLeadTimeDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColLeadTimeQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColLeadTimeDays.Index).Value, rowdgv.Cells(Me.ColExplainLeadTimeQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                        End If
                    End If

                    ''rnew::
                    If Me.chkAllOptimumStockStrategy.Checked = True Or Me.chkAllOptimumStockParams.Checked = True Then
                        ''hitung optimum stock
                        If rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage) Then
                            Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                                    Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                    rowdgv.Cells(Me.ColDecimal.Index).Value)
                        ElseIf rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = Me.cboOptimumStockStrategy.Items.Item(Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined) Then
                            Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                        End If
                    End If

                Next
            End If
            ''save to DB........D :D
            Dim dtaB As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
            dtaB.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                Me.Prg1.Value = Me.dgvItem.Rows.IndexOf(rowdgv) + 1
                Me.LblProgress.Text = "Saving to database - data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " & rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString + " - " + rowdgv.Cells(Me.ColDescription.Index).Value

                '''tambahan batasan qty:
                'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                '    End If
                '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                '    End If
                '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                '    End If
                '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                'End If

                Me.SaveParamsB(Me.dgvItem.Rows.IndexOf(rowdgv), dtaB)
            Next
            If Me.dgvItem.Rows.Count > 0 Then
                Dim rowDGV As DataGridViewRow
                Me.dgvItem.ClearSelection()
                rowDGV = Me.dgvItem.Rows(0)
                rowDGV.Selected = True
            End If
            dtaB.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            LSales = Nothing
            LBundling = Nothing
            LAdj = Nothing
            tblIndexSales.Dispose()
            tblIndexBundling.Dispose()
            tblIndexAdjustment.Dispose()
            tblHistorical.Dispose()
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.AktifkanForm()
            Me.LblProgress.Visible = False
            Me.Prg1.Visible = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            MsgBox("Average, Safety Stock, and Lead Time Parameters Updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Dim minOrder As Integer = 0
        Dim optOrder As Integer = 0
        Dim maxOrder As Integer = 0
        Dim minStock As Integer = 0
        Dim optStock As Integer = 0
        Dim maxStock As Integer = 0
        Dim isSomeQtyNotUpdated As Boolean = False
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            Me.ObjCD.UpdateThreading()
            Me.Prg1.Maximum = Me.dgvItem.Rows.Count
            Me.Prg1.Visible = True
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.Cursor = Cursors.WaitCursor
            dtaD.Connection.ConnectionString = Core_Module.GetConstring

            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                Me.Prg1.Value = Me.dgvItem.Rows.IndexOf(rowdgv) + 1
                Me.LblProgress.Text = "Updating data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " & rowdgv.Cells(Me.ColCodeNo.Index).Value.ToString + " - " + rowdgv.Cells(Me.ColDescription.Index).Value
                If Me.chkAllMinOrder.Checked = True Then
                    If Me.txtAllMinOrder.Text = "" Then
                        minOrder = 0
                    Else
                        minOrder = Me.txtAllMinOrder.Text
                    End If
                    rowdgv.Cells(Me.ColMinOrderQty.Index).Value = Me.ObjCD.DigitGrouping(minOrder, CInt(rowdgv.Cells(Me.ColDecimal.Index).Value))
                End If
                If Me.chkAllOptOrder.Checked = True Then
                    If Me.txtAllOptOrder.Text = "" Then
                        optOrder = 0
                    Else
                        optOrder = Me.txtAllOptOrder.Text
                    End If
                    rowdgv.Cells(Me.ColOptOrderQty.Index).Value = Me.ObjCD.DigitGrouping(optOrder, CInt(rowdgv.Cells(Me.ColDecimal.Index).Value))
                End If
                If Me.chkAllMaximumOrder.Checked = True Then
                    If Me.txtAllMaxOrder.Text = "" Then
                        maxOrder = 0
                    Else
                        maxOrder = Me.txtAllMaxOrder.Text
                    End If
                    rowdgv.Cells(Me.ColMaxOrderQty.Index).Value = Me.ObjCD.DigitGrouping(maxOrder, CInt(rowdgv.Cells(Me.ColDecimal.Index).Value))
                End If
                If Me.chkAllMinStock.Checked = True Then
                    If Me.txtAllMinStock.Text = "" Then
                        minStock = 0
                    Else
                        minStock = Me.txtAllMinStock.Text
                    End If
                    If minStock > rowdgv.Cells(Me.ColSafetyStockQty.Index).Value Then
                        rowdgv.Cells(Me.ColMinStockQty.Index).Value = Me.ObjCD.DigitGrouping(minStock, CInt(rowdgv.Cells(Me.ColDecimal.Index).Value))
                    Else
                        isSomeQtyNotUpdated = True
                    End If
                End If
                If Me.chkAllOptStock.Checked = True Then
                    If Me.txtAllOptStock.Text = "" Then
                        optStock = 0
                    Else
                        optStock = Me.txtAllOptStock.Text
                    End If
                    If optStock > rowdgv.Cells(Me.ColSafetyStockQty.Index).Value Then
                        rowdgv.Cells(Me.ColOptStockQty.Index).Value = Me.ObjCD.DigitGrouping(optStock, CInt(rowdgv.Cells(Me.ColDecimal.Index).Value))
                    Else
                        isSomeQtyNotUpdated = True
                    End If
                End If
                If Me.chkAllMaxStock.Checked = True Then
                    If Me.txtAllMaxStock.Text = "" Then
                        maxStock = 0
                    Else
                        maxStock = Me.txtAllMaxStock.Text
                    End If
                    If maxStock > rowdgv.Cells(Me.ColSafetyStockQty.Index).Value Then
                        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = Me.ObjCD.DigitGrouping(maxStock, CInt(rowdgv.Cells(Me.ColDecimal.Index).Value))
                    Else
                        isSomeQtyNotUpdated = True
                    End If
                End If
                Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, _
                        rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                        rowdgv.Cells(Me.ColMaxStockQty.Index).Value, Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
            Next
            If Me.dgvItem.Rows.Count > 0 Then
                Dim rowDGV As DataGridViewRow
                Me.dgvItem.ClearSelection()
                rowDGV = Me.dgvItem.Rows(0)
                rowDGV.Selected = True
            End If

            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.Prg1.Visible = False
            Me.LblProgress.Visible = False
            Me.AktifkanForm()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            MsgBox("Quantity parameters updated. All qty under safety stock is not updated.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Dim dtaMasterProd As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim tblMasterProd As New DseTrans_MaterialPlanningDashboard_Data.tblmaster_product1DataTable
        Dim rowMasterProd As DseTrans_MaterialPlanningDashboard_Data.tblmaster_product1Row
        'Dim dtaReplicaProd As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product11TableAdapter
        'Dim tblReplicaProd As New DseTrans_MaterialPlanningDashboard_Data.tblmaster_product11DataTable
        'Dim rowReplicaProd As DseTrans_MaterialPlanningDashboard_Data.tblmaster_product11Row
        Dim minStock, optStock, maxStock As Double
        Dim startPeriodMin, finishPeriodMax As Date
        Dim includeSales, includeBundling, includeAdj As Boolean
        Dim safetyStockTime, leadTime, optimumStockTime As Integer
        Dim safetyStockQty, leadStockQty, optimumStockQty As Double
        Dim avgMonthly As Double
        Dim explainAvgMonthly, explainSafetyStock, explainLeadTime, explainOptimumStock As String
        Dim LSales As New List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable)
        Dim LBundling As New List(Of DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable)
        Dim LAdj As New List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable)
        Dim tblIndexSales As New DseTrans_MaterialPlanningDashboard_Data.indexSalesDataTable
        Dim tblIndexBundling As New DseTrans_MaterialPlanningDashboard_Data.indexBundlingDataTable
        Dim tblIndexAdjustment As New DseTrans_MaterialPlanningDashboard_Data.indexAdjustmentDataTable
        Dim tblHistorical As New DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable

        Try
            Me.ObjCD.UpdateThreading()
            Me.Cursor = Cursors.WaitCursor

            explainAvgMonthly = ""
            explainSafetyStock = ""
            explainLeadTime = ""
            explainOptimumStock = ""

            'Me.Prg1.Visible = True
            'Me.LblProgress.Visible = True

            dtaMasterProd.Connection.ConnectionString = Core_Module.GetConstring
            'dtaReplicaProd.Connection.ConnectionString = dtaMasterProd.Connection.ConnectionString
            If dtaMasterProd.Fill(tblMasterProd) > 0 Then
                ''fill historical, sales, bundling, dan adj sekaligus u/ semua produk
                startPeriodMin = Me.ObjCD.GetFirstDateofTheMonth(dtaMasterProd.ScalarQuery_GetMinPeriod().Value)
                finishPeriodMax = Me.ObjCD.GetLasDateOfTheMonth(dtaMasterProd.ScalarQuery_GetMaxPeriod().Value)
                Me.obj.GetHistoricalByPrimary(tblHistorical, -1, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax))
                Me.obj.GetListSales(LSales, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexSales)
                Me.obj.GetListBundling(LBundling, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexBundling)
                Me.obj.GetListAdj(LAdj, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexAdjustment)
                Me.Prg1.Visible = True
                Me.Prg1.Maximum = tblMasterProd.Rows.Count + 1
                Me.LblProgress.Visible = True

                For Each rowMasterProd In tblMasterProd
                    Me.Prg1.Value = tblMasterProd.Rows.IndexOf(rowMasterProd) + 1
                    Me.LblProgress.Text = "Updating data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " _
                    & rowMasterProd.Product_Code + " - " + rowMasterProd.Description

                    With rowMasterProd
                        minStock = .Minimum_Stock
                        optStock = .optimum_stock
                        maxStock = .maximum_stock
                        safetyStockTime = .safety_stock_days
                        leadTime = .lead_time_days
                        optimumStockTime = .optimum_stock_days

                        If .include_sales = 0 Then
                            includeSales = False
                        Else
                            includeSales = True
                        End If
                        If .include_bundling = 0 Then
                            includeBundling = False
                        Else
                            includeBundling = True
                        End If
                        If .include_adjustment = 0 Then
                            includeAdj = False
                        Else
                            includeAdj = True
                        End If
                        ''hitung avg, safety stock, dan lead stock serta explanation nya
                        Me.obj.CalculateAverageMonthlyDemandAll(.Primary, Format(.start_period, "dd MMMM yyyy"), Format(.end_period, "dd MMMM yyyy"), Me.txtWorkDays.Text, includeSales, includeBundling, includeAdj, _
                                avgMonthly, explainAvgMonthly, tblHistorical, LSales, LBundling, LAdj, tblIndexSales, tblIndexBundling, tblIndexAdjustment, ._Decimal)

                        If .safety_stock_strategy = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                            Me.obj.CalculateSafetyStock(avgMonthly, .safety_stock_days, Me.txtWorkDays.Text, safetyStockQty, explainSafetyStock, ._Decimal)
                        ElseIf .safety_stock_strategy = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined Then
                            Me.obj.CalculateSafetyStockTime(avgMonthly, .safety_stock_qty, Me.txtWorkDays.Text, safetyStockTime, explainSafetyStock, ._Decimal)
                        End If
                        If .lead_time_strategy = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage Then
                            Me.obj.CalculateLeadTimeQty(avgMonthly, .lead_time_days, Me.txtWorkDays.Text, leadStockQty, explainLeadTime, ._Decimal)
                        ElseIf .lead_time_strategy = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined Then
                            Me.obj.CalculatedLeadTimeDays(avgMonthly, .lead_time_qty, Me.txtWorkDays.Text, leadTime, explainLeadTime, ._Decimal)
                        End If

                        If .optimum_stock_strategy = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                            Me.obj.CalculateOptimumStockQty(avgMonthly, .optimum_stock_days, Me.txtWorkDays.Text, optimumStockQty, explainOptimumStock, ._Decimal)
                        ElseIf .optimum_stock_strategy = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined Then
                            Me.obj.CalculatedOptimumStockDays(avgMonthly, .optimum_stock, Me.txtWorkDays.Text, optimumStockTime, explainOptimumStock, ._Decimal)
                        End If

                        ''betulkan qty2 yang kurang dari safety stock
                        If minStock < safetyStockQty Then
                            minStock = safetyStockQty
                        End If
                        If optStock < safetyStockQty Then
                            optStock = safetyStockQty
                        End If
                        If maxStock < safetyStockQty Then
                            maxStock = safetyStockQty
                        End If

                        ''simpan perubahan ke db
                        dtaMasterProd.UpdateQuery_UpdateTrigeredByWorkDays(minStock, optStock, maxStock, avgMonthly, safetyStockTime, safetyStockQty, _
                        leadTime, leadStockQty, explainAvgMonthly, explainSafetyStock, explainLeadTime, Me.UserCode, Now, Now, optimumStockTime, explainOptimumStock, .Primary)
                    End With
                Next
                If Me.dgvItem.Rows.Count > 0 Then
                    Dim rowDGV As DataGridViewRow
                    Me.dgvItem.ClearSelection()
                    rowDGV = Me.dgvItem.Rows(0)
                    rowDGV.Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Prg1.Visible = False
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            Me.Prg1.Visible = False
            Me.LblProgress.Visible = False

            Me.AktifkanForm()
            Me.btnFilter.PerformClick()
            Me.Cursor = Cursors.Arrow
            MsgBox("Recalculation completed.", MsgBoxStyle.Information, "Information")
            Me.CheckForRecalculateAll = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkProcurementStrategy_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkProcurementStrategy.CheckedChanged
        Try
            If Me.ChkProcurementStrategy.CheckState = CheckState.Checked Then
                Me.cboProcurementStrategy.Enabled = True
            ElseIf Me.ChkProcurementStrategy.CheckState = CheckState.Unchecked Then
                Me.cboProcurementStrategy.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub chkAllOptimumStockStrategy_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllOptimumStockStrategy.CheckedChanged
        Try
            If chkAllOptimumStockStrategy.Checked = True Then
                cboAllOptimumStock.Enabled = True
                If Me.cboAllOptimumStock.SelectedIndex = 0 Then
                    Me.cboAllOptimumStockParams.SelectedIndex = 0
                ElseIf Me.cboAllOptimumStock.SelectedIndex = 1 Then
                    Me.cboAllOptimumStockParams.SelectedIndex = 1
                End If
                cboAllOptimumStockParams.Enabled = False
            Else
                cboAllOptimumStock.Enabled = False
                If Me.cboAllOptimumStock.SelectedIndex = 0 Then
                    Me.cboAllOptimumStockParams.SelectedIndex = 0
                ElseIf Me.cboAllOptimumStock.SelectedIndex = 1 Then
                    Me.cboAllOptimumStockParams.SelectedIndex = 1
                End If
                If Me.chkAllOptimumStockParams.Checked = True Then
                    cboAllOptimumStockParams.Enabled = True
                Else
                    cboAllOptimumStockParams.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub cboAllOptimumStock_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAllOptimumStock.SelectedIndexChanged
        Try
            If Me.cboAllOptimumStock.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                Me.cboAllOptimumStockParams.SelectedIndex = 0
            Else
                Me.cboAllOptimumStockParams.SelectedIndex = 1
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkAllOptimumStockParams_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllOptimumStockParams.CheckedChanged
        Try
            If Me.chkAllOptimumStockParams.Checked = True Then
                If Me.chkAllOptimumStockStrategy.Checked = False Then
                    Me.cboAllOptimumStockParams.Enabled = True
                Else
                    Me.cboAllOptimumStockParams.Enabled = False
                End If
                Me.txtAllOptimumStock.Enabled = True
                'Me.lblAllOptimumStockUnit.Enabled = True
            Else
                Me.cboAllOptimumStockParams.Enabled = False
                Me.txtAllOptimumStock.Enabled = False
                'Me.lblAllOptimumStockUnit.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub cboAllOptimumStockParams_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAllOptimumStockParams.SelectedIndexChanged
        Try
            If Me.cboAllOptimumStockParams.Text = "Optimum Time" Then
                Me.lblAllOptimumStockUnit.Text = "Days"
            Else
                Me.lblAllOptimumStockUnit.Text = "Unit"
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllOptimumStock_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllOptimumStock.GotFocus
        Try
            Me.txtAllOptimumStock.BackColor = Me.BackColorAll
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtAllOptimumStock_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAllOptimumStock.LostFocus
        Try
            If Me.txtAllOptimumStock.Enabled = True Then
                Me.txtAllOptimumStock.BackColor = Me.TxtPrimaryGrop.BackColor
            Else
                Me.txtAllOptimumStock.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboOptimumStockStrategy_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboOptimumStockStrategy.Enter
        Try
            Me.buffer = cboOptimumStockStrategy.Text
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboOptimumStockStrategy_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboOptimumStockStrategy.Leave
        Dim rowdgv As DataGridViewRow
        Dim dtaB As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim dtaD As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter

        Try
            If Me.cboOptimumStockStrategy.Text = Me.buffer Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            dtaB.Connection.ConnectionString = Core_Module.GetConstring
            dtaD.Connection.ConnectionString = dtaB.Connection.ConnectionString

            If Me.cboOptimumStockStrategy.SelectedIndex >= 0 Then
                If Me.dgvItem.SelectedCells.Count > 0 Then
                    rowdgv = Me.dgvItem.Rows(Me.dgvItem.SelectedCells(0).RowIndex)
                    If Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                        rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = "Based On Usage"
                        ''hitung ulang opt stock qty
                        Me.obj.CalculateOptimumStockQty(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockDays.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    ElseIf Me.cboOptimumStockStrategy.SelectedIndex = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined Then
                        rowdgv.Cells(Me.ColOptStockStrategy.Index).Value = "User Defined"
                        ''hitung ulang opt stock days
                        Me.obj.CalculatedOptimumStockDays(rowdgv.Cells(Me.ColAverageMonthlyDemand.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, _
                                Me.txtWorkDays.Text, rowdgv.Cells(Me.ColOptStockDays.Index).Value, rowdgv.Cells(Me.ColExplainOptStockQty.Index).Value, _
                                rowdgv.Cells(Me.ColDecimal.Index).Value)
                    End If
                    ''save to DB........D :D
                    Me.SaveParamsB(Me.dgvItem.SelectedCells(0).RowIndex, dtaB)

                    '''tambahan batasan qty:
                    'If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Or _
                    '    CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '    If CDbl(rowdgv.Cells(Me.ColMinStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMinStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColOptStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColOptStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    If CDbl(rowdgv.Cells(Me.ColMaxStockQty.Index).Value) < CDbl(rowdgv.Cells(Me.ColSafetyStockQty.Index).Value) Then
                    '        rowdgv.Cells(Me.ColMaxStockQty.Index).Value = rowdgv.Cells(Me.ColSafetyStockQty.Index).Value
                    '    End If
                    '    Me.obj.Update_D_MinOptMaxOrderStock(rowdgv.Cells(Me.ColMinOrderQty.Index).Value, rowdgv.Cells(Me.ColOptOrderQty.Index).Value, rowdgv.Cells(Me.ColMaxOrderQty.Index).Value, _
                    '            rowdgv.Cells(Me.ColMinStockQty.Index).Value, rowdgv.Cells(Me.ColOptStockQty.Index).Value, rowdgv.Cells(Me.ColMaxStockQty.Index).Value, _
                    '            Me.UserCode, rowdgv.Cells(Me.ColPrimaryCodeNo.Index).Value, dtaD)
                    'End If

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            dtaB.Dispose()
            dtaD.Dispose()
        End Try
    End Sub

    Private Sub txtProductCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.btnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtProductDesc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtProductDesc.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.btnFilter.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpTo.ValueChanged
        Try
            ''rnew::
            'max 36 bulan
            Dim d1 As DateTime = Me.dtpTo.Value
            Dim d2 As DateTime = Me.dtpFrom.Value
            Dim M As Integer = Math.Abs((d1.Year - d2.Year))
            Dim months As Integer = ((M * 12) + Math.Abs((d1.Month - d2.Month)))
            If months > 36 Then
                MsgBox("Maximum different value of period allowed is 3 years (36 months).", MsgBoxStyle.Information, "Information")
                Me.dtpTo.Value = Me.dtpFrom.Value
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.ValueChanged
        Try
            ''rnew::
            'max 36 bulan
            Dim d1 As DateTime = Me.dtpTo.Value
            Dim d2 As DateTime = Me.dtpFrom.Value
            Dim M As Integer = Math.Abs((d1.Year - d2.Year))
            Dim months As Integer = ((M * 12) + Math.Abs((d1.Month - d2.Month)))
            If months > 36 Then
                MsgBox("Maximum different value of period allowed is 3 years (36 months).", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Value = Me.dtpTo.Value
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMaterialPlanningGroup.LinkClicked
        Try
            Dim ij As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_Data1
            ij.UserCode = Me.UserCode
            ' Set the caption to be unique.
            ij.BackColorAll = Core_Module.GetSMESSBackColorAll
            ij.ShowDialog()
            Me.btnFilter.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUpperTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUpperTolerance.Enter
        Try
            Me.buffer = Me.TxtUpperTolerance.Text
            Me.ObjCD.Number_Enter(Me.TxtUpperTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUpperTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUpperTolerance.GotFocus
        Try
            If Me.TxtUpperTolerance.Enabled = True Then
                If Me.TxtUpperTolerance.ReadOnly = True Then
                    Me.TxtUpperTolerance.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtUpperTolerance.BackColor = Me.BackColorAll
                End If
            Else
                Me.TxtUpperTolerance.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUpperTolerance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUpperTolerance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUpperTolerance, sender, e, 50)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUpperTolerance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUpperTolerance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUpperTolerance, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLowerTolerance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLowerTolerance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtLowerTolerance, sender, e, 50)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLowerTolerance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLowerTolerance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtLowerTolerance, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUpperTolerance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUpperTolerance.Leave
        Try
            'cek data
            '1. pastikan angka non negatif
            If Me.ObjCD.CekNonNegatif(Me.TxtUpperTolerance.Text) = False Then
                Me.TxtUpperTolerance.Text = Me.buffer
                Exit Sub
            End If
            '2. pastikan tidak lebih dr 50%
            If CDec(Me.TxtUpperTolerance.Text) > 50 Then
                MsgBox("Maximum allowed tolerance value is 50%.", MsgBoxStyle.Information, "Information")
                Me.TxtUpperTolerance.Text = Me.buffer
                Exit Sub
            End If
            '3. digit grouping
            Me.TxtUpperTolerance.Text = Me.ObjCD.DigitGrouping(Me.TxtUpperTolerance.Text, 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUpperTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUpperTolerance.LostFocus
        Try
            If Me.TxtUpperTolerance.Enabled = True Then
                If Me.TxtUpperTolerance.ReadOnly = True Then
                    Me.TxtUpperTolerance.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtUpperTolerance.BackColor = Me.TxtPrimaryGrop.BackColor
                End If
            Else
                Me.TxtUpperTolerance.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLowerTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLowerTolerance.Enter
        Try
            Me.buffer = Me.TxtLowerTolerance.Text
            Me.ObjCD.Number_Enter(Me.TxtLowerTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLowerTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLowerTolerance.GotFocus
        Try
            If Me.TxtLowerTolerance.Enabled = True Then
                If Me.TxtLowerTolerance.ReadOnly = True Then
                    Me.TxtLowerTolerance.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtLowerTolerance.BackColor = Me.BackColorAll
                End If
            Else
                Me.TxtLowerTolerance.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLowerTolerance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLowerTolerance.Leave
        Try
            'cek data
            '1. pastikan angka non negatif
            If Me.ObjCD.CekNonNegatif(Me.TxtLowerTolerance.Text) = False Then
                Me.TxtLowerTolerance.Text = Me.buffer
                Exit Sub
            End If
            '2. pastikan tidak lebih dr 50%
            If CDec(Me.TxtLowerTolerance.Text) > 50 Then
                MsgBox("Maximum allowed tolerance value is 50%.", MsgBoxStyle.Information, "Information")
                Me.TxtLowerTolerance.Text = Me.buffer
                Exit Sub
            End If
            '3. digit grouping
            Me.TxtLowerTolerance.Text = Me.ObjCD.DigitGrouping(Me.TxtLowerTolerance.Text, 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLowerTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLowerTolerance.LostFocus
        Try
            If Me.TxtLowerTolerance.Enabled = True Then
                If Me.TxtLowerTolerance.ReadOnly = True Then
                    Me.TxtLowerTolerance.BackColor = Me.txtReadOnly.BackColor
                Else
                    Me.TxtLowerTolerance.BackColor = Me.TxtPrimaryGrop.BackColor
                End If
            Else
                Me.TxtLowerTolerance.BackColor = Me.txtReadOnly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUpperTolerance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUpperTolerance.CheckedChanged
        Try
            If Me.chkUpperTolerance.Checked = True Then
                Me.TxtUpperTolerance.ReadOnly = False
                Me.TxtUpperTolerance.Enabled = True
            Else
                Me.TxtUpperTolerance.ReadOnly = True
                Me.TxtUpperTolerance.Text = Me.ObjCD.DigitGrouping(0, 2)
                Me.TxtUpperTolerance.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkLowerTolerance_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLowerTolerance.CheckedChanged
        Try
            If Me.chkLowerTolerance.Checked = True Then
                Me.TxtLowerTolerance.ReadOnly = False
                Me.TxtLowerTolerance.Enabled = True
            Else
                Me.TxtLowerTolerance.ReadOnly = True
                Me.TxtLowerTolerance.Text = Me.ObjCD.DigitGrouping(0, 2)
                Me.TxtLowerTolerance.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAdvanceExclusion_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAdvanceExclusion.LinkClicked
        Dim ijkl As New FrmTrans_MaterialPlanningDashboard_Adv_Exclusion

        Try
            ijkl.UserCode = Me.UserCode
            ijkl.TipeTransaksiExclusion = ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion.MaterialPlanning
            ijkl.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijkl = Nothing
        End Try
    End Sub

    Private Sub chkProcurementType_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkProcurementType.CheckedChanged
        Try
            If Me.chkProcurementType.CheckState = CheckState.Checked Then
                Me.cboProcurementType.Enabled = True
            ElseIf Me.chkProcurementType.CheckState = CheckState.Unchecked Then
                Me.cboProcurementType.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class