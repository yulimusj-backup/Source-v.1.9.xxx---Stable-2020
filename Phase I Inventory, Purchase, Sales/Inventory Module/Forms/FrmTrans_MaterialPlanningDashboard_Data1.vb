Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTrans_MaterialPlanningDashboard_Data1
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim obj As New ClsTransaksiMaterialPlanningGroup
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim objItem As New ClsTransaksiMaterialPlanningGroupItem
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1
    Dim objPlanning As New ClsTransaksiMaterialPlanning_Data
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser


    Dim buffer As String = ""
    Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
    Dim StrictEdit As Boolean = True
    Dim StrictDelete As Boolean = True
    Dim DummyIndex As Integer = -1
    Dim BufferFrom As Date = Today
    Dim BufferTo As Date = Today
    Dim BufferBoolean_IncudeSales As Boolean = False
    Dim BufferBoolean_IncudeAdjustment As Boolean = False
    Dim BufferBoolean_IncudeBundling As Boolean = False
    Dim BufferBoolean_IncudeStopBeforeRefValue As Boolean = False
    Dim BufferProductCodeFilter As String = ""
    Dim BufferProductDescFilter As String = ""

    Dim TBLItem_Invoice As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable
    Dim rowItem_Invoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceRow
    Dim TBLItem_SR As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable
    Dim rowItem_SR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srRow
    Dim TBLItem_Bundling As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable
    Dim rowItem_Bundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingRow
    Dim TBLItem_Adjustment As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable
    Dim rowItem_Adjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentRow

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TabControl1.SelectedIndex = 1 And Me.TabControl2.SelectedIndex = 0 Then
                    Me.ObjCD.SetBackColor(Me.TabPage_ProductGroup.Controls, Me.txtPrimary.BackColor, False)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Dim MinRow As Integer = 0

        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TabControl1.SelectedIndex = 1 And Me.TabControl2.SelectedIndex = 0 Then
                    Me.ObjCD.SetBackColor(Me.TabPage_ProductGroup.Controls, Me.txtPrimary.BackColor, False)
                ElseIf Me.TabControl1.SelectedIndex = 0 Then
                    Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, False)
                End If
            End If
            If Me.TabControl1.SelectedTab.Name = Me.TabPage_Group.Name Then
                Me.llbCopyGroup.Visible = True
            Else
                Me.llbCopyGroup.Visible = False
            End If
            If Me.TabControl1.SelectedTab.Name = Me.TabPage_Product.Name Then
                Me.llbFindProduct.Visible = True
                Me.chkShowDetail.Visible = True
            Else
                Me.llbFindProduct.Visible = False
                Me.chkShowDetail.Visible = False
            End If
            If Me.TabControl1.SelectedTab.Name = Me.TabPage_Scope.Name Then
                If Me.TxtRemarks.Enabled = True Then
                    MinRow = 1
                Else
                    MinRow = 0
                End If
                If Me.dgvItemSupplier.Rows.Count > MinRow Then
                    If Me.txtProductCodeFilter.Text = "" And Me.TxtProductDescFilter.Text = "" And Me.dgvItemProductGroup.Rows.Count = MinRow Then
                        Me.TabControl2.SelectedTab = Me.TabPage_PreferedSupplier
                    End If
                End If
                If Me.dgvItemProductGroup.Rows.Count > MinRow Or Me.txtProductCodeFilter.Text <> "" Or Me.TxtProductDescFilter.Text <> "" Then
                    If Me.dgvItemSupplier.Rows.Count = MinRow Then
                        Me.TabControl2.SelectedTab = Me.TabPage_ProductGroup
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage_ProductGroup.Controls, Me.txtPrimary.BackColor, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.TabPage_ProductGroup.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.btnExit.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItemGroup.AllowUserToDeleteRows = True
            Me.dgvItemGroup.ReadOnly = False
            Me.dgvItemGroup.AllowUserToAddRows = True
            Me.dgvItemProductGroup.AllowUserToDeleteRows = True
            Me.dgvItemProductGroup.ReadOnly = False
            Me.dgvItemProductGroup.AllowUserToAddRows = True
            Me.dgvItemSupplier.AllowUserToDeleteRows = True
            Me.dgvItemSupplier.ReadOnly = False
            Me.dgvItemSupplier.AllowUserToAddRows = True
            Me.dgvItemProduct.ReadOnly = False
            Me.dgvItemInvoiceException.AllowUserToDeleteRows = True
            Me.dgvItemInvoiceException.ReadOnly = False
            Me.dgvItemInvoiceException.AllowUserToAddRows = True

            Me.llbRecalculate.Enabled = True
            Me.llbCopyGroup.Enabled = True
            Me.chkShowDetail.Enabled = True

            Me.chkAllIncludeAdjustment.Enabled = True
            Me.chkAllIncludeBundling.Enabled = True
            Me.chkAllIncludeSales.Enabled = True
            Me.chkStopBeforeRefValue.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.TabPage_ProductGroup.Controls, True, False, True, False, True, False, True, False)

            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)

            Me.CancelButton = Me.btnExit
            Me.btnCancel.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvItemGroup.AllowUserToDeleteRows = False
            Me.dgvItemGroup.ReadOnly = True
            Me.dgvItemGroup.AllowUserToAddRows = False
            Me.dgvItemSupplier.AllowUserToDeleteRows = False
            Me.dgvItemSupplier.ReadOnly = True
            Me.dgvItemSupplier.AllowUserToAddRows = False
            Me.dgvItemProductGroup.AllowUserToDeleteRows = False
            Me.dgvItemProductGroup.ReadOnly = True
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            Me.dgvItemProduct.AllowUserToDeleteRows = False
            Me.dgvItemProduct.ReadOnly = True
            Me.dgvItemInvoiceException.AllowUserToDeleteRows = False
            Me.dgvItemInvoiceException.ReadOnly = True
            Me.dgvItemInvoiceException.AllowUserToAddRows = False

            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage_ProductGroup.Controls, Me.BackColor, False)

            Me.llbRecalculate.Enabled = False
            Me.llbCopyGroup.Enabled = False
            Me.chkShowDetail.Enabled = True

            Me.chkAllIncludeAdjustment.Enabled = False
            Me.chkAllIncludeBundling.Enabled = False
            Me.chkAllIncludeSales.Enabled = False
            Me.chkStopBeforeRefValue.Enabled = False

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_Data1_New") = False Then
                Me.btnNew.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_Data1_Edit") = False Then
                Me.btnEdit.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_Data1_Delete") = False Then
                Me.btnDelete.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_Data1_Print") = False Then
                Me.btnPrint.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_Data1_Apply") = False Then
                Me.btnApply.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_Data1_ShowDetail") = False Then
                Me.chkShowDetail.Enabled = False
                Me.chkShowDetail.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage_ProductGroup.Controls, True, True, True)

            Me.txtPrimary.Text = "-1"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dtpDate.Value = Today
            Me.dtpTo.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpTo.Value)
            Me.DtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value)

            Me.dgvItemGroup.Rows.Clear()
            Me.dgvItemProduct.Rows.Clear()
            Me.dgvItemProductGroup.Rows.Clear()
            Me.dgvItemSupplier.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini mengasumsikan dgvGroup, dgvSupplier, dan dgvProductGroup -> allow user to add row = false
    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim Jml As Double
        Dim vDocNoAcuan As String = ""
        Dim vPrimaryAcuan As Double = -1
        Dim TipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe
        Dim rowDGV, rowDgvBefore As DataGridViewRow
        Dim TotalPerc As Decimal = 0
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_prefered_supplier

        Try
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            Me.dgvItemSupplier.AllowUserToAddRows = False
            Me.dgvItemGroup.AllowUserToAddRows = False
            '1.a. cek data yang harus diisi
            Me.txtDocumentNo.Text = Me.txtDocumentNo.Text.Trim
            Me.TxtRemarks.Text = Me.TxtRemarks.Text.Trim
            Me.txtDescription.Text = Me.txtDescription.Text.Trim
            KoleksiNamaData.Add("document number")
            KoleksiTextBox.Add(Me.txtDocumentNo)
            KoleksiNamaData.Add("description")
            KoleksiTextBox.Add(Me.txtDescription)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '1.b. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If Me.baru = True Then
                Jml = Me.obj.GetJml(Me.txtDocumentNo.Text)
                If Jml > 0 Then
                    Me.txtDocumentNo.Focus()
                    MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                vDocNoAcuan = Me.obj.GetDocumentNo(Me.obj.Primary)
                If vDocNoAcuan <> Me.txtDocumentNo.Text Then
                    vPrimaryAcuan = Me.obj.GetPrimary(Me.txtDocumentNo.Text)
                    If vPrimaryAcuan <> Me.obj.Primary And vPrimaryAcuan <> -1 Then
                        MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                        Me.txtDocumentNo.Focus()
                        Return False
                    End If
                End If
            End If

            '2. jika salah satu dari dgv berikut kosong: group, dgv product
            If Me.dgvItemGroup.Rows.Count = 0 Then
                MsgBox("Please specify at least 1 group.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 0
                Me.dgvItemGroup.Focus()
                Return False
            End If
            If Me.dgvItemProduct.Rows.Count = 0 Then
                MsgBox("Please specify at least 1 product.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 2
                Me.dgvItemProduct.Focus()
                Return False
            End If

            '2.b. jika product scope, tidak ada sama sekali
            If Me.dgvItemProductGroup.Rows.Count = 0 And Me.dgvItemSupplier.Rows.Count = 0 _
            And Me.txtProductCodeFilter.Text = "" And Me.TxtProductDescFilter.Text = "" Then
                Me.TabControl1.SelectedTab = Me.TabPage_Scope
                MsgBox("Please specify calculation scope", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. period from harus lebih kecil dari period to
            If Me.DtpFrom.Value > Me.dtpTo.Value Then
                MsgBox("Start period must be smaller than end period.", MsgBoxStyle.Information, "Information")
                Me.DtpFrom.Focus()
                Return False
            End If

            '5. jika include sales, include adjustment, dan include bundling semuanya bernilai 0
            If Me.chkAllIncludeSales.Checked = False And Me.chkAllIncludeBundling.Checked = False And Me.chkAllIncludeAdjustment.Checked = False Then
                MsgBox("Please make sure at least 1 of these data is checked: 'Include Sales', 'Include Bundling', or 'Include Adjustment'.", _
                MsgBoxStyle.Information, "Information")
                Me.chkAllIncludeAdjustment.Focus()
                Return False
            End If

            '6. untuk tipe transaksi cumulative percetage, total cumulative percetage harus 100 %
            TipeTrans = Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text)
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Then
                For Each rowDGV In dgvItemGroup.Rows
                    TotalPerc += CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value)
                Next
                If TotalPerc <> 100 Then
                    MsgBox("Please make sure the sum of ref. value data is 100%.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.dgvItemGroup.Focus()
                    Me.dgvItemGroup.ClearSelection()
                    Return False
                End If
            End If


            '7. untuk tipe transaksi percetage, nilai tidak boleh melebih 100%, baris yang atas harus selalu lebih besar nilainya dari baris yang dibawahnya
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                For Each rowDGV In dgvItemGroup.Rows
                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) > 100 Then
                        MsgBox("Please make sure maximum ref. value data is always < 100%.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvItemGroup.Focus()
                        Me.dgvItemGroup.ClearSelection()
                        rowDGV.Selected = True
                        Return False
                    End If
                    If rowDGV.Index > 0 Then
                        rowDgvBefore = dgvItemGroup.Rows(rowDGV.Index - 1)
                        If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) >= CDec(rowDgvBefore.Cells(Me.ColGroup_RefValue.Index).Value) Then
                            MsgBox("Please make sure ref. value data is smaller than the previous row.", MsgBoxStyle.Information, "Information")
                            Me.TabControl1.SelectedIndex = 0
                            Me.dgvItemGroup.Focus()
                            Me.dgvItemGroup.ClearSelection()
                            rowDGV.Selected = True
                            Return False
                        End If
                    End If
                Next
            End If

            '8. untuk setiap item ref value pada dgv group tidak boleh nilai 0
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Then
                'kecuali baris terakhir
                For Each rowDGV In dgvItemGroup.Rows
                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) = 0 And rowDGV.Index <> Me.dgvItemGroup.RowCount - 1 Then
                        MsgBox("Please make sure ref. value data is always > 0", MsgBoxStyle.Information, "Information")

                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvItemGroup.Focus()
                        Me.dgvItemGroup.ClearSelection()
                        rowDGV.Selected = True

                        Return False
                    End If
                Next
            Else
                For Each rowDGV In dgvItemGroup.Rows
                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) = 0 Then
                        MsgBox("Please make sure ref. value data is always > 0", MsgBoxStyle.Information, "Information")

                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvItemGroup.Focus()
                        Me.dgvItemGroup.ClearSelection()
                        rowDGV.Selected = True

                        Return False
                    End If
                Next
            End If

            '9. jika tipe transaksi melibatkan gross margin, maka harus sudah closing mulai dari start period - end period
            'If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            'TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            'TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Then
            If ObjClosing.GetPrimaryClosing(Me.dtpTo.Value) = -1 Then
                MsgBox("Please make sure there is inventory closing on " + Format(Me.dtpTo.Value, "MMM yyyy"), MsgBoxStyle.Information, "Information")
                Return False
            End If
            'End If

            '10. untuk setiap item group, jika strategynya user defined dan qty =0
            '11. untuk setiap item group, jika strategynya based on usage dan days =0
            'For Each rowDGV In dgvItemGroup.Rows
            '    'safety stock
            '    TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_SafetyStockStrategy.Index).Value)
            '    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
            '        If rowDGV.Cells(Me.ColGroup_SafetyStockDays.Index).Value = 0 Then
            '            MsgBox("Please specify a positive number for safety stock days in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '            MsgBoxStyle.Information, "Information")
            '            Me.TabControl1.SelectedIndex = 0
            '            Me.dgvItemGroup.ClearSelection()
            '            rowDGV.Selected = True
            '            Return False
            '        End If
            '    ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
            '        If rowDGV.Cells(Me.ColGroup_SafetyStockQty.Index).Value = 0 Then
            '            MsgBox("Please specify a positive number for safety stock qty in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '            MsgBoxStyle.Information, "Information")
            '            Me.TabControl1.SelectedIndex = 0
            '            Me.dgvItemGroup.ClearSelection()
            '            rowDGV.Selected = True
            '            Return False
            '        End If
            '    End If

            '    'optimum stock
            '    TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_OptimumStockTrategy.Index).Value)
            '    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
            '        If rowDGV.Cells(Me.ColGroup_OptimumStockDays.Index).Value = 0 Then
            '            MsgBox("Please specify a positive number for optimum stock days in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '            MsgBoxStyle.Information, "Information")
            '            Me.TabControl1.SelectedIndex = 0
            '            Me.dgvItemGroup.ClearSelection()
            '            rowDGV.Selected = True
            '            Return False
            '        End If
            '    ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
            '        If rowDGV.Cells(Me.ColGroup_OptimumStockQty.Index).Value = 0 Then
            '            MsgBox("Please specify a positive number for optimum stock qty in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '            MsgBoxStyle.Information, "Information")
            '            Me.TabControl1.SelectedIndex = 0
            '            Me.dgvItemGroup.ClearSelection()
            '            rowDGV.Selected = True
            '            Return False
            '        End If
            '    End If

            '    'lead time stock
            '    TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_LeadTimeStrategy.Index).Value)
            '    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
            '        If rowDGV.Cells(Me.ColGroup_LeadTimeDays.Index).Value = 0 Then
            '            MsgBox("Please specify a positive number for lead time stock days in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '            MsgBoxStyle.Information, "Information")
            '            Me.TabControl1.SelectedIndex = 0
            '            Me.dgvItemGroup.ClearSelection()
            '            rowDGV.Selected = True
            '            Return False
            '        End If
            '    ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
            '        If rowDGV.Cells(Me.ColGroup_LeadTimeQty.Index).Value = 0 Then
            '            MsgBox("Please specify a positive number for lead time stock qty in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '            MsgBoxStyle.Information, "Information")
            '            Me.TabControl1.SelectedIndex = 0
            '            Me.dgvItemGroup.ClearSelection()
            '            rowDGV.Selected = True
            '            Return False
            '        End If
            '    End If
            'Next

            '12. jika ada baris dalam product group yang hanya ditentukan levelnya saja, kodenya kosong
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                    MsgBox("Please specify a valid product group/product.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_Scope
                    Me.TabControl2.SelectedTab = Me.TabPage_ProductGroup
                    Me.dgvItemProductGroup.ClearSelection()
                    rowDGV.Selected = True
                    Return False
                End If
            Next

            '13. kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.obj.TimeStamp <> Me.obj.GetTimestamp(Me.obj.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.obj.CariDetail(Me.obj.Primary, Me.dgvItemGroup, _
                        Me.dgvItemSupplier, Me.dgvItemProduct, Me.dgvItemProductGroup, Me, _
                        Me.TBLItem_Invoice, Me.TBLItem_SR, Me.TBLItem_Bundling, Me.TBLItem_Adjustment)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            Me.dgvItemSupplier.AllowUserToAddRows = False
            Me.dgvItemGroup.AllowUserToAddRows = False
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.obj
                .ProductDescFilter = Me.TxtProductDescFilter.Text
                .Remarks = Me.TxtRemarks.Text
                .StartPeriod = Me.DtpFrom.Value
                .StopBeforeRefValue = Me.chkStopBeforeRefValue.Checked
                .TanggalTransaksi = Me.dtpDate.Value
                .Tipe = Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text)
                .CodeNoFilter = Me.txtProductCodeFilter.Text
                .Desc = Me.txtDescription.Text
                .DocumentNo = Me.txtDocumentNo.Text
                .EndPeriod = Me.dtpTo.Value
                .IncludeAdjustment = Me.chkAllIncludeAdjustment.Checked
                .IncludeBundling = Me.chkAllIncludeBundling.Checked
                .IncludeSales = Me.chkAllIncludeSales.Checked
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Dim ObjLocation As New clsMasterStorageArea

        Try
            With Me.obj
                Me.txtPrimary.Text = .Primary

                Me.TxtProductDescFilter.Text = .ProductDescFilter
                Me.TxtRemarks.Text = .Remarks
                Me.DtpFrom.Value = .StartPeriod
                Me.chkStopBeforeRefValue.Checked = .StopBeforeRefValue
                Me.dtpDate.Value = .TanggalTransaksi
                Me.cboGroupBy.Text = Me.obj.GetTipeTransaksiStr(.Tipe)
                Me.txtProductCodeFilter.Text = .CodeNoFilter
                Me.txtDescription.Text = .Desc
                Me.txtDocumentNo.Text = .DocumentNo
                Me.dtpTo.Value = .EndPeriod
                Me.chkAllIncludeAdjustment.Checked = .IncludeAdjustment
                Me.chkAllIncludeBundling.Checked = .IncludeBundling
                Me.chkAllIncludeSales.Checked = .IncludeSales

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With

            SetKolomGroup_LeadTime()
            SetKolomGroup_OptimumStock()
            SetKolomGroup_SafetyStock()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjLocation = Nothing
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()

            Me.DtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value)
            Me.dtpTo.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpTo.Value)

            Me.dgvItemGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProduct.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemSupplier.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemInvoiceException.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.SetKolomGroup_LeadTime()
            Me.SetKolomGroup_OptimumStock()
            Me.SetKolomGroup_SafetyStock()

            Me.TabControl2.TabPages.Remove(Me.TabPage_InvoiceException)
            If Me.chkShowDetail.Enabled = True Then
                If Me.ObjSettingUser.GetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanningGrouping, _
                ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanningGrouping_ShowDetail, Me.UserCode) = "1" Then
                    Me.chkShowDetail.Checked = True
                Else
                    Me.chkShowDetail.Checked = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.SimpanBaru(Me.dgvItemGroup, Me.dgvItemSupplier, Me.dgvItemProductGroup, Me.dgvItemProduct, Me, Me.UserCode, _
                            Me.TBLItem_Invoice, Me.TBLItem_SR, Me.TBLItem_Bundling, Me.TBLItem_Adjustment) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.SimpanPerubahan(Me.dgvItemGroup, Me.dgvItemSupplier, Me.dgvItemProductGroup, Me.dgvItemProduct, Me, Me.UserCode, _
                            Me.TBLItem_Invoice, Me.TBLItem_SR, Me.TBLItem_Bundling, Me.TBLItem_Adjustment) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                End If
            Else
                Me.dgvItemProductGroup.AllowUserToAddRows = True
                Me.dgvItemSupplier.AllowUserToAddRows = True
                Me.dgvItemGroup.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemarks.Text = "" And Me.dgvItemSupplier.RowCount <= 1 And Me.dgvItemProductGroup.RowCount <= 1 And _
            Me.dgvItemGroup.RowCount <= 1 And Me.dgvItemProduct.RowCount = 0 And Me.dtpDate.Value.Date = Today And _
            Me.txtDocumentNo.Text = "" And Me.txtDescription.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.obj.CariDetail(Me.txtPrimary.Text, Me.dgvItemGroup, Me.dgvItemSupplier, Me.dgvItemProduct, Me.dgvItemProductGroup, Me, _
                        Me.TBLItem_Invoice, Me.TBLItem_SR, Me.TBLItem_Bundling, Me.TBLItem_Adjustment)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            Me.txtDocumentNo.Focus()
            Me.obj.ResetObject(Me.dgvItemGroup, Me.dgvItemSupplier, Me.dgvItemProductGroup, Me.dgvItemProduct)
            Me.chkAllIncludeAdjustment.Checked = True
            Me.chkAllIncludeBundling.Checked = True
            Me.chkAllIncludeSales.Checked = True

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            Me.TabControl1.SelectedIndex = 0

            Me.SetCboGroupBy()
            SetKolomGroup_LeadTime()
            SetKolomGroup_OptimumStock()
            SetKolomGroup_SafetyStock()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_MaterialPlanningDashboard_Data1_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                If Me.obj.CariDetail(hj.PrimaryTransaksi, _
                Me.dgvItemGroup, Me.dgvItemSupplier, Me.dgvItemProduct, Me.dgvItemProductGroup, Me, _
                Me.TBLItem_Invoice, Me.TBLItem_SR, Me.TBLItem_Bundling, Me.TBLItem_Adjustment) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()

                    Me.SetKolomGroup_LeadTime()
                    Me.SetKolomGroup_LeadTime_ItemProduct()
                    Me.SetKolomGroup_OptimumStock()
                    Me.SetKolomGroup_OptimumStock_ItemProduct()
                    Me.SetKolomGroup_SafetyStock()
                    Me.SetKolomGroup_SafetyStock_ItemProduct()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                If Me.obj.ApakahBolehEdit(Me.StrictEdit, Me.UserCode, Me.obj.Primary) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                    Me.dgvItemGroup.Focus()
                    Me.SetCboGroupBy()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtDocumentNo.Text <> "" Then
                If Me.obj.ApakahBolehHapus(Me.StrictDelete, Me.UserCode, Me.txtPrimary.Text) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.obj.Hapus(Me.txtPrimary.Text, Me.StrictDelete, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtDocumentNo.Text <> "" Or _
                Me.txtDescription.Text <> "" Or Me.dgvItemGroup.Rows.Count > 1 Or _
                Me.dgvItemProduct.Rows.Count > 0 Or Me.dgvItemProductGroup.Rows.Count > 1 Or _
                Me.dgvItemSupplier.Rows.Count > 1 Then
                    Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ApakahDataBenar() = True Then
                            Me.SalinKeObjek()
                            If baru = True Then
                                'simpan baru
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.obj.SimpanBaru(Me.dgvItemGroup, Me.dgvItemSupplier, Me.dgvItemProductGroup, Me.dgvItemProduct, Me, Me.UserCode, _
                                Me.TBLItem_Invoice, Me.TBLItem_SR, Me.TBLItem_Bundling, Me.TBLItem_Adjustment) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            Else
                                'simpan perubahan
                                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                If Me.obj.SimpanPerubahan(Me.dgvItemGroup, Me.dgvItemSupplier, Me.dgvItemProductGroup, Me.dgvItemProduct, Me, Me.UserCode, _
                                Me.TBLItem_Invoice, Me.TBLItem_SR, Me.TBLItem_Bundling, Me.TBLItem_Adjustment) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnEdit.Focus()
                                Else
                                    e.Cancel = True
                                End If
                            End If
                        Else
                            e.Cancel = True
                        End If
                    ElseIf hasil = MsgBoxResult.Cancel Then
                        e.Cancel = True
                    End If
                Else
                    'langsung tutup
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemGroup.CellBeginEdit
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy

        Try
            'yang bisa diedit
            '1. jika baris baru: code
            If e.RowIndex = Me.dgvItemGroup.RowCount - 1 Then
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                '2. jika baris yang lama: code, desc, ref.value, procurement strategy, remarks,
                If e.ColumnIndex = Me.ColGroup_Code.Index Or _
                e.ColumnIndex = Me.ColGroup_Desc.Index Or _
                e.ColumnIndex = Me.ColGroup_RefValue.Index Or _
                e.ColumnIndex = Me.ColGroup_ProcurementStrategy.Index Or _
                e.ColumnIndex = Me.ColGroup_Remarks.Index Or _
                e.ColumnIndex = Me.ColGroup_OptimumStockTrategy.Index Or _
                e.ColumnIndex = Me.ColGroup_LeadTimeStrategy.Index Or _
                e.ColumnIndex = Me.ColGroup_SafetyStockStrategy.Index Then
                    Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ElseIf e.ColumnIndex = Me.ColGroup_LeadTimeDays.Index Or _
                e.ColumnIndex = Me.ColGroup_LeadTimeQty.Index Then
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_LeadTimeStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        'yang bisa diedit: qty
                        If e.ColumnIndex = Me.ColGroup_LeadTimeQty.Index Then
                            Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        'yang bisa diedit: days
                        If e.ColumnIndex = Me.ColGroup_LeadTimeDays.Index Then
                            Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGroup_SafetyStockDays.Index Or _
                e.ColumnIndex = Me.ColGroup_SafetyStockQty.Index Then
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        'yang bisa diedit: qty
                        If e.ColumnIndex = Me.ColGroup_SafetyStockQty.Index Then
                            Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        'yang bisa diedit: days
                        If e.ColumnIndex = Me.ColGroup_SafetyStockDays.Index Then
                            Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGroup_OptimumStockDays.Index Or _
                e.ColumnIndex = Me.ColGroup_OptimumStockQty.Index Then
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockTrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        'yang bisa diedit: qty
                        If e.ColumnIndex = Me.ColGroup_OptimumStockQty.Index Then
                            Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        'yang bisa diedit: days
                        If e.ColumnIndex = Me.ColGroup_OptimumStockDays.Index Then
                            Me.buffer = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemGroup.CellEndEdit
        Dim DummyIndex As Integer = -1
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy
        Dim TipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage
        Dim rowDGVBefore As DataGridViewRow
        Dim TipeProcurement As Core_Module.clsMasterProduct.EnumProcurementStrategy = clsMasterProduct.EnumProcurementStrategy.Build_To_Order

        Try
            If IsNothing(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColGroup_Code.Index Then
                    'kalo kode akun dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColGroup_RefValue.Index Or _
                e.ColumnIndex = Me.ColGroup_Desc.Index Or _
                e.ColumnIndex = Me.ColGroup_OptimumStockDays.Index Or _
                e.ColumnIndex = Me.ColGroup_OptimumStockQty.Index Or _
                e.ColumnIndex = Me.ColGroup_LeadTimeDays.Index Or _
                e.ColumnIndex = Me.ColGroup_LeadTimeQty.Index Or _
                e.ColumnIndex = Me.ColGroup_SafetyStockDays.Index Or _
                e.ColumnIndex = Me.ColGroup_SafetyStockQty.Index Then
                    'desc, ref.value
                    'optimum stock days, optimum stock qty, lead time stock days, lead time stock qty
                    'safety stock days, safety stock qty
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColGroup_Code.Index Then
                    'code
                    'pastikan ditrim dan huruf besar, max 250 karakter
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'pastikan tidak double dengan yang lainnya
                    DummyIndex = Me.GetIndex_ItemGroup_ByCode(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                    If DummyIndex <> -1 Then
                        MsgBox("Group code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                        If Me.buffer <> "" Then
                            Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        Else
                            Me.DummyIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                    End If
                    'tidak boleh ada karakter aneh2: %, ', *, #, "
                    If Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'kalo sampe sini maka datanya sudah bener, update ke item product
                    Me.UpdateGroupPadaDGVProduct(Me.dgvItemGroup.Item(Me.ColGroup_Identifier.Index, e.RowIndex).Value, _
                    Me.dgvItemGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value, _
                    Me.dgvItemGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value)
                ElseIf e.ColumnIndex = Me.ColGroup_Desc.Index Or _
                e.ColumnIndex = Me.ColGroup_Remarks.Index Then
                    'desc, remarks
                    'pastikan ditrim dan huruf besar, max 250 karakter
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'pastikan tidak double dengan yang lainnya
                    If e.ColumnIndex = Me.ColGroup_Desc.Index Then
                        DummyIndex = Me.GetIndex_ItemGroup_ByDesc(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                        If DummyIndex <> -1 Then
                            MsgBox("Group desc duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'kalo sampe sini maka datanya sudah bener, update ke item product
                    If e.ColumnIndex = Me.ColGroup_Desc.Index Then
                        Me.UpdateGroupPadaDGVProduct(Me.dgvItemGroup.Item(Me.ColGroup_Identifier.Index, e.RowIndex).Value, _
                        Me.dgvItemGroup.Item(Me.ColGroup_Code.Index, e.RowIndex).Value, _
                        Me.dgvItemGroup.Item(Me.ColGroup_Desc.Index, e.RowIndex).Value)
                    End If
                ElseIf e.ColumnIndex = Me.ColGroup_ProcurementStrategy.Index Then
                    'procurement strategy - ga ada batasan
                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        '2. jika ada perubahan procurement strategy
                        If Me.buffer <> Me.dgvItemGroup.Item(Me.ColGroup_ProcurementStrategy.Index, e.RowIndex).Value Then
                            If MsgBox("Changing procurement strategy data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGroup_LeadTimeStrategy.Index Then
                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        '2. jika ada perubahan lead time strategy
                        If Me.buffer <> Me.dgvItemGroup.Item(Me.ColGroup_LeadTimeStrategy.Index, e.RowIndex).Value Then
                            If MsgBox("Changing lead time strategy will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    'jika based on usage: qty = 0
                    'jika user defined: days = 0
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_LeadTimeStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        Me.dgvItemGroup.Item(Me.ColGroup_LeadTimeQty.Index, e.RowIndex).Value = 0
                    Else
                        Me.dgvItemGroup.Item(Me.ColGroup_LeadTimeDays.Index, e.RowIndex).Value = 0
                    End If
                    SetKolomGroup_LeadTime()
                ElseIf e.ColumnIndex = Me.ColGroup_OptimumStockTrategy.Index Then
                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        '2. jika ada perubahan optimum stock strategy
                        If Me.buffer <> Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockTrategy.Index, e.RowIndex).Value Then
                            If MsgBox("Changing optimum strategy will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    'jika based on usage: qty = 0
                    'jika user defined: days = 0
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockTrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockQty.Index, e.RowIndex).Value = 0
                    Else
                        Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockDays.Index, e.RowIndex).Value = 0
                    End If
                    SetKolomGroup_OptimumStock()
                ElseIf e.ColumnIndex = Me.ColGroup_SafetyStockStrategy.Index Then
                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        '2. jika ada perubahan safety stock strategy
                        If Me.buffer <> Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockStrategy.Index, e.RowIndex).Value Then
                            If MsgBox("Changing safety stock strategy will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    'jika based on usage: qty = 0
                    'jika user defined: days = 0
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockQty.Index, e.RowIndex).Value = 0
                    Else
                        Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockDays.Index, e.RowIndex).Value = 0
                    End If
                    SetKolomGroup_SafetyStock()
                ElseIf e.ColumnIndex = Me.ColGroup_LeadTimeDays.Index Or _
                e.ColumnIndex = Me.ColGroup_OptimumStockDays.Index Or _
                e.ColumnIndex = Me.ColGroup_SafetyStockDays.Index Then
                    'optimum stock days, lead time stock days, safety stock days
                    'pastikan angka
                    If IsNumeric(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit grouping 0 decimal
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value, 0)
                    'pastikan tidak negatif
                    If CInt(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                        MsgBox("Please specify non-negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        '2. jika ada perubahan data
                        If CInt(Me.buffer) <> CInt(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            If MsgBox("Changing this data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    'pastikan optimum stock > safety stock
                    If e.ColumnIndex = Me.ColGroup_OptimumStockDays.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockStrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            TipeProcurement = Me.objItem.GetProcurementStrategy(Me.dgvItemGroup.Item(Me.ColGroup_ProcurementStrategy.Index, e.RowIndex).Value)
                            If TipeProcurement = clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                                If CInt(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockDays.Index, e.RowIndex).Value) <= _
                                CInt(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days > safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            Else
                                If CInt(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockDays.Index, e.RowIndex).Value) < _
                                CInt(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days => safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColGroup_SafetyStockDays.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockTrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            TipeProcurement = Me.objItem.GetProcurementStrategy(Me.dgvItemGroup.Item(Me.ColGroup_ProcurementStrategy.Index, e.RowIndex).Value)
                            If TipeProcurement = clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                                If CInt(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockDays.Index, e.RowIndex).Value) <= _
                                CInt(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days > safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            Else
                                If CInt(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockDays.Index, e.RowIndex).Value) < _
                                CInt(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days => safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGroup_LeadTimeQty.Index Or _
                e.ColumnIndex = Me.ColGroup_OptimumStockQty.Index Or _
                e.ColumnIndex = Me.ColGroup_SafetyStockQty.Index Then
                    'optimum stock qty, lead time stock qty, safety stock qty
                    'pastikan angka
                    If IsNumeric(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit grouping berdasarkan decimal master product
                    Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value, 6))
                    'pastikan tidak negatif
                    If CInt(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                        MsgBox("Please specify non-negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        '2. jika ada perubahan data
                        If CDec(Me.buffer) <> CDec(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            If MsgBox("Changing this data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    'pastikan optimum stock > safety stock
                    If e.ColumnIndex = Me.ColGroup_OptimumStockQty.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockStrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            If CDec(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockQty.Index, e.RowIndex).Value) <= _
                            CDec(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockQty.Index, e.RowIndex).Value) Then
                                MsgBox("Please make sure optimum stock qty > safety stock qty.", MsgBoxStyle.Information, "Information")
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColGroup_SafetyStockQty.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockTrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            If CDec(Me.dgvItemGroup.Item(Me.ColGroup_OptimumStockQty.Index, e.RowIndex).Value) <= _
                            CDec(Me.dgvItemGroup.Item(Me.ColGroup_SafetyStockQty.Index, e.RowIndex).Value) Then
                                MsgBox("Please make sure optimum stock qty > safety stock qty.", MsgBoxStyle.Information, "Information")
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColGroup_RefValue.Index Then
                    'ref value
                    'pastikan angka
                    If IsNumeric(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'pastikan tidak nol
                    If CDec(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                        MsgBox("Please specify positive number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit grouping
                    TipeTrans = Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text)
                    If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                        'jika persen, maka digit grouping 2
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        'maximal 100%
                        If CDec(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            MsgBox("Maximum value is 100%", MsgBoxStyle.Information, "Information")
                            Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Then
                        'jika currency, maka digit grouping sesuai IDR
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value, Me.ObjCur.GetDecimal_DefaultCurrency)
                    ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Then
                        'jika currency, maka digit grouping sesuai IDR
                        Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value, 6))
                    End If
                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        '2. jika ada perubahan ref. value
                        If CDec(Me.buffer) <> CDec(Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            If MsgBox("Changing ref. value will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    'jika percetage biasa, maka nilai ref. value harus lebih kecil dari baris sebelumnya
                    'dan lebih besar dari baris setelahnya
                    If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Or _
                    TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Then
                        If e.RowIndex > 0 Then
                            rowDGVBefore = Me.dgvItemGroup.Rows(e.RowIndex - 1)
                            If CDec(Me.dgvItemGroup.Item(Me.ColGroup_RefValue.Index, e.RowIndex).Value) >= _
                            CDec(rowDGVBefore.Cells(Me.ColGroup_RefValue.Index).Value) Then
                                MsgBox("Please specify value less than: " + rowDGVBefore.Cells(Me.ColGroup_RefValue.Index).Value.ToString _
                                + ".", MsgBoxStyle.Information, "Information")
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                        If e.RowIndex < Me.dgvItemGroup.Rows.Count - 1 Then
                            rowDGVBefore = Me.dgvItemGroup.Rows(e.RowIndex + 1)
                            If CDec(Me.dgvItemGroup.Item(Me.ColGroup_RefValue.Index, e.RowIndex).Value) <= _
                            CDec(rowDGVBefore.Cells(Me.ColGroup_RefValue.Index).Value) Then
                                MsgBox("Please specify value more than: " + rowDGVBefore.Cells(Me.ColGroup_RefValue.Index).Value.ToString _
                                + ".", MsgBoxStyle.Information, "Information")
                                Me.dgvItemGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function GetIndex_ItemSupplier(ByVal vCode As String, ByVal vRefIndex As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemSupplier.Rows
                If rowDGV.Cells(Me.ColSupplier_Code.Index).Value = vCode And rowDGV.Index <> vRefIndex Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function GetIndex_ItemGroup_ByCode(ByVal vCode As String, ByVal vRefIndex As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                If rowDGV.Cells(Me.ColGroup_Code.Index).Value = vCode And rowDGV.Index <> vRefIndex Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Function GetIndex_ItemGroup_ByDesc(ByVal vDesc As String, ByVal vRefIndex As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                If rowDGV.Cells(Me.ColGroup_Desc.Index).Value = vDesc And rowDGV.Index <> vRefIndex Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvItemGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Function GetNextIdentifier() As String
        Dim rowDGV As DataGridViewRow
        Dim MaxVal As Double = 0
        Dim Index As Integer = -1

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                If rowDGV.Cells(Me.ColGroup_Identifier.Index).Value <> "" Then
                    Index = rowDGV.Cells(Me.ColGroup_Identifier.Index).Value.ToString.LastIndexOf("-")
                    If Index <> -1 Then
                        If MaxVal < CDbl(rowDGV.Cells(Me.ColGroup_Identifier.Index).Value.ToString.Substring(Index + 1)) Then
                            MaxVal = rowDGV.Cells(Me.ColGroup_Identifier.Index).Value.ToString.Substring(Index + 1)
                        End If
                    End If
                End If
            Next

            Return Me.txtPrimary.Text + "-" + (MaxVal + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-0"
        End Try
    End Function

    Private Sub dgvItem_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            'tentukan nilai default
            If Me.TxtRemarks.Enabled = True Then
                rowDGV = Me.dgvItemGroup.Rows(Me.dgvItemGroup.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColGroup_Code.Index).Value = ""
                    .Cells(Me.ColGroup_Desc.Index).Value = ""
                    .Cells(Me.ColGroup_LeadTimeDays.Index).Value = 0
                    .Cells(Me.ColGroup_LeadTimeQty.Index).Value = 0
                    .Cells(Me.ColGroup_LeadTimeStrategy.Index).Value = Me.objItem.GetLeadTimeStrategyStr(ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage)
                    .Cells(Me.ColGroup_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColGroup_OptimumStockDays.Index).Value = 0
                    .Cells(Me.ColGroup_OptimumStockQty.Index).Value = 0
                    .Cells(Me.ColGroup_OptimumStockTrategy.Index).Value = .Cells(Me.ColGroup_LeadTimeStrategy.Index).Value
                    .Cells(Me.ColGroup_SafetyStockDays.Index).Value = 0
                    .Cells(Me.ColGroup_SafetyStockQty.Index).Value = 0
                    .Cells(Me.ColGroup_SafetyStockStrategy.Index).Value = .Cells(Me.ColGroup_LeadTimeStrategy.Index).Value
                    .Cells(Me.ColGroup_Remarks.Index).Value = ""
                    .Cells(Me.ColGroup_RefValue.Index).Value = 0
                    .Cells(Me.ColGroup_ProcurementStrategy.Index).Value = Me.objItem.GetProcurementStrategyStr(clsMasterProduct.EnumProcurementStrategy.Build_To_Stock)
                    .Cells(Me.ColGroup_Identifier.Index).Value = Me.GetNextIdentifier
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateGroupPadaDGVProduct(ByVal vIdentifier As String, ByVal vCode As String, ByVal vDesc As String)
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvItemProduct.Rows.Count > 0 Then
                Me.dgvItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvItemProduct.Rows
                    If rowDGV.Cells(Me.ColCodeNo_IdentifierGroup.Index).Value = vIdentifier Then
                        rowDGV.Cells(Me.ColCodeNo_Group.Index).Value = vCode + " - " + vDesc
                    End If
                Next
                Me.dgvItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemGroup.RowsRemoved
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.SetNumber(Me.dgvItemGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemGroup_UserAddedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemGroup.UserAddedRow
        Try
            Me.cboGroupBy.Enabled = False
            Me.SetKolomGroup_LeadTime()
            Me.SetKolomGroup_OptimumStock()
            Me.SetKolomGroup_SafetyStock()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemGroup)
            'kalo baris sisa yang belum dimasukkin semua, maka cbo bisa di edit kembali
            Me.SetCboGroupBy()
            Me.SetKolomGroup_LeadTime()
            Me.SetKolomGroup_OptimumStock()
            Me.SetKolomGroup_SafetyStock()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCboGroupBy()
        Try
            If Me.dgvItemGroup.AllowUserToAddRows = True Then
                If Me.dgvItemGroup.Rows.Count = 1 Then
                    Me.cboGroupBy.Enabled = True
                Else
                    Me.cboGroupBy.Enabled = False
                End If
            Else
                If Me.dgvItemGroup.Rows.Count = 0 Then
                    Me.cboGroupBy.Enabled = True
                Else
                    Me.cboGroupBy.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus, txtDocumentNo.GotFocus, _
    txtDescription.GotFocus, txtProductCodeFilter.GotFocus, TxtProductDescFilter.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                Dim txt As TextBox = TryCast(sender, TextBox)
                If txt.ReadOnly = True Then
                    txt.BackColor = Me.TxtReadOnly.BackColor
                Else
                    txt.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCodeFilter_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCodeFilter.Enter
        Me.BufferProductCodeFilter = Me.txtProductCodeFilter.Text.Trim
    End Sub

    Private Sub txtProductCodeFilter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCodeFilter.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItemProduct.Rows.Count > 0 Then
                    Me.txtProductCodeFilter.Text = Me.txtProductCodeFilter.Text.Trim
                    If Me.txtProductCodeFilter.Text <> Me.BufferProductCodeFilter Then
                        If MsgBox("Changing this data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.txtProductCodeFilter.Text = Me.BufferProductCodeFilter
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtProductDescFilter_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductDescFilter.Enter
        Me.BufferProductDescFilter = Me.TxtProductDescFilter.Text.Trim
    End Sub

    Private Sub TxtProductDescFilter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtProductDescFilter.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItemProduct.Rows.Count > 0 Then
                    Me.TxtProductDescFilter.Text = Me.TxtProductDescFilter.Text.Trim
                    If Me.TxtProductDescFilter.Text <> Me.BufferProductDescFilter Then
                        If MsgBox("Changing this data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.TxtProductDescFilter.Text = Me.BufferProductDescFilter
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus, txtDocumentNo.LostFocus, _
    txtDescription.LostFocus, txtProductCodeFilter.LostFocus, TxtProductDescFilter.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                Dim txt As TextBox = TryCast(sender, TextBox)
                If txt.ReadOnly = True Then
                    txt.BackColor = Me.TxtReadOnly.BackColor
                Else
                    txt.BackColor = Me.txtPrimary.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbGroup_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCopyGroup.LinkClicked
        Dim hj As New FrmTrans_MaterialPlanningDashboard_Data1_Cari
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtRemarks.Enabled = True Then
                'konfirmasi
                If MsgBox("Copy data from other transaction?" + vbCrLf + "This will delete your current group and reset calculation data.", MsgBoxStyle.YesNo, "Confirm.") = MsgBoxResult.Yes Then
                    hj.BackColorAll = Me.BackColorAll
                    hj.ShowDialog()

                    'pastikan ada yang dipilih
                    If Not hj.PrimaryTransaksi = -1 Then
                        'pastikan tipe transaksi-nya sama
                        Me.cboGroupBy.Text = Me.obj.GetTipeTransaksiStr(Me.obj.GetTipeTransaksi_ByPrimary(hj.PrimaryTransaksi))
                        'copy data
                        Me.objItem.IsiDGV_ItemGroup(hj.PrimaryTransaksi, _
                        Me.dgvItemGroup, Me, Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text), Me.ObjCD)
                        'set identifier, untuk transaksi ini
                        For Each rowDGV In Me.dgvItemGroup.Rows
                            rowDGV.Cells(Me.ColGroup_Identifier.Index).Value = Me.txtPrimary.Text + "-" + (rowDGV.Index + 1).ToString
                        Next
                        Me.SetCboGroupBy()
                        Me.dgvItemProduct.Rows.Clear()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemGroup.Rows.RemoveAt(DummyIndex)
                DummyIndex = -1
            End If
            Me.ObjCD.SetNumber(Me.dgvItemGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSupplier_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemSupplier.CellBeginEdit
        Try
            If e.RowIndex = Me.dgvItemSupplier.Rows.Count - 1 Then
                'baris terakhir hanya supplier code saja yang bisa di edit
                If e.ColumnIndex = Me.ColSupplier_Code.Index Then
                    Me.buffer = Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'bukan baris terakhir, semua bisa edit kecuali no dan nama supplier
                If e.ColumnIndex = Me.ColSupplier_Code.Index Or _
                e.ColumnIndex = Me.ColSupplier_Remarks.Index Then
                    Me.buffer = Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSupplier_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemSupplier.CellEndEdit
        Try
            If IsNothing(Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColSupplier_Remarks.Index Then
                    'Remarks
                    Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColSupplier_Code.Index Then
                    'kalo kode supplier dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColSupplier_Code.Index Then
                    'code
                    'pastikan ditrim dan huruf besar, max 250 karakter
                    Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'cek apakah ada ada dalam master supplier
                    'kemungkinan
                    If Me.ObjSupplier.ApakahSudahAda(Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode supplier ada  --> load data nama + primary
                        Me.dgvItemSupplier.Item(Me.ColSupplier_Primary.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvItemSupplier.Item(Me.ColSupplier_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.dgvItemSupplier.Item(Me.ColSupplier_Primary.Index, e.RowIndex).Value)
                    Else
                        '2. kode supplier yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmCariMasterSupplier
                        ij.txtCari.Text = Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimarySupplier <> -1 Then
                            '2.a. supplier yang dicari ada --> load data deskripsi + primary + kode
                            Me.dgvItemSupplier.Item(Me.ColSupplier_Primary.Index, e.RowIndex).Value = ij.PrimarySupplier
                            Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjSupplier.GetSupplierCode(ij.PrimarySupplier)
                            Me.dgvItemSupplier.Item(Me.ColSupplier_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(ij.PrimarySupplier)
                        ElseIf ij.PrimarySupplier = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            If Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.DummyIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                                ij = Nothing
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                    'pastikan tidak double dengan yang lainnya
                    DummyIndex = Me.GetIndex_ItemSupplier(Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                    If DummyIndex <> -1 Then
                        MsgBox("Supplier code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                        If buffer = "" Then
                            Me.DummyIndex = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemSupplier.Item(Me.ColSupplier_Primary.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.buffer)
                            Me.dgvItemSupplier.Item(Me.ColSupplier_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.dgvItemSupplier.Item(Me.ColSupplier_Primary.Index, e.RowIndex).Value)
                        End If
                        Exit Sub
                    End If

                    'update ke dalam dgv recalculate
                    If Me.dgvItemProduct.Rows.Count > 0 Then
                        If Me.buffer <> "" Then
                            '1. jika ada perubahan kode supplier
                            If Me.dgvItemSupplier.Item(Me.ColSupplier_Code.Index, e.RowIndex).Value <> Me.buffer Then
                                If MsgBox("Changing prefered supplier data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                    Me.dgvItemProduct.Rows.Clear()
                                Else
                                    Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Me.dgvItemSupplier.Item(Me.ColSupplier_Primary.Index, e.RowIndex).Value = Me.ObjSupplier.GetPrimary(Me.buffer)
                                    Me.dgvItemSupplier.Item(Me.ColSupplier_Name.Index, e.RowIndex).Value = Me.ObjSupplier.GetSupplierName(Me.dgvItemSupplier.Item(Me.ColSupplier_Primary.Index, e.RowIndex).Value)
                                End If
                                Exit Sub
                            End If
                        Else
                            '2. jika ada penambahan kode supplier
                            If MsgBox("Changing prefered supplier data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.dgvItemProduct.Rows.Clear()
                            Else
                                Me.DummyIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            End If
                            Exit Sub
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColSupplier_Remarks.Index Then
                    'desc, remarks
                    'pastikan ditrim dan huruf besar, max 250 karakter
                    Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemSupplier.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSupplier_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemSupplier.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TxtRemarks.Enabled = True Then
                rowDGV = Me.dgvItemSupplier.Rows(Me.dgvItemSupplier.RowCount - 1)
                rowDGV.Cells(Me.ColSupplier_Code.Index).Value = ""
                rowDGV.Cells(Me.ColSupplier_Name.Index).Value = ""
                rowDGV.Cells(Me.ColSupplier_No.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColSupplier_Primary.Index).Value = -1
                rowDGV.Cells(Me.ColSupplier_Remarks.Index).Value = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSupplier_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemSupplier.UserDeletedRow
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.SetNumber(Me.dgvItemSupplier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.objItem.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If

            'update ke dalam dgv recalculate
            If Me.dgvItemProduct.Rows.Count > 0 Then
                If Me.buffer <> "" Then
                    '1. jika ada perubahan kode product group
                    If Me.dgvItemProductGroup.Item(Me.ColSupplier_Code.Index, e.RowIndex).Value <> Me.buffer Then
                        If MsgBox("Changing product group data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.buffer, vProductGroup_Level)
                            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                        End If
                        Exit Sub
                    End If
                Else
                    '2. jika ada penambahan kode product group
                    If MsgBox("Changing product group data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.dgvItemProduct.Rows.Clear()
                    Else
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                    End If
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If

            'update ke dalam dgv recalculate
            If Me.dgvItemProduct.Rows.Count > 0 Then
                If Me.buffer <> "" Then
                    '1. jika ada perubahan kode product group
                    If Me.dgvItemProductGroup.Item(Me.ColSupplier_Code.Index, e.RowIndex).Value <> Me.buffer Then
                        If MsgBox("Changing product group data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.buffer)
                            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                        End If
                        Exit Sub
                    End If
                Else
                    '2. jika ada penambahan kode product group
                    If MsgBox("Changing product group data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.dgvItemProduct.Rows.Clear()
                    Else
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                    End If
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If

            'update ke dalam dgv recalculate
            If Me.dgvItemProduct.Rows.Count > 0 Then
                If Me.buffer <> "" Then
                    '1. jika ada perubahan kode product group
                    If Me.dgvItemProductGroup.Item(Me.ColSupplier_Code.Index, e.RowIndex).Value <> Me.buffer Then
                        If MsgBox("Changing product group data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.buffer)
                            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                        End If
                        Exit Sub
                    End If
                Else
                    '2. jika ada penambahan kode product group
                    If MsgBox("Changing product group data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.dgvItemProduct.Rows.Clear()
                    Else
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                    End If
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemSupplier.Rows.RemoveAt(Me.DummyIndex)
                Me.DummyIndex = -1
            End If
            Me.ObjCD.SetNumber(Me.dgvItemSupplier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolomGroup_LeadTime()
        Dim rowDGV As DataGridViewRow
        Dim AdaUserDefined As Boolean = False
        Dim AdaBasedOnUsage As Boolean = False
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_LeadTimeStrategy.Index).Value)
                If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                    If Me.dgvItemGroup.Rows.Count = 1 Then
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            AdaUserDefined = True
                        ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            AdaBasedOnUsage = True
                        End If
                    Else
                        If Me.dgvItemGroup.AllowUserToAddRows = False Then
                            If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                AdaUserDefined = True
                            ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                AdaBasedOnUsage = True
                            End If
                        End If
                    End If
                Else
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        AdaUserDefined = True
                    ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        AdaBasedOnUsage = True
                    End If
                End If
            Next
            If AdaBasedOnUsage = True Then
                Me.ColGroup_LeadTimeDays.Visible = True
            Else
                Me.ColGroup_LeadTimeDays.Visible = False
            End If
            If AdaUserDefined = True Then
                Me.ColGroup_LeadTimeQty.Visible = True
            Else
                Me.ColGroup_LeadTimeQty.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolomGroup_LeadTime_ItemProduct()
        Dim rowDGV As DataGridViewRow
        Dim AdaUserDefined As Boolean = False
        Dim AdaBasedOnUsage As Boolean = False
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage

        Try
            If Me.chkShowDetail.Checked = True Then
                For Each rowDGV In Me.dgvItemProduct.Rows
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColCodeNo_LeadTimeStrategy.Index).Value)
                    If rowDGV.Index = Me.dgvItemProduct.Rows.Count - 1 Then
                        If Me.dgvItemProduct.Rows.Count = 1 Then
                            If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                AdaUserDefined = True
                            ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                AdaBasedOnUsage = True
                            End If
                        Else
                            If Me.dgvItemProduct.AllowUserToAddRows = False Then
                                If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                    AdaUserDefined = True
                                ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                    AdaBasedOnUsage = True
                                End If
                            End If
                        End If
                    Else
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            AdaUserDefined = True
                        ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            AdaBasedOnUsage = True
                        End If
                    End If
                Next
                If AdaBasedOnUsage = True Then
                    Me.ColCodeNo_LeadTimeDays.Visible = True
                Else
                    Me.ColCodeNo_LeadTimeDays.Visible = False
                End If
                If AdaUserDefined = True Then
                    Me.ColCodeNo_LeadTimeQty.Visible = True
                Else
                    Me.ColCodeNo_LeadTimeQty.Visible = False
                End If
            Else
                Me.ColCodeNo_LeadTimeQty.Visible = False
                Me.ColCodeNo_LeadTimeDays.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolomGroup_OptimumStock()
        Dim rowDGV As DataGridViewRow
        Dim AdaUserDefined As Boolean = False
        Dim AdaBasedOnUsage As Boolean = False
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_OptimumStockTrategy.Index).Value)
                If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                    If Me.dgvItemGroup.Rows.Count = 1 Then
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            AdaUserDefined = True
                        ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            AdaBasedOnUsage = True
                        End If
                    Else
                        If Me.dgvItemGroup.AllowUserToAddRows = False Then
                            If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                AdaUserDefined = True
                            ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                AdaBasedOnUsage = True
                            End If
                        End If
                    End If
                Else
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        AdaUserDefined = True
                    ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        AdaBasedOnUsage = True
                    End If
                End If
            Next
            If AdaBasedOnUsage = True Then
                Me.ColGroup_OptimumStockDays.Visible = True
            Else
                Me.ColGroup_OptimumStockDays.Visible = False
            End If
            If AdaUserDefined = True Then
                Me.ColGroup_OptimumStockQty.Visible = True
            Else
                Me.ColGroup_OptimumStockQty.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolomGroup_OptimumStock_ItemProduct()
        Dim rowDGV As DataGridViewRow
        Dim AdaUserDefined As Boolean = False
        Dim AdaBasedOnUsage As Boolean = False
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage

        Try
            If Me.chkShowDetail.Checked = True Then
                For Each rowDGV In Me.dgvItemProduct.Rows
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColCodeNo_OptimumStockStrategy.Index).Value)
                    If rowDGV.Index = Me.dgvItemProduct.Rows.Count - 1 Then
                        If Me.dgvItemProduct.Rows.Count = 1 Then
                            If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                AdaUserDefined = True
                            ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                AdaBasedOnUsage = True
                            End If
                        Else
                            If Me.dgvItemProduct.AllowUserToAddRows = False Then
                                If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                    AdaUserDefined = True
                                ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                    AdaBasedOnUsage = True
                                End If
                            End If
                        End If
                    Else
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            AdaUserDefined = True
                        ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            AdaBasedOnUsage = True
                        End If
                    End If
                Next
                If AdaBasedOnUsage = True Then
                    Me.ColCodeNo_OptimumStockDays.Visible = True
                Else
                    Me.ColCodeNo_OptimumStockDays.Visible = False
                End If
                If AdaUserDefined = True Then
                    Me.ColCodeNo_OptimumStockQty.Visible = True
                Else
                    Me.ColCodeNo_OptimumStockQty.Visible = False
                End If
            Else
                Me.ColCodeNo_OptimumStockQty.Visible = False
                Me.ColCodeNo_OptimumStockDays.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolomGroup_SafetyStock()
        Dim rowDGV As DataGridViewRow
        Dim AdaUserDefined As Boolean = False
        Dim AdaBasedOnUsage As Boolean = False
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_SafetyStockStrategy.Index).Value)
                If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                    If Me.dgvItemGroup.Rows.Count = 1 Then
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            AdaUserDefined = True
                        ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            AdaBasedOnUsage = True
                        End If
                    Else
                        If Me.dgvItemGroup.AllowUserToAddRows = False Then
                            If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                AdaUserDefined = True
                            ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                AdaBasedOnUsage = True
                            End If
                        End If
                    End If
                Else
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        AdaUserDefined = True
                    ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        AdaBasedOnUsage = True
                    End If
                End If
            Next
            If AdaBasedOnUsage = True Then
                Me.ColGroup_SafetyStockDays.Visible = True
            Else
                Me.ColGroup_SafetyStockDays.Visible = False
            End If
            If AdaUserDefined = True Then
                Me.ColGroup_SafetyStockQty.Visible = True
            Else
                Me.ColGroup_SafetyStockQty.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetKolomGroup_SafetyStock_ItemProduct()
        Dim rowDGV As DataGridViewRow
        Dim AdaUserDefined As Boolean = False
        Dim AdaBasedOnUsage As Boolean = False
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage

        Try
            If Me.chkShowDetail.Checked = True Then
                For Each rowDGV In Me.dgvItemProduct.Rows
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColCodeNo_SafetyStockStrategy.Index).Value)
                    If rowDGV.Index = Me.dgvItemProduct.Rows.Count - 1 Then
                        If Me.dgvItemProduct.Rows.Count = 1 Then
                            If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                AdaUserDefined = True
                            ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                AdaBasedOnUsage = True
                            End If
                        Else
                            If Me.dgvItemProduct.AllowUserToAddRows = False Then
                                If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                                    AdaUserDefined = True
                                ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                                    AdaBasedOnUsage = True
                                End If
                            End If
                        End If
                    Else
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            AdaUserDefined = True
                        ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            AdaBasedOnUsage = True
                        End If
                    End If
                Next
                If AdaBasedOnUsage = True Then
                    Me.ColCodeNo_SafetyStockStrategyDays.Visible = True
                Else
                    Me.ColCodeNo_SafetyStockStrategyDays.Visible = False
                End If
                If AdaUserDefined = True Then
                    Me.ColCodeNo_SafetyStockQty.Visible = True
                Else
                    Me.ColCodeNo_SafetyStockQty.Visible = False
                End If
            Else
                Me.ColCodeNo_SafetyStockQty.Visible = False
                Me.ColCodeNo_SafetyStockStrategyDays.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboGroupBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGroupBy.SelectedIndexChanged
        Dim TipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage

        Try
            TipeTrans = Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text)
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                Me.ColGroup_RefValue.HeaderText = "Ref Value (%)"
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Then
                Me.ColGroup_RefValue.HeaderText = "Ref Value (" + Me.ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + ")"
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Then
                Me.ColGroup_RefValue.HeaderText = "Ref Value (Unit)"
            End If

            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Then
                Me.chkStopBeforeRefValue.Visible = True
                Me.chkStopBeforeRefValue.TabStop = True
            Else
                Me.chkStopBeforeRefValue.Visible = False
                Me.chkStopBeforeRefValue.TabStop = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbRecalculate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbRecalculate.Click
        Dim ij As New FrmTrans_MaterialPlanningDashboard_Data1_Exclusion
        Dim TipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage
        Dim TotalPerc As Decimal = 0
        Dim rowDGV, rowDgvBefore As DataGridViewRow
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_prefered_supplier

        Try
            Me.dgvItemGroup.AllowUserToAddRows = False
            Me.dgvItemSupplier.AllowUserToAddRows = False
            Me.dgvItemProductGroup.AllowUserToAddRows = False

            'harus ada daftar invoice yang di exclude untuk memfasilitasi penjualan yang diluar kebiasaan, sehingga tidak over stock
            'jadi user kalo tahu, ada invoice yang diluar kebiasaan dalam periode ybs, user bisa tidak memasukkan invoice tersebut dalam perhitungan
            'misal kontrak besar yang terjadi hanya seumur hidup sekali
            'LATER

            '2. jika salah satu dari dgv berikut kosong: group
            If Me.dgvItemGroup.Rows.Count = 0 Then
                MsgBox("Please specify at least 1 group.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 0
                Me.dgvItemGroup.Focus()
                Exit Sub
            End If
            '2.b. jika scope tidak terdefinisi
            If Me.dgvItemProductGroup.Rows.Count = 0 And Me.dgvItemSupplier.Rows.Count = 0 And Me.txtProductCodeFilter.Text = "" And Me.TxtProductDescFilter.Text = "" Then
                Me.TabControl1.SelectedTab = Me.TabPage_Scope
                MsgBox("Please specify calculation scope", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '3. period from harus lebih kecil dari period to
            If Me.DtpFrom.Value > Me.dtpTo.Value Then
                MsgBox("Start period must be smaller than end period.", MsgBoxStyle.Information, "Information")
                Me.DtpFrom.Focus()
                Exit Sub
            End If

            '5. jika include sales, include adjustment, dan include bundling semuanya bernilai 0
            If Me.chkAllIncludeSales.Checked = False And Me.chkAllIncludeBundling.Checked = False And Me.chkAllIncludeAdjustment.Checked = False Then
                MsgBox("Please make sure at least 1 of these data is checked: 'Include Sales', 'Include Bundling', or 'Include Adjustment'.", _
                MsgBoxStyle.Information, "Information")
                Me.chkAllIncludeAdjustment.Focus()
                Exit Sub
            End If

            '6. untuk tipe transaksi cumulative percetage, total cumulative percetage harus 100 %
            TipeTrans = Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text)
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Then
                For Each rowDGV In dgvItemGroup.Rows
                    TotalPerc += CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value)
                Next
                If TotalPerc <> 100 Then
                    MsgBox("Please make sure the sum of ref. value data is 100%.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.dgvItemGroup.Focus()
                    Me.dgvItemGroup.ClearSelection()
                    Exit Sub
                End If
            End If


            '7. untuk tipe transaksi percetage, nilai tidak boleh melebih 100%, baris yang atas harus selalu lebih besar nilainya dari baris yang dibawahnya
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                For Each rowDGV In dgvItemGroup.Rows
                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) > 100 Then
                        MsgBox("Please make sure maximum ref. value data is always < 100%.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvItemGroup.Focus()
                        Me.dgvItemGroup.ClearSelection()
                        rowDGV.Selected = True
                        Exit Sub
                    End If
                Next
            End If
            'harus selalu urut yang paling atas yang paling besar
            For Each rowDGV In dgvItemGroup.Rows
                If rowDGV.Index > 0 Then
                    rowDgvBefore = dgvItemGroup.Rows(rowDGV.Index - 1)
                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) >= CDec(rowDgvBefore.Cells(Me.ColGroup_RefValue.Index).Value) Then
                        MsgBox("Please make sure ref. value data is smaller than the previous row.", MsgBoxStyle.Information, "Information")
                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvItemGroup.Focus()
                        Me.dgvItemGroup.ClearSelection()
                        rowDGV.Selected = True
                        Exit Sub
                    End If
                End If
            Next

            '8. untuk setiap item ref value pada dgv group tidak boleh nilai 0
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Then
                'kecuali baris terakhir
                For Each rowDGV In dgvItemGroup.Rows
                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) = 0 And rowDGV.Index <> Me.dgvItemGroup.RowCount - 1 Then
                        MsgBox("Please make sure ref. value data is always > 0", MsgBoxStyle.Information, "Information")

                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvItemGroup.Focus()
                        Me.dgvItemGroup.ClearSelection()
                        rowDGV.Selected = True

                        Exit Sub
                    End If
                Next
            Else
                For Each rowDGV In dgvItemGroup.Rows
                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) = 0 Then
                        MsgBox("Please make sure ref. value data is always > 0", MsgBoxStyle.Information, "Information")

                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvItemGroup.Focus()
                        Me.dgvItemGroup.ClearSelection()
                        rowDGV.Selected = True

                        Exit Sub
                    End If
                Next
            End If

            '9. jika tipe transaksi melibatkan gross margin, maka harus sudah closing mulai dari start period - end period
            'If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            'TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            'TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Then
            If ObjClosing.GetPrimaryClosing(Me.dtpTo.Value) = -1 Then
                MsgBox("Please make sure there is inventory closing on " + Format(Me.dtpTo.Value, "MMM yyyy"), MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            'End If

            '10. untuk setiap item group, jika strategynya user defined dan qty =0
            '11. untuk setiap item group, jika strategynya based on usage dan days =0
            'For Each rowDGV In dgvItemGroup.Rows
            'safety stock
            'TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_SafetyStockStrategy.Index).Value)
            'If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
            '    If rowDGV.Cells(Me.ColGroup_SafetyStockDays.Index).Value = 0 Then
            '        MsgBox("Please specify a positive number for safety stock days in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '        MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.dgvItemGroup.ClearSelection()
            '        rowDGV.Selected = True
            '        Exit Sub
            '    End If
            'ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
            '    If rowDGV.Cells(Me.ColGroup_SafetyStockQty.Index).Value = 0 Then
            '        MsgBox("Please specify a positive number for safety stock qty in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '        MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.dgvItemGroup.ClearSelection()
            '        rowDGV.Selected = True
            '        Exit Sub
            '    End If
            'End If

            'optimum stock
            'TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_OptimumStockTrategy.Index).Value)
            'If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
            '    If rowDGV.Cells(Me.ColGroup_OptimumStockDays.Index).Value = 0 Then
            '        MsgBox("Please specify a positive number for optimum stock days in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '        MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.dgvItemGroup.ClearSelection()
            '        rowDGV.Selected = True
            '        Exit Sub
            '    End If
            'ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
            '    If rowDGV.Cells(Me.ColGroup_OptimumStockQty.Index).Value = 0 Then
            '        MsgBox("Please specify a positive number for optimum stock qty in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '        MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.dgvItemGroup.ClearSelection()
            '        rowDGV.Selected = True
            '        Exit Sub
            '    End If
            'End If

            'lead time stock
            'TipeStrategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_LeadTimeStrategy.Index).Value)
            'If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
            '    If rowDGV.Cells(Me.ColGroup_LeadTimeDays.Index).Value = 0 Then
            '        MsgBox("Please specify a positive number for lead time stock days in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '        MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.dgvItemGroup.ClearSelection()
            '        rowDGV.Selected = True
            '        Exit Sub
            '    End If
            'ElseIf TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
            '    If rowDGV.Cells(Me.ColGroup_LeadTimeQty.Index).Value = 0 Then
            '        MsgBox("Please specify a positive number for lead time stock qty in group " + rowDGV.Cells(Me.ColGroup_Code.Index).Value.ToString, _
            '        MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.dgvItemGroup.ClearSelection()
            '        rowDGV.Selected = True
            '        Exit Sub
            '    End If
            'End If
            'Next

            '12. jika ada baris dalam product group yang hanya ditentukan levelnya saja, kodenya kosong
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                    MsgBox("Please specify a valid product group/product.", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedTab = Me.TabPage_Scope
                    Me.TabControl2.SelectedTab = Me.TabPage_ProductGroup
                    Me.dgvItemProductGroup.ClearSelection()
                    rowDGV.Selected = True
                    Exit Sub
                End If
            Next

            '13. Jika tipe transaksi = gross margin atau nett sales maka chk include sales harus diinput
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Then
                If Me.chkAllIncludeSales.Checked = False Then
                    MsgBox("Please check include sales in calculation.", MsgBoxStyle.Information, "Information")
                    Me.chkAllIncludeSales.Focus()
                    Exit Sub
                End If
            End If

            'product code yang dimasukan ke dalam daftar exclusion harus di daftarkan terlebih dahulu, jika di setujui, maka product tersebut tidak dimasukkan ke dalam perhitungan
            'hanya jika ada yang di exclude
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                If MsgBox("Recalculate data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'tambahkan baris terakhir jika baris terakhir tidak benilai 0, khusus yang tipe transaksi-nya bukan kumulatif
                    If TipeTrans <> ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage And _
                    TipeTrans <> ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage And _
                    TipeTrans <> ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Then
                        rowDGV = Me.dgvItemGroup.Rows(Me.dgvItemGroup.Rows.Count - 1)
                        If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) > 0 Then
                            Me.dgvItemGroup.Rows.Add()
                            rowDGV = Me.dgvItemGroup.Rows(Me.dgvItemGroup.RowCount - 1)
                            With rowDGV
                                .Cells(Me.ColGroup_Code.Index).Value = "NG"
                                .Cells(Me.ColGroup_Desc.Index).Value = "No Group"
                                .Cells(Me.ColGroup_LeadTimeDays.Index).Value = 0
                                .Cells(Me.ColGroup_LeadTimeQty.Index).Value = 0
                                .Cells(Me.ColGroup_No.Index).Value = rowDGV.Index + 1
                                .Cells(Me.ColGroup_OptimumStockDays.Index).Value = 0
                                .Cells(Me.ColGroup_OptimumStockQty.Index).Value = 0
                                .Cells(Me.ColGroup_SafetyStockDays.Index).Value = 0
                                .Cells(Me.ColGroup_SafetyStockQty.Index).Value = 0
                                .Cells(Me.ColGroup_Remarks.Index).Value = ""
                                .Cells(Me.ColGroup_RefValue.Index).Value = 0
                            End With
                            If rowDGV.Index > 0 Then
                                rowDgvBefore = Me.dgvItemGroup.Rows(rowDGV.Index - 1)
                                With rowDGV
                                    .Cells(Me.ColGroup_ProcurementStrategy.Index).Value = rowDgvBefore.Cells(Me.ColGroup_ProcurementStrategy.Index).Value
                                    .Cells(Me.ColGroup_SafetyStockStrategy.Index).Value = rowDgvBefore.Cells(Me.ColGroup_LeadTimeStrategy.Index).Value
                                    .Cells(Me.ColGroup_OptimumStockTrategy.Index).Value = rowDgvBefore.Cells(Me.ColGroup_OptimumStockTrategy.Index).Value
                                    .Cells(Me.ColGroup_LeadTimeStrategy.Index).Value = rowDgvBefore.Cells(Me.ColGroup_LeadTimeStrategy.Index).Value
                                End With
                            End If
                        End If

                        Me.SetKolomGroup_LeadTime()
                        Me.SetKolomGroup_OptimumStock()
                        Me.SetKolomGroup_SafetyStock()
                    End If

                    Me.Hitung()
                    Me.TabControl1.SelectedTab = Me.TabPage_Product
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            Me.dgvItemGroup.AllowUserToAddRows = True
            Me.dgvItemSupplier.AllowUserToAddRows = True
            Me.dgvItemProductGroup.AllowUserToAddRows = True
        End Try
    End Sub

    Private Sub IsikanTabelDummy(ByVal TblDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable, ByVal vPrimaryAcuan As Double, _
    ByVal vLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel, ByVal DTAProduct1 As DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_product1TableAdapter, _
    ByVal TBLProduct1 As DseTransaksiMaterialPlanningGroup.tblmaster_product1DataTable)
        Dim rowProduct1 As DseTransaksiMaterialPlanningGroup.tblmaster_product1Row
        Dim rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow

        Try
            If vLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                DTAProduct1.FillBy_ProductGroup(TBLProduct1, vPrimaryAcuan)
            ElseIf vLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                DTAProduct1.FillBy_Brand(TBLProduct1, vPrimaryAcuan)
            ElseIf vLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                DTAProduct1.FillBy_SubBrand(TBLProduct1, vPrimaryAcuan)
            End If

            For Each rowProduct1 In TBLProduct1.Rows
                TblDummy.DefaultView.RowFilter = "Primary = " + rowProduct1.Primary.ToString
                If TblDummy.DefaultView.Count = 0 Then
                    rowDummy = TblDummy.Newtblmaster_productRow
                    With rowDummy
                        .Primary = rowProduct1.Primary
                        .Product_Code = rowProduct1.Product_Code
                        .Description = rowProduct1.Description
                        .exclude_on_material_planning = rowProduct1.exclude_on_material_planning
                        .Unit_Code = rowProduct1.Unit_Code
                        ._Decimal = rowProduct1._Decimal
                        .Primary_Unit = rowProduct1.Primary_Unit

                        .Gross_Margin = 0
                        .Gross_Margin_Avg = 0
                        .Gross_Margin_Perc = 0
                        .Identifier_Group = ""
                        .Lead_Time_Days = 0
                        .Lead_Time_Qty = 0
                        .Lead_Time_Strategy = 0
                        .Nett_Sales = 0
                        .Nett_Sales_Avg = 0
                        .Nett_Sales_Perc = 0
                        .Optimum_Stock_Days = 0
                        .Optimum_Stock_Qty = 0
                        .Optimum_Stock_Strategy = 0
                        .Procurement_strategy = 0
                        .Qty = 0
                        .Qty_Avg = 0
                        .Qty_Perc = 0
                        .Override = 0
                        .Urutan_Group = -1
                        .Group_Code = ""
                        .Group_Desc = ""
                        .Safety_Stock_Strategy = 0
                        .Safety_Stock_Days = 0
                        .Safety_Stock_Qty = 0
                    End With
                    TblDummy.Rows.Add(rowDummy)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung()
        Dim rowDGV As DataGridViewRow
        Dim TBLProduct As New DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable
        Dim rowProduct, rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow
        Dim DTAProduct As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_productTableAdapter
        Dim Level As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim TBLDummy As New DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable
        Dim Filter As String = ""
        Dim StartPeriode, EndPeriod As Date
        Dim DTA1 As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_salesinvoice_item_view0TableAdapter
        Dim TBL1 As New DseTransaksiMaterialPlanningGroup.tbltrans_salesinvoice_item_view0DataTable
        Dim row1 As DseTransaksiMaterialPlanningGroup.tbltrans_salesinvoice_item_view0Row
        Dim DTAReturn As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_salesinvoice_item_returnTableAdapter
        Dim TBLReturn As New DseTransaksiMaterialPlanningGroup.tbltrans_salesinvoice_item_returnDataTable
        Dim DTASR_Raw As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_sales_returnTableAdapter
        Dim TBLSR_Raw As New DseTransaksiMaterialPlanningGroup.tbltrans_sales_returnDataTable
        Dim rowSR_Raw As DseTransaksiMaterialPlanningGroup.tbltrans_sales_returnRow
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DV As Data.DataRowView
        Dim DTABundling As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_bundlingTableAdapter
        Dim TBLBundling As New DseTransaksiMaterialPlanningGroup.tbltrans_bundlingDataTable
        Dim rowBundling As DseTransaksiMaterialPlanningGroup.tbltrans_bundlingRow
        Dim DTAAdjustment As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_stockout_adjustTableAdapter
        Dim TBLAdjustment As New DseTransaksiMaterialPlanningGroup.tbltrans_stockout_adjustDataTable
        Dim rowAdjustment As DseTransaksiMaterialPlanningGroup.tbltrans_stockout_adjustRow
        Dim TipeTransaksiOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim MasukException As Boolean = False
        Dim FrmException As New FrmTrans_MaterialPlanningDashboard_Data1_Exception
        Dim DefPrecIDR As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim JmlBulan As Integer = 0
        Dim StartPeriodDummy As Date = Me.DtpFrom.Value
        Dim RefValue As Decimal = 0
        Dim TipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage
        Dim TotalGrossMargin As Decimal = 0
        Dim TotalNettSales As Decimal = 0
        Dim TotalQty As Decimal = 0
        Dim TotalValue As Decimal = 0
        Dim TBLProduct1 As New DseTransaksiMaterialPlanningGroup.tblmaster_product1DataTable
        Dim DTAProduct1 As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_product1TableAdapter

        Try
            Me.Cursor = Cursors.WaitCursor

            DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
            DTAProduct1.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTA1.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAReturn.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTASR_Raw.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTABundling.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAAdjustment.Connection.ConnectionString = DTAProduct.Connection.ConnectionString

            'load data product dahulu, pastikan pertimbangkan exclusion
            TBLDummy.Rows.Clear()
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                TBLProduct.Rows.Clear()

                Level = Me.objItem.GetProductLevel(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                If Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Or _
                Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Or _
                Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                    Me.IsikanTabelDummy(TBLDummy, rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value, Level, DTAProduct1, TBLProduct1)
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                    DTAProduct.FillBy_Category(TBLProduct, rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value)
                ElseIf Level = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                    DTAProduct.FillBy_Primary(TBLProduct, rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value)
                End If

                For Each rowProduct In TBLProduct.Rows
                    With rowProduct
                        .Gross_Margin = 0
                        .Gross_Margin_Avg = 0
                        .Gross_Margin_Perc = 0
                        .Identifier_Group = ""
                        .Lead_Time_Days = 0
                        .Lead_Time_Qty = 0
                        .Lead_Time_Strategy = 0
                        .Nett_Sales = 0
                        .Nett_Sales_Avg = 0
                        .Nett_Sales_Perc = 0
                        .Optimum_Stock_Days = 0
                        .Optimum_Stock_Qty = 0
                        .Optimum_Stock_Strategy = 0
                        .Procurement_strategy = 0
                        .Qty = 0
                        .Qty_Avg = 0
                        .Qty_Perc = 0
                        .Override = 0
                        .Urutan_Group = -1
                        .Group_Code = ""
                        .Group_Desc = ""
                        .Safety_Stock_Strategy = 0
                        .Safety_Stock_Days = 0
                        .Safety_Stock_Qty = 0
                    End With
                    TBLDummy.DefaultView.RowFilter = "Primary = " + rowProduct.Primary.ToString
                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.Newtblmaster_productRow
                        For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                            rowDummy.Item(ij) = rowProduct.Item(ij)
                        Next
                        TBLDummy.Rows.Add(rowDummy)
                    End If
                Next
            Next

            'load data berdasarkan prefered supplier
            For Each rowDGV In Me.dgvItemSupplier.Rows
                TBLProduct.Rows.Clear()
                DTAProduct.FillBy_PrimarySupplier(TBLProduct, rowDGV.Cells(Me.ColSupplier_Primary.Index).Value)
                For Each rowProduct In TBLProduct.Rows
                    With rowProduct
                        .Gross_Margin = 0
                        .Gross_Margin_Avg = 0
                        .Gross_Margin_Perc = 0
                        .Identifier_Group = ""
                        .Lead_Time_Days = 0
                        .Lead_Time_Qty = 0
                        .Lead_Time_Strategy = 0
                        .Nett_Sales = 0
                        .Nett_Sales_Avg = 0
                        .Nett_Sales_Perc = 0
                        .Optimum_Stock_Days = 0
                        .Optimum_Stock_Qty = 0
                        .Optimum_Stock_Strategy = 0
                        .Procurement_strategy = 0
                        .Qty = 0
                        .Qty_Avg = 0
                        .Qty_Perc = 0
                        .Override = 0
                        .Urutan_Group = -1
                        .Group_Code = ""
                        .Group_Desc = ""
                    End With
                    TBLDummy.DefaultView.RowFilter = "Primary = " + rowProduct.Primary.ToString
                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.Newtblmaster_productRow
                        For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                            rowDummy.Item(ij) = rowProduct.Item(ij)
                        Next
                        TBLDummy.Rows.Add(rowDummy)
                    End If
                Next
            Next

            'load berdasarkan product code
            Me.txtProductCodeFilter.Text = Me.txtProductCodeFilter.Text.Trim
            If Me.txtProductCodeFilter.Text <> "" Then
                TBLProduct.Rows.Clear()
                DTAProduct.FillBy_ProductCodeLIKE(TBLProduct, Me.txtProductCodeFilter.Text)
                For Each rowProduct In TBLProduct.Rows
                    With rowProduct
                        .Gross_Margin = 0
                        .Gross_Margin_Avg = 0
                        .Gross_Margin_Perc = 0
                        .Identifier_Group = ""
                        .Lead_Time_Days = 0
                        .Lead_Time_Qty = 0
                        .Lead_Time_Strategy = 0
                        .Nett_Sales = 0
                        .Nett_Sales_Avg = 0
                        .Nett_Sales_Perc = 0
                        .Optimum_Stock_Days = 0
                        .Optimum_Stock_Qty = 0
                        .Optimum_Stock_Strategy = 0
                        .Procurement_strategy = 0
                        .Qty = 0
                        .Qty_Avg = 0
                        .Qty_Perc = 0
                        .Override = 0
                        .Urutan_Group = -1
                        .Group_Code = ""
                        .Group_Desc = ""
                    End With
                    TBLDummy.DefaultView.RowFilter = "Primary = " + rowProduct.Primary.ToString
                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.Newtblmaster_productRow
                        For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                            rowDummy.Item(ij) = rowProduct.Item(ij)
                        Next
                        TBLDummy.Rows.Add(rowDummy)
                    End If
                Next
            End If

            'load berdasarkan product desc
            Me.TxtProductDescFilter.Text = Me.TxtProductDescFilter.Text.Trim
            If Me.TxtProductDescFilter.Text <> "" Then
                TBLProduct.Rows.Clear()
                DTAProduct.FillBy_ProductDescLIKE(TBLProduct, Me.TxtProductDescFilter.Text)
                For Each rowProduct In TBLProduct.Rows
                    With rowProduct
                        .Gross_Margin = 0
                        .Gross_Margin_Avg = 0
                        .Gross_Margin_Perc = 0
                        .Identifier_Group = ""
                        .Lead_Time_Days = 0
                        .Lead_Time_Qty = 0
                        .Lead_Time_Strategy = 0
                        .Nett_Sales = 0
                        .Nett_Sales_Avg = 0
                        .Nett_Sales_Perc = 0
                        .Optimum_Stock_Days = 0
                        .Optimum_Stock_Qty = 0
                        .Optimum_Stock_Strategy = 0
                        .Procurement_strategy = 0
                        .Qty = 0
                        .Qty_Avg = 0
                        .Qty_Perc = 0
                        .Override = 0
                        .Urutan_Group = -1
                        .Group_Code = ""
                        .Group_Desc = ""
                    End With
                    TBLDummy.DefaultView.RowFilter = "Primary = " + rowProduct.Primary.ToString
                    If TBLDummy.DefaultView.Count = 0 Then
                        rowDummy = TBLDummy.Newtblmaster_productRow
                        For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                            rowDummy.Item(ij) = rowProduct.Item(ij)
                        Next
                        TBLDummy.Rows.Add(rowDummy)
                    End If
                Next
            End If

            'untuk setiap baris yang ada di dgv, yang diovveride, pastikan masuk ke dalam scope, jika tidak masuk maka hapuskan baris tersebut
            For ij As Double = Me.dgvItemProduct.Rows.Count - 1 To 0 Step -1
                rowDGV = Me.dgvItemProduct.Rows(ij)
                If rowDGV.Cells(Me.ColCodeNo_Override.Index).Value = Boolean.TrueString Then
                    TBLProduct.DefaultView.RowFilter = "Primary = " + rowDGV.Cells(Me.ColCodeNo_PrimaryCodeNo.Index).Value.ToString
                    If TBLProduct.DefaultView.Count = 0 Then
                        Me.dgvItemProduct.Rows.RemoveAt(ij)
                    End If
                End If
            Next

            'langsung keluar dari fungsi jika memang tidak ada itemnya
            If TBLDummy.Rows.Count = 0 Then
                'kosongkan dgv
                Me.dgvItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItemProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.dgvItemProduct.Rows.Clear()
                Exit Sub
            End If

            'load data penjualan + return
            StartPeriode = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value)
            EndPeriod = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value)

            'isikan tabel invoice yang akan digunakan untuk perhitungan
            Me.TBLItem_Invoice.Rows.Clear()
            Me.TBLItem_SR.Rows.Clear()
            If Me.chkAllIncludeSales.Checked = True Then
                DTA1.Fill(TBL1, EndPeriod, StartPeriode, CInt(Core_Module.clsMasterProduct.EnumTipeProduct.Goods))
                DTAReturn.Fill(TBLReturn, EndPeriod, StartPeriode)
                Me.KurangkanItemReturn(TBL1, TBLReturn)
                'isikan tabel sales return nota return pajak
                DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, EndPeriod, StartPeriode)
                Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, StartPeriode, EndPeriod)

                'pindahkan ke dalam tabel product
                For Each rowDummy In TBLDummy.Rows
                    TBL1.DefaultView.RowFilter = "[Primary Product] = " + rowDummy.Primary.ToString
                    For Each DV In TBL1.DefaultView
                        row1 = DV.Row
                        rowDummy.Gross_Margin += row1.Sales_Amount - row1.Cost
                        rowDummy.Nett_Sales += row1.Sales_Amount
                        If rowDummy.Nett_Sales < 0 Then
                            MsgBox("")
                        End If
                        rowDummy.Qty += Math.Round(row1.Total_Qty, rowDummy._Decimal, MidpointRounding.AwayFromZero)
                        'pindahkan ke table keterangan
                        Me.rowItem_Invoice = Me.TBLItem_Invoice.Newtbltrans_material_planning_group_item_invoiceRow
                        With Me.rowItem_Invoice
                            .cost = row1.Cost
                            .customer_code_historical = row1.Customer_Code
                            .customer_name_historical = row1.Customer_Name
                            .invoice_date_historical = row1.Invoice_Date
                            .invoice_no_historical = row1.No_Transaksi
                            .primary = Me.TBLItem_Invoice.Rows.Count + 1
                            .primary_code_no = row1.Primary_Product
                            .primary_customer = row1.Primary_Customer
                            .primary_invoice = row1.Primary_Invoice
                            .primary_transaksi = Me.txtPrimary.Text
                            .primary_unit = row1.Primary_Unit
                            .qty = row1.Total_Qty
                            .sales_amount = row1.Sales_Amount
                        End With
                        Me.TBLItem_Invoice.Rows.Add(Me.rowItem_Invoice)
                    Next
                    TBLSR_Raw.DefaultView.RowFilter = "[Primary Product] = " + rowDummy.Primary.ToString
                    For Each DV In TBLSR_Raw.DefaultView
                        rowSR_Raw = DV.Row
                        rowDummy.Gross_Margin += (rowSR_Raw.Amount - rowSR_Raw.Cost) * -1
                        rowDummy.Nett_Sales += rowSR_Raw.Amount * -1
                        rowDummy.Qty += Math.Round(rowSR_Raw.Total_Qty * -1, rowDummy._Decimal, MidpointRounding.AwayFromZero)
                        'pindahkan ke table keterangan
                        Me.rowItem_SR = Me.TBLItem_SR.Newtbltrans_material_planning_group_item_srRow
                        With Me.rowItem_SR
                            .cost = rowSR_Raw.Cost
                            .customer_code_historical = rowSR_Raw.Customer_Code
                            .customer_name_historical = rowSR_Raw.Customer_Name
                            .primary = Me.TBLItem_SR.Rows.Count + 1
                            .primary_code_no = rowSR_Raw.Primary_Product
                            .primary_customer = rowSR_Raw.Primary_Customer
                            .primary_sr = rowSR_Raw.Primary_SR
                            .primary_transaksi = Me.txtPrimary.Text
                            .primary_unit = rowSR_Raw.Primary_Unit
                            .qty = rowSR_Raw.Total_Qty
                            .sales_amount = rowSR_Raw.Amount
                            .sr_date_historical = rowSR_Raw.Tanggal_Transaksi
                            .sr_no_historical = rowSR_Raw.SR_Number
                        End With
                        Me.TBLItem_SR.Rows.Add(Me.rowItem_SR)
                    Next
                Next
            End If

            'bundling
            Me.TBLItem_Bundling.Rows.Clear()
            If Me.chkAllIncludeBundling.Checked = True Then
                DTABundling.Fill(TBLBundling, StartPeriode, EndPeriod)
                'pindahkan ke dalam tabel product
                For Each rowDummy In TBLDummy.Rows
                    TBLBundling.DefaultView.RowFilter = "[Primary Product] = " + rowDummy.Primary.ToString
                    For Each DV In TBLBundling.DefaultView
                        rowBundling = DV.Row
                        rowDummy.Qty += Math.Round(rowBundling.Qty, rowDummy._Decimal, MidpointRounding.AwayFromZero)
                        'pindahkan ke tabel keterangan
                        Me.rowItem_Bundling = Me.TBLItem_Bundling.Newtbltrans_material_planning_group_item_bundlingRow
                        With Me.rowItem_Bundling
                            .bundling_date_historical = rowBundling.Tanggal_Transaksi
                            .bundling_no_historical = rowBundling.No_Transaksi
                            .primary = Me.TBLItem_Bundling.Rows.Count + 1
                            .primary_bundling = rowBundling.Primary_Transaksi
                            .primary_code_no = rowBundling.Primary_Product
                            .primary_transaksi = Me.txtPrimary.Text
                            .primary_unit = rowBundling.Primary_Unit
                            .qty = rowBundling.Qty
                        End With
                        Me.TBLItem_Bundling.Rows.Add(Me.rowItem_Bundling)
                    Next
                Next
            End If

            'out adjustment
            Me.TBLItem_Adjustment.Rows.Clear()
            If Me.chkAllIncludeAdjustment.Checked = True Then
                DTAAdjustment.Fill(TBLAdjustment, TipeTransaksiOut, EndPeriod, StartPeriode)
                For Each rowDummy In TBLDummy.Rows
                    TBLAdjustment.DefaultView.RowFilter = "[Primary Product] = " + rowDummy.Primary.ToString
                    For Each DV In TBLAdjustment.DefaultView
                        rowAdjustment = DV.Row
                        rowDummy.Qty += Math.Round(rowAdjustment.Qty, rowDummy._Decimal, MidpointRounding.AwayFromZero)
                        'pindahkan ke tabel keterangan
                        Me.rowItem_Adjustment = Me.TBLItem_Adjustment.Newtbltrans_material_planning_group_item_adjustmentRow
                        With Me.rowItem_Adjustment
                            .adj_date_historical = rowAdjustment.Tanggal_Transaksi
                            .adj_no_historical = rowAdjustment.No_Transaksi
                            .primary = Me.TBLItem_Adjustment.Rows.Count + 1
                            .primary_adj = rowAdjustment.Primary_Transaksi
                            .primary_code_no = rowAdjustment.Primary_Product
                            .primary_transaksi = Me.txtPrimary.Text
                            .primary_unit = rowAdjustment.Primary_Unit
                            .qty = rowAdjustment.Qty
                            .primary_tipe_adjustment = rowAdjustment.PrimaryAdjType
                        End With
                        Me.TBLItem_Adjustment.Rows.Add(Me.rowItem_Adjustment)
                    Next
                Next
            End If

            'kalo sampe ketemu yang nilai qty-nya, nett sales, atau gross marginnya minus, dianggap nol
            'nanti harus diberikan peringatan ketika selesai
            For Each rowDummy In TBLDummy.Rows
                MasukException = False
                If rowDummy.Qty < 0 Then
                    MasukException = True
                End If
                If rowDummy.Nett_Sales < 0 Then
                    MasukException = True
                End If
                If rowDummy.Gross_Margin < 0 Then
                    MasukException = True
                End If
                If MasukException = True Then
                    'demo
                    FrmException.DataGridView2.Rows.Add()
                    rowDGV = FrmException.DataGridView2.Rows(FrmException.DataGridView2.RowCount - 1)
                    With rowDGV
                        .Cells(FrmException.ColCodeNo.Index).Value = rowDummy.Product_Code
                        .Cells(FrmException.ColDesc.Index).Value = rowDummy.Description
                        .Cells(FrmException.ColGrossMargin.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Gross_Margin, DefPrecIDR)
                        .Cells(FrmException.ColNo.Index).Value = rowDGV.Index + 1
                        .Cells(FrmException.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Qty, rowDummy._Decimal)
                        .Cells(FrmException.ColSalesAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Nett_Sales, DefPrecIDR)
                        .Cells(FrmException.ColUnit.Index).Value = rowDummy.Unit_Code
                    End With
                End If
                If rowDummy.Qty < 0 Then
                    rowDummy.Qty = 0
                End If
                If rowDummy.Nett_Sales < 0 Then
                    rowDummy.Nett_Sales = 0
                End If
                If rowDummy.Gross_Margin < 0 Then
                    rowDummy.Gross_Margin = 0
                End If
            Next
            If MasukException = True Then
                'berikan peringatan
                MsgBox("The following product have either minus qty, nett sales, or gross masrgin data. Zero value is assumed for each minus data. Please be aware of it.", _
                MsgBoxStyle.Information, "Information")
                FrmException.ShowDialog()
            End If

            'hitung average
            JmlBulan = 0
            StartPeriodDummy = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value)
            While StartPeriodDummy <= ObjCD.GetFirstDateofTheMonth(Me.dtpTo.Value)
                JmlBulan += 1
                StartPeriodDummy = Me.ObjCD.GetFirstDateofTheMonth(StartPeriodDummy + TimeSpan.FromDays(32))
            End While
            For Each rowDummy In TBLDummy.Rows
                rowDummy.Gross_Margin_Avg = Math.Round(rowDummy.Gross_Margin / JmlBulan, DefPrecIDR, MidpointRounding.AwayFromZero)
                rowDummy.Nett_Sales_Avg = Math.Round(rowDummy.Nett_Sales / JmlBulan, DefPrecIDR, MidpointRounding.AwayFromZero)
                rowDummy.Qty_Avg = Math.Round(rowDummy.Qty / JmlBulan, rowDummy._Decimal, MidpointRounding.AwayFromZero)
            Next

            'hitung percent gross margin
            TotalValue = 0
            For Each rowDummy In TBLDummy.Rows
                TotalValue += rowDummy.Gross_Margin
            Next
            If TotalValue <> 0 Then
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Gross_Margin_Perc = Math.Round(rowDummy.Gross_Margin * 100 / TotalValue, 2, MidpointRounding.AwayFromZero)
                Next
            Else
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Gross_Margin_Perc = 0
                Next
            End If
            TotalGrossMargin = TotalValue
            'hitung percent nett sales
            TotalValue = 0
            For Each rowDummy In TBLDummy.Rows
                TotalValue += rowDummy.Nett_Sales
            Next
            If TotalValue <> 0 Then
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Nett_Sales_Perc = Math.Round(rowDummy.Nett_Sales * 100 / TotalValue, 2, MidpointRounding.AwayFromZero)
                Next
            Else
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Nett_Sales_Perc = 0
                Next
            End If
            TotalNettSales = TotalValue
            'hitung percent qty
            TotalValue = 0
            For Each rowDummy In TBLDummy.Rows
                TotalValue += rowDummy.Qty
            Next
            If TotalValue <> 0 Then
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Qty_Perc = Math.Round(rowDummy.Qty * 100 / TotalValue, 2, MidpointRounding.AwayFromZero)
                Next
            Else
                For Each rowDummy In TBLDummy.Rows
                    rowDummy.Qty_Perc = 0
                Next
            End If
            TotalQty = TotalValue

            'klasifikasikan product
            TBLDummy.DefaultView.RowFilter = ""
            TipeTrans = Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text)
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Then
                'sort berdasarkan gross margin
                TBLDummy.DefaultView.Sort = "[Gross Margin] DESC"
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Then
                'sort berdasarkan nett sales
                TBLDummy.DefaultView.Sort = "[Nett Sales] DESC"
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                'sort berdasarkan nett sales
                TBLDummy.DefaultView.Sort = "[Qty] DESC"
            End If

            TBLProduct.Rows.Clear()
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Then
                Me.Klasifikasi_GrossMarginCumulativePercentage(TBLDummy, DefPrecIDR, TotalGrossMargin)
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Then
                Me.Klasifikasi_NettSalesCumulativePercentage(TBLDummy, DefPrecIDR, TotalNettSales)
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Then
                Me.Klasifikasi_QtyCumulativePercentage(TBLDummy, DefPrecIDR, TotalQty)
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Then
                Me.Klasifikasi_Percentage_Qty(TBLDummy, TipeTrans)
            End If

            'untuk setiap baris yang ada di dgv, yang override, tandai sebagai override, masukkan data hasil override ke dalam data table
            Me.SalinDataOverrideDariDGV(TBLDummy)
            'kosongkan dgv
            Me.dgvItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItemProduct.Rows.Clear()
            'load ke dalam dgv
            Me.SalinDataKeDGVItemProduct(TBLDummy, DefPrecIDR)
            Me.SetKolomGroup_LeadTime_ItemProduct()
            Me.SetKolomGroup_OptimumStock_ItemProduct()
            Me.SetKolomGroup_SafetyStock_ItemProduct()
            Me.dgvItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLProduct.Dispose()
            DTAProduct.Dispose()
            TBLProduct1.Dispose()
            DTAProduct1.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
            DTAAdjustment.Dispose()
            TBLAdjustment.Dispose()
            DTABundling.Dispose()
            TBLBundling.Dispose()
            FrmException.Dispose()
            Me.Cursor = Cursors.Arrow
        End Try
    End Sub

    Private Sub SalinDataKeDGVItemProduct(ByVal TBLDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable, ByVal DefPrecIDR As Integer)
        Dim rowDGV As DataGridViewRow
        Dim DV As Data.DataRowView
        Dim rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow
        Dim TipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage

        Try
            TipeTrans = Me.obj.GetTipeTransaksi(Me.cboGroupBy.Text)

            TBLDummy.DefaultView.RowFilter = ""
            If TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Then
                TBLDummy.DefaultView.Sort = "[Urutan Group] ASC, [Gross Margin] DESC, [Product Code] ASC"
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Then
                TBLDummy.DefaultView.Sort = "[Urutan Group] ASC, [Nett Sales] DESC, [Product Code] ASC"
            ElseIf TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Or _
            TipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                TBLDummy.DefaultView.Sort = "[Urutan Group] ASC, Qty DESC, [Product Code] ASC"
            End If
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                Me.dgvItemProduct.Rows.Add()
                rowDGV = Me.dgvItemProduct.Rows(Me.dgvItemProduct.Rows.Count - 1)
                With rowDummy
                    rowDGV.Cells(Me.ColCodeNo_AverageGrossMargin.Index).Value = ObjCD.DigitGrouping(.Gross_Margin_Avg, DefPrecIDR)
                    rowDGV.Cells(Me.ColCodeNo_AverageNettSales.Index).Value = ObjCD.DigitGrouping(.Nett_Sales_Avg, DefPrecIDR)
                    rowDGV.Cells(Me.ColCodeNo_AverageQty.Index).Value = ObjCD.DigitGrouping(.Qty_Avg, ._Decimal)
                    rowDGV.Cells(Me.ColCodeNo_CodeNo.Index).Value = .Product_Code
                    rowDGV.Cells(Me.ColCodeNo_Desc.Index).Value = .Description
                    rowDGV.Cells(Me.ColCodeNo_GrossMargin.Index).Value = ObjCD.DigitGrouping(.Gross_Margin, DefPrecIDR)
                    rowDGV.Cells(Me.ColCodeNo_GrossMarginPerc.Index).Value = ObjCD.DigitGrouping(.Gross_Margin_Perc, 2)
                    rowDGV.Cells(Me.ColCodeNo_LeadTimeDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.Lead_Time_Days)
                    rowDGV.Cells(Me.ColCodeNo_LeadTimeQty.Index).Value = ObjCD.DigitGrouping(.Lead_Time_Qty, ._Decimal)
                    rowDGV.Cells(Me.ColCodeNo_LeadTimeStrategy.Index).Value = Me.objItem.GetLeadTimeStrategyStr(.Lead_Time_Strategy)
                    rowDGV.Cells(Me.ColCodeNo_NettSales.Index).Value = ObjCD.DigitGrouping(.Nett_Sales, DefPrecIDR)
                    rowDGV.Cells(Me.ColCodeNo_NettSalesPerc.Index).Value = ObjCD.DigitGrouping(.Nett_Sales_Perc, 2)
                    rowDGV.Cells(Me.ColCodeNo_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColCodeNo_OptimumStockDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.Optimum_Stock_Days)
                    rowDGV.Cells(Me.ColCodeNo_OptimumStockQty.Index).Value = ObjCD.DigitGrouping(.Optimum_Stock_Qty, DefPrecIDR)
                    rowDGV.Cells(Me.ColCodeNo_OptimumStockStrategy.Index).Value = Me.objItem.GetLeadTimeStrategyStr(.Optimum_Stock_Strategy)
                    If .Override = 0 Then
                        rowDGV.Cells(Me.ColCodeNo_Override.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(Me.ColCodeNo_Override.Index).Value = Boolean.TrueString
                    End If
                    rowDGV.Cells(Me.ColCodeNo_PrimaryUnit.Index).Value = .Primary_Unit
                    rowDGV.Cells(Me.ColCodeNo_ProcurementStrategy.Index).Value = Me.objItem.GetProcurementStrategyStr(.Procurement_strategy)
                    rowDGV.Cells(Me.ColCodeNo_Qty.Index).Value = ObjCD.DigitGrouping(.Qty, ._Decimal)
                    rowDGV.Cells(Me.ColCodeNo_QtyPerc.Index).Value = ObjCD.DigitGrouping(.Qty_Perc, 2)
                    rowDGV.Cells(Me.ColCodeNo_SafetyStockQty.Index).Value = ObjCD.DigitGrouping(.Safety_Stock_Qty, DefPrecIDR)
                    rowDGV.Cells(Me.ColCodeNo_SafetyStockStrategy.Index).Value = Me.objItem.GetLeadTimeStrategyStr(.Safety_Stock_Strategy)
                    rowDGV.Cells(Me.ColCodeNo_SafetyStockStrategyDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.Safety_Stock_Days)
                    rowDGV.Cells(Me.ColCodeNo_Unit.Index).Value = .Unit_Code
                    rowDGV.Cells(Me.ColCodeNo_Remarks.Index).Value = ""
                    rowDGV.Cells(Me.ColCodeNo_PrimaryCodeNo.Index).Value = .Primary
                    rowDGV.Cells(Me.ColCodeNo_IdentifierGroup.Index).Value = .Identifier_Group
                    rowDGV.Cells(Me.ColCodeNo_Group.Index).Value = .Group_Code + " - " + .Group_Desc
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinDataOverrideDariDGV(ByVal vTBLDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable)
        Dim rowDGV As DataGridViewRow
        Dim rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow

        Try
            For Each rowDGV In Me.dgvItemProduct.Rows
                If rowDGV.Cells(Me.ColCodeNo_Override.Index).Value = Boolean.TrueString Then
                    vTBLDummy.DefaultView.RowFilter = "[Primary] = " + rowDGV.Cells(Me.ColCodeNo_PrimaryCodeNo.Index).Value.ToString
                    If vTBLDummy.DefaultView.Count > 0 Then
                        rowDummy = vTBLDummy.DefaultView(0).Row
                        With rowDummy
                            .Procurement_strategy = Me.objItem.GetProcurementStrategy(rowDGV.Cells(Me.ColCodeNo_ProcurementStrategy.Index).Value)
                            .Optimum_Stock_Strategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColCodeNo_OptimumStockStrategy.Index).Value)
                            .Optimum_Stock_Days = rowDGV.Cells(Me.ColCodeNo_OptimumStockDays.Index).Value
                            .Optimum_Stock_Qty = rowDGV.Cells(Me.ColCodeNo_OptimumStockQty.Index).Value
                            .Safety_Stock_Strategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColCodeNo_SafetyStockStrategy.Index).Value)
                            .Safety_Stock_Days = rowDGV.Cells(Me.ColCodeNo_SafetyStockStrategyDays.Index).Value
                            .Safety_Stock_Qty = rowDGV.Cells(Me.ColCodeNo_SafetyStockQty.Index).Value
                            .Lead_Time_Strategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColCodeNo_LeadTimeStrategy.Index).Value)
                            .Lead_Time_Days = rowDGV.Cells(Me.ColCodeNo_LeadTimeDays.Index).Value
                            .Lead_Time_Qty = rowDGV.Cells(Me.ColCodeNo_LeadTimeQty.Index).Value
                        End With
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Klasifikasi_Percentage_Qty(ByVal TBLDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable, _
    ByVal vTipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe)
        Dim rowDGV As DataGridViewRow
        Dim rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow
        Dim DV As Data.DataRowView
        Dim JmlKonsisten As Integer = 0
        Dim RefValue As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                TBLDummy.DefaultView.RowFilter = "[Identifier Group] = ''"

                For Each DV In TBLDummy.DefaultView
                    rowDummy = DV.Row
                    If vTipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Then
                        RefValue = rowDummy.Gross_Margin_Perc
                    ElseIf vTipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Then
                        RefValue = rowDummy.Nett_Sales_Perc
                    ElseIf vTipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                        RefValue = rowDummy.Qty_Perc
                    ElseIf vTipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Then
                        RefValue = rowDummy.Gross_Margin
                    ElseIf vTipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Then
                        RefValue = rowDummy.Nett_Sales
                    ElseIf vTipeTrans = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Minimum_Value Then
                        RefValue = rowDummy.Qty
                    End If

                    If CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) <= RefValue Then
                        'masuk kriteria, salin data
                        Me.SalinData(rowDummy, rowDGV)
                    Else
                        'jika 4 kali tidak masuk kriteria, karena sudah diurutkan berdasarkan kriteria yang bersesuai sebelumnya
                        'maka bisa diasumsikan semua baris selanjutnya akan tidak masuk kriteria, langsung keluar loop
                        JmlKonsisten += 1
                        If JmlKonsisten > 3 Then
                            Exit For
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Klasifikasi_QtyCumulativePercentage(ByVal TBLDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable, _
    ByVal DefPrecIDR As Integer, ByVal TotalQty As Decimal)
        Dim rowDGV As DataGridViewRow
        Dim RefValue As Decimal = 0
        Dim TotalValue As Decimal = 0
        Dim StartIndex As Integer = 0
        Dim rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow

        Try
            StartIndex = 0
            RefValue = 0
            TotalValue = 0

            For Each rowDGV In Me.dgvItemGroup.Rows
                If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                    RefValue = TotalQty
                Else
                    RefValue += Math.Round((CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) * TotalQty) / 100, DefPrecIDR, MidpointRounding.AwayFromZero)
                End If
                If RefValue = 0 Then
                    If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                        While StartIndex <= TBLDummy.DefaultView.Count - 1
                            rowDummy = TBLDummy.DefaultView(StartIndex).Row
                            'pas, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'baris product selanjutnya pindah ke group selanjutnya
                            StartIndex += 1
                        End While
                    Else
                        'pindah baris selanjutnya pada dgv item group
                    End If
                Else
                    While StartIndex <= TBLDummy.DefaultView.Count - 1
                        rowDummy = TBLDummy.DefaultView(StartIndex).Row
                        TotalValue += rowDummy.Qty

                        If TotalValue = RefValue Then
                            'pas, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'baris product selanjutnya pindah ke group selanjutnya
                            StartIndex += 1
                            Exit While
                        ElseIf TotalValue > RefValue Then
                            If Me.chkStopBeforeRefValue.Checked = True Then
                                'baris ini masuk ke dalam group selanjutnya, tidak menambahkan start index, langsung keluar while
                                'kembali kurangkan total value, akan ditambahkan ketika mengevaluasi baris ini lagi
                                TotalValue -= rowDummy.Qty
                                Exit While
                            Else
                                'baris ini masuk, salin data
                                Me.SalinData(rowDummy, rowDGV)
                                'baris product selanjutnya pindah ke group selanjutnya
                                StartIndex += 1
                                Exit While
                            End If
                        Else
                            'baris ini masuk group ybs, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'cek baris selanjutnya
                            StartIndex += 1
                        End If
                    End While
                End If
            Next

            'klasifikasikan baris yang tidak masuk ke dalam baris terakhir
            If StartIndex <> 0 Then
                While StartIndex <= TBLDummy.DefaultView.Count - 1
                    rowDummy = TBLDummy.DefaultView(StartIndex).Row
                    rowDGV = Me.dgvItemGroup.Rows(Me.dgvItemGroup.Rows.Count - 1)
                    'baris ini masuk group ybs, salin data
                    Me.SalinData(rowDummy, rowDGV)
                    'cek baris selanjutnya
                    StartIndex += 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Klasifikasi_GrossMarginCumulativePercentage(ByVal TBLDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable, _
    ByVal DefPrecIDR As Integer, ByVal TotalGrossMargin As Decimal)
        Dim rowDGV As DataGridViewRow
        Dim RefValue As Decimal = 0
        Dim TotalValue As Decimal = 0
        Dim StartIndex As Integer = 0
        Dim rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow

        Try
            StartIndex = 0
            RefValue = 0
            TotalValue = 0

            For Each rowDGV In Me.dgvItemGroup.Rows
                If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                    RefValue = TotalGrossMargin
                Else
                    RefValue += Math.Round((CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) * TotalGrossMargin) / 100, DefPrecIDR, MidpointRounding.AwayFromZero)
                End If

                If RefValue = 0 Then
                    If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                        While StartIndex <= TBLDummy.DefaultView.Count - 1
                            rowDummy = TBLDummy.DefaultView(StartIndex).Row
                            'pas, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'baris product selanjutnya pindah ke group selanjutnya
                            StartIndex += 1
                        End While
                    Else
                        'pindah baris selanjutnya pada dgv item group
                    End If
                Else
                    While StartIndex <= TBLDummy.DefaultView.Count - 1
                        rowDummy = TBLDummy.DefaultView(StartIndex).Row
                        TotalValue += rowDummy.Gross_Margin

                        If TotalValue = RefValue Then
                            'pas, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'baris product selanjutnya pindah ke group selanjutnya
                            StartIndex += 1
                            Exit While
                        ElseIf TotalValue > RefValue Then
                            If Me.chkStopBeforeRefValue.Checked = True Then
                                'baris ini masuk ke dalam group selanjutnya, tidak menambahkan start index, langsung keluar while
                                'kembali kurangkan total value, akan ditambahkan ketika mengevaluasi baris ini lagi
                                TotalValue -= rowDummy.Gross_Margin
                                Exit While
                            Else
                                'baris ini masuk, salin data
                                Me.SalinData(rowDummy, rowDGV)
                                'baris product selanjutnya pindah ke group selanjutnya
                                StartIndex += 1
                                Exit While
                            End If
                        Else
                            'baris ini masuk group ybs, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'cek baris selanjutnya
                            StartIndex += 1
                        End If
                    End While
                End If
            Next

            'klasifikasikan baris yang tidak masuk ke dalam baris terakhir
            If StartIndex <> 0 Then
                While StartIndex <= TBLDummy.DefaultView.Count - 1
                    rowDummy = TBLDummy.DefaultView(StartIndex).Row
                    rowDGV = Me.dgvItemGroup.Rows(Me.dgvItemGroup.Rows.Count - 1)
                    'baris ini masuk group ybs, salin data
                    Me.SalinData(rowDummy, rowDGV)
                    'cek baris selanjutnya
                    StartIndex += 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Klasifikasi_NettSalesCumulativePercentage(ByVal TBLDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable, _
    ByVal DefPrecIDR As Integer, ByVal TotalNettSales As Decimal)
        Dim rowDGV As DataGridViewRow
        Dim RefValue As Decimal = 0
        Dim TotalValue As Decimal = 0
        Dim StartIndex As Integer = 0
        Dim rowDummy As DseTransaksiMaterialPlanningGroup.tblmaster_productRow

        Try
            StartIndex = 0
            RefValue = 0
            TotalValue = 0

            For Each rowDGV In Me.dgvItemGroup.Rows
                If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                    RefValue = TotalNettSales
                Else
                    RefValue += Math.Round((CDec(rowDGV.Cells(Me.ColGroup_RefValue.Index).Value) * TotalNettSales) / 100, DefPrecIDR, MidpointRounding.AwayFromZero)
                End If

                If RefValue = 0 Then
                    If rowDGV.Index = Me.dgvItemGroup.Rows.Count - 1 Then
                        While StartIndex <= TBLDummy.DefaultView.Count - 1
                            rowDummy = TBLDummy.DefaultView(StartIndex).Row
                            'pas, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'baris product selanjutnya pindah ke group selanjutnya
                            StartIndex += 1
                        End While
                    Else
                        'pindah baris selanjutnya pada dgv item group
                    End If
                Else
                    While StartIndex <= TBLDummy.DefaultView.Count - 1
                        rowDummy = TBLDummy.DefaultView(StartIndex).Row
                        TotalValue += rowDummy.Nett_Sales

                        If TotalValue = RefValue Then
                            'pas, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'baris product selanjutnya pindah ke group selanjutnya
                            StartIndex += 1
                            Exit While
                        ElseIf TotalValue > RefValue Then
                            If Me.chkStopBeforeRefValue.Checked = True Then
                                'baris ini masuk ke dalam group selanjutnya, tidak menambahkan start index, langsung keluar while
                                'kembali kurangkan total value, akan ditambahkan ketika mengevaluasi baris ini lagi
                                TotalValue -= rowDummy.Nett_Sales
                                Exit While
                            Else
                                'baris ini masuk, salin data
                                Me.SalinData(rowDummy, rowDGV)
                                'baris product selanjutnya pindah ke group selanjutnya
                                StartIndex += 1
                                Exit While
                            End If
                        Else
                            'baris ini masuk group ybs, salin data
                            Me.SalinData(rowDummy, rowDGV)
                            'cek baris selanjutnya
                            StartIndex += 1
                        End If
                    End While
                End If
            Next

            'klasifikasikan baris yang tidak masuk ke dalam baris terakhir
            If StartIndex <> 0 Then
                While StartIndex <= TBLDummy.DefaultView.Count - 1
                    rowDummy = TBLDummy.DefaultView(StartIndex).Row
                    rowDGV = Me.dgvItemGroup.Rows(Me.dgvItemGroup.Rows.Count - 1)
                    'baris ini masuk group ybs, salin data
                    Me.SalinData(rowDummy, rowDGV)
                    'cek baris selanjutnya
                    StartIndex += 1
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinData(ByVal vRowProduct As DseTransaksiMaterialPlanningGroup.tblmaster_productRow, ByVal rowDGV As DataGridViewRow)
        Try
            vRowProduct.Identifier_Group = rowDGV.Cells(Me.ColGroup_Identifier.Index).Value
            vRowProduct.Procurement_strategy = Me.objItem.GetProcurementStrategy(rowDGV.Cells(Me.ColGroup_ProcurementStrategy.Index).Value)
            vRowProduct.Optimum_Stock_Strategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_OptimumStockTrategy.Index).Value)
            vRowProduct.Optimum_Stock_Days = rowDGV.Cells(Me.ColGroup_OptimumStockDays.Index).Value
            vRowProduct.Optimum_Stock_Qty = rowDGV.Cells(Me.ColGroup_OptimumStockQty.Index).Value
            vRowProduct.Safety_Stock_Strategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_SafetyStockStrategy.Index).Value)
            vRowProduct.Safety_Stock_Days = rowDGV.Cells(Me.ColGroup_SafetyStockDays.Index).Value
            vRowProduct.Safety_Stock_Qty = rowDGV.Cells(Me.ColGroup_SafetyStockQty.Index).Value
            vRowProduct.Lead_Time_Strategy = Me.objItem.GetLeadTimeStrategy(rowDGV.Cells(Me.ColGroup_LeadTimeStrategy.Index).Value)
            vRowProduct.Lead_Time_Days = rowDGV.Cells(Me.ColGroup_LeadTimeDays.Index).Value
            vRowProduct.Lead_Time_Qty = rowDGV.Cells(Me.ColGroup_LeadTimeQty.Index).Value
            vRowProduct.Urutan_Group = rowDGV.Cells(Me.ColGroup_No.Index).Value
            vRowProduct.Group_Code = rowDGV.Cells(Me.ColGroup_Code.Index).Value
            vRowProduct.Group_Desc = rowDGV.Cells(Me.ColGroup_Desc.Index).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanReturnCost_AmountNonIDR(ByVal vTBLSR_Raw As DseTransaksiMaterialPlanningGroup.tbltrans_sales_returnDataTable, ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim rowSR_Raw As DseTransaksiMaterialPlanningGroup.tbltrans_sales_returnRow
        Dim DTASR As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_sales_return_itemcostTableAdapter
        Dim TBLSR As New DseTransaksiMaterialPlanningGroup.tbltrans_sales_return_itemcostDataTable
        Dim rowSR As DseTransaksiMaterialPlanningGroup.tbltrans_sales_return_itemcostRow
        Dim Cost As Decimal = 0
        Dim DV As Data.DataRowView
        Dim DTASR1 As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_sales_return2TableAdapter
        Dim TBLSR1 As New DseTransaksiMaterialPlanningGroup.tbltrans_sales_return2DataTable
        Dim rowSR1 As DseTransaksiMaterialPlanningGroup.tbltrans_sales_return2Row
        Dim SubTotal As Decimal = 0
        Dim DefPrec As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim StartDateDummy As Date = vStartDate

        Try
            DTASR.Connection.ConnectionString = Core_Module.GetConstring
            DTASR1.Connection.ConnectionString = DTASR.Connection.ConnectionString

            While StartDateDummy <= vEndDate
                DTASR.Fill(TBLSR, Me.ObjClosing.GetPrimaryClosing(StartDateDummy), _
                ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR)
                For Each rowSR_Raw In vTBLSR_Raw.Rows
                    Cost = 0
                    TBLSR.DefaultView.RowFilter = "[Primary Stock In Sales Return] = " + rowSR_Raw.Primary_Stock_In_Sales_Return.ToString + " AND " + _
                    "[Primary Product] = " + rowSR_Raw.Primary_Product.ToString
                    If TBLSR.DefaultView.Count > 0 Then
                        For Each DV In TBLSR.DefaultView
                            rowSR = DV.Row
                            Cost += rowSR.Cost
                        Next
                        rowSR_Raw.Cost = Cost
                    End If
                Next
                StartDateDummy = ObjCD.GetFirstDateofTheMonth(StartDateDummy + TimeSpan.FromDays(32))
            End While

            'isikan data return amount yang currencynya bukan IDR
            DTASR1.Fill(TBLSR1, Me.ObjCD.GetLasDateOfTheMonth(vEndDate), vStartDate, Core_Module.PrimaryBaseCurrency, CInt(TipeSR))
            For Each rowSR1 In TBLSR1.Rows
                rowSR1.Sub_Total_Base = Math.Round(rowSR1.Sub_Total * rowSR1.Tax_Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)

                vTBLSR_Raw.DefaultView.RowFilter = "[Primary SR] = " + rowSR1.Primary.ToString
                SubTotal = 0
                If vTBLSR_Raw.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To vTBLSR_Raw.DefaultView.Count - 1
                        rowSR_Raw = vTBLSR_Raw.DefaultView(ij).Row
                        If ij = vTBLSR_Raw.DefaultView.Count - 1 Then
                            rowSR_Raw.Amount = rowSR1.Sub_Total_Base - SubTotal
                        Else
                            'bukan baris terakhir, hitung biasa, catatkan sub total
                            If rowSR_Raw.Primary_Currency <> Core_Module.PrimaryBaseCurrency Then
                                rowSR_Raw.Amount = Math.Round(rowSR_Raw.Amount * rowSR_Raw.Exchange_Rate, DefPrec, _
                                    MidpointRounding.AwayFromZero)
                                SubTotal += rowSR_Raw.Amount
                            End If
                        End If
                    Next
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR.Dispose()
            TBLSR.Dispose()
            DTASR1.Dispose()
            TBLSR1.Dispose()
        End Try
    End Sub

    Private Sub KurangkanItemReturn(ByVal TBL1 As DseTransaksiMaterialPlanningGroup.tbltrans_salesinvoice_item_view0DataTable, _
    ByVal TBLReturn As DseTransaksiMaterialPlanningGroup.tbltrans_salesinvoice_item_returnDataTable)
        Dim Filter As String = ""
        Dim row1 As DseTransaksiMaterialPlanningGroup.tbltrans_salesinvoice_item_view0Row
        Dim i As Integer = 0
        Dim rowReturn As DseTransaksiMaterialPlanningGroup.tbltrans_salesinvoice_item_returnRow

        Try
            For Each rowReturn In TBLReturn.Rows
                Filter = "[Primary Invoice] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                TBL1.DefaultView.RowFilter = Filter

                i = 0
                While TBL1.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = TBL1.DefaultView(i).Row
                    If row1.Total_Qty >= rowReturn.Total_Qty Then
                        row1.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    End If

                    If row1.Total_Qty = 0 Then
                        TBL1.Rows.Remove(row1)
                    End If
                End While
            Next

            TBL1.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmTrans_MaterialPlanningDashboard_Data1_Print

        Try
            If Me.txtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.txtPrimary.Text <> "-1" Then
                ij.Obj = Me.obj
                ij.vPrimary = Me.txtPrimary.Text
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub dgvItemInvoiceException_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemInvoiceException.CellBeginEdit
        'later, bareng modifikasi material planning pada master product, harus ada padanannya di master product
    End Sub

    Private Sub dgvItemInvoiceException_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemInvoiceException.CellEndEdit
        'later, bareng modifikasi material planning pada master product, harus ada padanannya di master product
        'KETIKA MEMASUKKAN INVOICE BARU, USER DITANYAKAN PRODUCT CODE APA SAJA YANG PERLU DIMASUKKAN DENGAN MENGGUNAKAN FORM BARU, BISA MEMASUKAN BEBERAPA PRODUCT CODE SEKALIGUS
    End Sub

    Private Sub dgvItemInvoiceException_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemInvoiceException.RowsAdded
        'later, bareng modifikasi material planning pada master product, harus ada padanannya di master product
    End Sub

    Private Sub dgvItemInvoiceException_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemInvoiceException.UserDeletedRow
        'later, bareng modifikasi material planning pada master product, harus ada padanannya di master product
    End Sub

    Private Sub llbFindProduct_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbFindProduct.LinkClicked
        Dim i As New Core_Module.frmCariMasterProduct
        Dim PrimaryProduct As Double = -1
        Dim rowdgv As DataGridViewRow

        Try
            If Me.dgvItemProduct.Rows.Count = 0 Then
                Exit Sub
            End If
            'cari product yang akan diset posisinya
            i.cboCriteria.SelectedIndex = 0
            i.BackColorAll = Me.BackColorAll
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                PrimaryProduct = i.PrimaryProduct
                'cari baris yang berikan product tsb
                For Each rowdgv In Me.dgvItemProduct.Rows
                    If rowdgv.Cells(Me.ColCodeNo_PrimaryCodeNo.Index).Value = PrimaryProduct Then
                        'fokus ke baris yang bersangkutan
                        Me.dgvItemProduct.FirstDisplayedScrollingRowIndex = rowdgv.Index
                        Me.dgvItemProduct.ClearSelection()
                        rowdgv.Selected = True
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub chkShowDetail_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDetail.CheckedChanged
        Try
            If Me.chkShowDetail.Checked = True Then
                Me.ColCodeNo_Qty.Visible = True
                Me.ColCodeNo_QtyPerc.Visible = True
                Me.ColCodeNo_AverageQty.Visible = True
                Me.ColCodeNo_AverageNettSales.Visible = True
                Me.ColCodeNo_NettSales.Visible = True
                Me.ColCodeNo_NettSalesPerc.Visible = True
                Me.ColCodeNo_GrossMargin.Visible = True
                Me.ColCodeNo_GrossMarginPerc.Visible = True
                Me.ColCodeNo_AverageGrossMargin.Visible = True

                Me.ColCodeNo_SafetyStockStrategy.Visible = True
                Me.ColCodeNo_ProcurementStrategy.Visible = True
                Me.ColCodeNo_OptimumStockStrategy.Visible = True
                Me.ColCodeNo_LeadTimeStrategy.Visible = True
            Else
                Me.ColCodeNo_Qty.Visible = False
                Me.ColCodeNo_QtyPerc.Visible = False
                Me.ColCodeNo_AverageQty.Visible = False
                Me.ColCodeNo_AverageNettSales.Visible = False
                Me.ColCodeNo_NettSales.Visible = False
                Me.ColCodeNo_NettSalesPerc.Visible = False
                Me.ColCodeNo_GrossMargin.Visible = False
                Me.ColCodeNo_GrossMarginPerc.Visible = False
                Me.ColCodeNo_AverageGrossMargin.Visible = False

                Me.ColCodeNo_SafetyStockStrategy.Visible = False
                Me.ColCodeNo_ProcurementStrategy.Visible = False
                Me.ColCodeNo_OptimumStockStrategy.Visible = False
                Me.ColCodeNo_LeadTimeStrategy.Visible = False
            End If

            Me.SetKolomGroup_SafetyStock_ItemProduct()
            Me.SetKolomGroup_OptimumStock_ItemProduct()
            Me.SetKolomGroup_LeadTime_ItemProduct()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProduct_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProduct.CellContentClick
        Dim ij As New FrmTrans_MaterialPlanningDashboard_Data1_Explain

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColCodeNo_NettSales.Index Or _
                e.ColumnIndex = Me.ColCodeNo_Qty.Index Or _
                e.ColumnIndex = Me.ColCodeNo_GrossMargin.Index Then
                    Me.TBLItem_Adjustment.DefaultView.RowFilter = "[primary code no] = " + Me.dgvItemProduct.Item(Me.ColCodeNo_PrimaryCodeNo.Index, e.RowIndex).Value.ToString
                    Me.TBLItem_Bundling.DefaultView.RowFilter = "[primary code no] = " + Me.dgvItemProduct.Item(Me.ColCodeNo_PrimaryCodeNo.Index, e.RowIndex).Value.ToString
                    Me.TBLItem_Invoice.DefaultView.RowFilter = "[primary code no] = " + Me.dgvItemProduct.Item(Me.ColCodeNo_PrimaryCodeNo.Index, e.RowIndex).Value.ToString
                    Me.TBLItem_SR.DefaultView.RowFilter = "[primary code no] = " + Me.dgvItemProduct.Item(Me.ColCodeNo_PrimaryCodeNo.Index, e.RowIndex).Value.ToString

                    If Me.TBLItem_Adjustment.DefaultView.Count > 0 Or _
                    Me.TBLItem_Bundling.DefaultView.Count > 0 Or _
                    Me.TBLItem_Invoice.DefaultView.Count > 0 Or _
                    Me.TBLItem_SR.DefaultView.Count > 0 Then
                        'isikan data adjustment
                        ij.TBLItem_Adjustment = Me.TBLItem_Adjustment
                        ij.TBLItem_Bundling = Me.TBLItem_Bundling
                        ij.TBLItem_Invoice = Me.TBLItem_Invoice
                        ij.TBLItem_SR = Me.TBLItem_SR
                        ij.CodeNo = Me.dgvItemProduct.Item(Me.ColCodeNo_CodeNo.Index, e.RowIndex).Value
                        ij.Desc = Me.dgvItemProduct.Item(Me.ColCodeNo_Desc.Index, e.RowIndex).Value
                        ij.Unit = Me.dgvItemProduct.Item(Me.ColCodeNo_Unit.Index, e.RowIndex).Value
                        ij.vPrimaryProduct = Me.dgvItemProduct.Item(Me.ColCodeNo_PrimaryCodeNo.Index, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll

                        ij.ShowDialog()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub dgvItemGroup_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemGroup.UserDeletingRow
        Try
            If Me.dgvItemProduct.RowCount <> 0 Then
                If MsgBox("Deleting this row will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                Else
                    Me.dgvItemProduct.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemProductGroup.UserDeletingRow
        Try
            If Me.dgvItemProduct.RowCount <> 0 Then
                If MsgBox("Deleting this row will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                Else
                    Me.dgvItemProduct.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSupplier_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemSupplier.UserDeletingRow
        Try
            If Me.dgvItemProduct.RowCount <> 0 Then
                If MsgBox("Deleting this row will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                Else
                    Me.dgvItemProduct.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.BufferFrom.Date <> Me.DtpFrom.Value.Date Then
                    If Me.dgvItemProduct.RowCount <> 0 Then
                        If MsgBox("Changing period will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                            Me.DtpFrom.Value = Me.BufferFrom
                        Else
                            Me.dgvItemProduct.Rows.Clear()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFrom_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpFrom.Enter
        Try
            Me.BufferFrom = Me.DtpFrom.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFrom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpFrom.Leave
        Try
            Me.BackgroundWorker5.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.Enter
        Try
            Me.BufferTo = Me.dtpTo.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.Leave
        Try
            Me.BackgroundWorker6.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllIncludeAdjustment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeAdjustment.Enter
        Try
            Me.BufferBoolean_IncudeAdjustment = Me.chkAllIncludeAdjustment.Checked
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllIncludeAdjustment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeAdjustment.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItemProduct.Rows.Count > 0 Then
                    If Me.chkAllIncludeAdjustment.Checked <> Me.BufferBoolean_IncudeAdjustment Then
                        If MsgBox("Changing data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.chkAllIncludeAdjustment.Checked = Me.BufferBoolean_IncudeAdjustment
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllIncludeBundling_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeBundling.Enter
        Try
            Me.BufferBoolean_IncudeBundling = Me.chkAllIncludeBundling.Checked
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllIncludeBundling_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeBundling.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItemProduct.Rows.Count > 0 Then
                    If Me.chkAllIncludeBundling.Checked <> Me.BufferBoolean_IncudeBundling Then
                        If MsgBox("Changing data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.chkAllIncludeBundling.Checked = Me.BufferBoolean_IncudeBundling
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllIncludeSales_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeSales.Enter
        Try
            Me.BufferBoolean_IncudeSales = Me.chkAllIncludeSales.Checked
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAllIncludeSales_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllIncludeSales.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItemProduct.Rows.Count > 0 Then
                    If Me.chkAllIncludeSales.Checked <> Me.BufferBoolean_IncudeSales Then
                        If MsgBox("Changing data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.chkAllIncludeSales.Checked = Me.BufferBoolean_IncudeSales
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkStopBeforeRefValue_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkStopBeforeRefValue.Enter
        Try
            Me.BufferBoolean_IncudeStopBeforeRefValue = Me.chkStopBeforeRefValue.Checked
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkStopBeforeRefValue_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkStopBeforeRefValue.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItemProduct.Rows.Count > 0 Then
                    If Me.chkStopBeforeRefValue.Checked <> Me.BufferBoolean_IncudeStopBeforeRefValue Then
                        If MsgBox("Changing data will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Me.dgvItemProduct.Rows.Clear()
                        Else
                            Me.chkStopBeforeRefValue.Checked = Me.BufferBoolean_IncudeStopBeforeRefValue
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.BufferTo.Date <> Me.dtpTo.Value.Date Then
                    If Me.dgvItemProduct.RowCount <> 0 Then
                        If MsgBox("Changing period will reset calculation, continue?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                            Me.dtpTo.Value = Me.BufferTo
                        Else
                            Me.dgvItemProduct.Rows.Clear()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProduct_CellBeginEdit1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProduct.CellBeginEdit
        Dim rowDGV As DataGridViewRow
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy

        Try
            rowDGV = Me.dgvItemProduct.Rows(e.RowIndex)
            'yang bisa diedit
            If rowDGV.Cells(Me.ColCodeNo_Override.Index).Value = Boolean.TrueString Then
                '1. jika sudah di override
                If e.ColumnIndex = Me.ColCodeNo_ProcurementStrategy.Index Or _
                e.ColumnIndex = Me.ColCodeNo_Remarks.Index Or _
                e.ColumnIndex = Me.ColCodeNo_OptimumStockStrategy.Index Or _
                e.ColumnIndex = Me.ColCodeNo_LeadTimeStrategy.Index Or _
                e.ColumnIndex = Me.ColCodeNo_SafetyStockStrategy.Index Or _
                e.ColumnIndex = Me.ColCodeNo_Override.Index Then
                    Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                ElseIf e.ColumnIndex = Me.ColCodeNo_LeadTimeDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_LeadTimeQty.Index Then
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_LeadTimeStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        'yang bisa diedit: qty
                        If e.ColumnIndex = Me.ColCodeNo_LeadTimeQty.Index Then
                            Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        'yang bisa diedit: days
                        If e.ColumnIndex = Me.ColCodeNo_LeadTimeDays.Index Then
                            Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColCodeNo_SafetyStockStrategyDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_SafetyStockQty.Index Then
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        'yang bisa diedit: qty
                        If e.ColumnIndex = Me.ColCodeNo_SafetyStockQty.Index Then
                            Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        'yang bisa diedit: days
                        If e.ColumnIndex = Me.ColCodeNo_SafetyStockStrategyDays.Index Then
                            Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColCodeNo_OptimumStockDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_OptimumStockQty.Index Then
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                        'yang bisa diedit: qty
                        If e.ColumnIndex = Me.ColCodeNo_OptimumStockQty.Index Then
                            Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    Else
                        'yang bisa diedit: days
                        If e.ColumnIndex = Me.ColCodeNo_OptimumStockDays.Index Then
                            Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                '2. jika belum di override: coloverride + remarks
                If e.ColumnIndex = Me.ColCodeNo_Remarks.Index Or _
                e.ColumnIndex = Me.ColCodeNo_Override.Index Then
                    Me.buffer = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexGroup(ByVal vGroupIdentifier As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemGroup.Rows
                If rowDGV.Cells(Me.ColGroup_Identifier.Index).Value = vGroupIdentifier Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvItemProduct_CellEndEdit1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProduct.CellEndEdit
        Dim DummyIndex As Integer = -1
        Dim TipeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy
        Dim TipeTrans As ClsTransaksiMaterialPlanningGroup.EnumTipe = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage
        Dim TipeProcurement As Core_Module.clsMasterProduct.EnumProcurementStrategy = clsMasterProduct.EnumProcurementStrategy.Build_To_Order
        Dim rowDGV, rowDGVGroup As DataGridViewRow

        Try
            If IsNothing(Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColCodeNo_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColCodeNo_OptimumStockDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_OptimumStockQty.Index Or _
                e.ColumnIndex = Me.ColCodeNo_LeadTimeDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_LeadTimeQty.Index Or _
                e.ColumnIndex = Me.ColCodeNo_SafetyStockStrategyDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_SafetyStockQty.Index Then
                    'desc, ref.value
                    'optimum stock days, optimum stock qty, lead time stock days, lead time stock qty
                    'safety stock days, safety stock qty
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColCodeNo_Remarks.Index Then
                    'desc, remarks
                    'pastikan ditrim dan max 250 karakter
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                ElseIf e.ColumnIndex = Me.ColCodeNo_Override.Index Then
                    'jika jadinya tidak override, maka kembalikan datanya ke semula
                    If CBool(Me.dgvItemProduct.Item(Me.ColCodeNo_Override.Index, e.RowIndex).Value) = CBool(Boolean.FalseString) Then
                        rowDGVGroup = Me.dgvItemGroup.Rows(Me.GetIndexGroup(Me.dgvItemProduct.Item(Me.ColCodeNo_IdentifierGroup.Index, e.RowIndex).Value))
                        rowDGV = Me.dgvItemProduct.Rows(e.RowIndex)
                        With rowDGV
                            .Cells(Me.ColCodeNo_LeadTimeDays.Index).Value = rowDGVGroup.Cells(Me.ColGroup_LeadTimeDays.Index).Value
                            .Cells(Me.ColCodeNo_LeadTimeQty.Index).Value = rowDGVGroup.Cells(Me.ColGroup_LeadTimeQty.Index).Value
                            .Cells(Me.ColCodeNo_LeadTimeStrategy.Index).Value = rowDGVGroup.Cells(Me.ColGroup_LeadTimeStrategy.Index).Value
                            .Cells(Me.ColCodeNo_OptimumStockDays.Index).Value = rowDGVGroup.Cells(Me.ColGroup_OptimumStockDays.Index).Value
                            .Cells(Me.ColCodeNo_OptimumStockQty.Index).Value = rowDGVGroup.Cells(Me.ColGroup_OptimumStockQty.Index).Value
                            .Cells(Me.ColCodeNo_OptimumStockStrategy.Index).Value = rowDGVGroup.Cells(Me.ColGroup_OptimumStockTrategy.Index).Value
                            .Cells(Me.ColCodeNo_ProcurementStrategy.Index).Value = rowDGVGroup.Cells(Me.ColGroup_ProcurementStrategy.Index).Value
                            .Cells(Me.ColCodeNo_SafetyStockQty.Index).Value = rowDGVGroup.Cells(Me.ColGroup_SafetyStockQty.Index).Value
                            .Cells(Me.ColCodeNo_SafetyStockStrategy.Index).Value = rowDGVGroup.Cells(Me.ColGroup_SafetyStockStrategy.Index).Value
                            .Cells(Me.ColCodeNo_SafetyStockStrategyDays.Index).Value = rowDGVGroup.Cells(Me.ColGroup_SafetyStockDays.Index).Value
                        End With
                        Me.SetKolomGroup_LeadTime_ItemProduct()
                        Me.SetKolomGroup_OptimumStock_ItemProduct()
                        Me.SetKolomGroup_SafetyStock_ItemProduct()
                    End If
                ElseIf e.ColumnIndex = Me.ColCodeNo_LeadTimeStrategy.Index Then
                    'jika based on usage: qty = 0
                    'jika user defined: days = 0
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_LeadTimeStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        Me.dgvItemProduct.Item(Me.ColCodeNo_LeadTimeQty.Index, e.RowIndex).Value = 0
                    Else
                        Me.dgvItemProduct.Item(Me.ColCodeNo_LeadTimeDays.Index, e.RowIndex).Value = 0
                    End If
                    SetKolomGroup_LeadTime_ItemProduct()
                ElseIf e.ColumnIndex = Me.ColCodeNo_OptimumStockStrategy.Index Then
                    'jika based on usage: qty = 0
                    'jika user defined: days = 0
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockQty.Index, e.RowIndex).Value = 0
                    Else
                        Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockDays.Index, e.RowIndex).Value = 0
                    End If
                    SetKolomGroup_OptimumStock_ItemProduct()
                ElseIf e.ColumnIndex = Me.ColCodeNo_SafetyStockStrategy.Index Then
                    'jika based on usage: qty = 0
                    'jika user defined: days = 0
                    TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategy.Index, e.RowIndex).Value)
                    If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                        Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockQty.Index, e.RowIndex).Value = 0
                    Else
                        Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategyDays.Index, e.RowIndex).Value = 0
                    End If
                    SetKolomGroup_SafetyStock_ItemProduct()
                ElseIf e.ColumnIndex = Me.ColCodeNo_LeadTimeDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_OptimumStockDays.Index Or _
                e.ColumnIndex = Me.ColCodeNo_SafetyStockStrategyDays.Index Then
                    'optimum stock days, lead time stock days, safety stock days
                    'pastikan angka
                    If IsNumeric(Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit grouping 0 decimal
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value, 0)
                    'pastikan tidak negatif
                    If CInt(Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                        MsgBox("Please specify non-negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'pastikan optimum stock > safety stock
                    If e.ColumnIndex = Me.ColCodeNo_OptimumStockDays.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            TipeProcurement = Me.objItem.GetProcurementStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_ProcurementStrategy.Index, e.RowIndex).Value)
                            If TipeProcurement = clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                                If CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockDays.Index, e.RowIndex).Value) <= _
                                CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategyDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days > safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            Else
                                If CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockDays.Index, e.RowIndex).Value) < _
                                CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategyDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days => safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColCodeNo_SafetyStockStrategyDays.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockStrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            TipeProcurement = Me.objItem.GetProcurementStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_ProcurementStrategy.Index, e.RowIndex).Value)
                            If TipeProcurement = clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                                If CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockDays.Index, e.RowIndex).Value) <= _
                                CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategyDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days > safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            Else
                                If CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockDays.Index, e.RowIndex).Value) < _
                                CInt(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategyDays.Index, e.RowIndex).Value) Then
                                    MsgBox("Please make sure optimum stock days => safety stock days.", MsgBoxStyle.Information, "Information")
                                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColCodeNo_LeadTimeQty.Index Or _
                e.ColumnIndex = Me.ColCodeNo_OptimumStockQty.Index Or _
                e.ColumnIndex = Me.ColCodeNo_SafetyStockQty.Index Then
                    'optimum stock qty, lead time stock qty, safety stock qty
                    'pastikan angka
                    If IsNumeric(Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit grouping berdasarkan decimal master product
                    Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value, _
                    Me.ObjProduct.GetAngkaDecimal(Me.dgvItemProduct.Item(Me.ColCodeNo_PrimaryCodeNo.Index, e.RowIndex).Value))
                    'pastikan tidak negatif
                    If CInt(Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                        MsgBox("Please specify non-negative number.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'pastikan optimum stock > safety stock
                    If e.ColumnIndex = Me.ColCodeNo_OptimumStockQty.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockStrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            If CDec(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockQty.Index, e.RowIndex).Value) <= _
                            CDec(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockQty.Index, e.RowIndex).Value) Then
                                MsgBox("Please make sure optimum stock qty > safety stock qty.", MsgBoxStyle.Information, "Information")
                                Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColCodeNo_SafetyStockQty.Index Then
                        TipeStrategy = Me.objItem.GetLeadTimeStrategy(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockStrategy.Index, e.RowIndex).Value)
                        If TipeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            If CDec(Me.dgvItemProduct.Item(Me.ColCodeNo_OptimumStockQty.Index, e.RowIndex).Value) <= _
                            CDec(Me.dgvItemProduct.Item(Me.ColCodeNo_SafetyStockQty.Index, e.RowIndex).Value) Then
                                MsgBox("Please make sure optimum stock qty > safety stock qty.", MsgBoxStyle.Information, "Information")
                                Me.dgvItemProduct.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker7_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker7.DoWork
        Me.UpdateData()
    End Sub

    Private Sub UpdateData()
        Dim dtaMasterProd As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_product2TableAdapter
        Dim tblMasterProd As New DseTransaksiMaterialPlanningGroup.tblmaster_product2DataTable
        Dim rowMasterProd As DseTransaksiMaterialPlanningGroup.tblmaster_product2Row

        Dim minStock, maxStock, DaysPerMonth As Double
        Dim startPeriodMin, finishPeriodMax, EndPeriod As Date
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
        Dim IncludeSalesInt, IncludeBundlingInt, IncludeAdjustmentInt As Integer

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DaysPerMonth = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan)
            Me.ObjCD.UpdateThreading()

            explainAvgMonthly = ""
            explainSafetyStock = ""
            explainLeadTime = ""
            explainOptimumStock = ""

            'Me.Prg1.Visible = True
            'Me.LblProgress.Visible = True

            If Me.chkAllIncludeAdjustment.Checked = True Then
                IncludeAdjustmentInt = 1
            Else
                IncludeAdjustmentInt = 0
            End If
            If Me.chkAllIncludeBundling.Checked = True Then
                IncludeBundlingInt = 1
            Else
                IncludeBundlingInt = 0
            End If
            If Me.chkAllIncludeSales.Checked = True Then
                IncludeSalesInt = 1
            Else
                IncludeSalesInt = 0
            End If

            dtaMasterProd.Connection.ConnectionString = Core_Module.GetConstring
            'dtaReplicaProd.Connection.ConnectionString = dtaMasterProd.Connection.ConnectionString
            If dtaMasterProd.Fill(tblMasterProd, Me.txtPrimary.Text) > 0 Then
                ''fill historical, sales, bundling, dan adj sekaligus u/ semua produk
                startPeriodMin = Me.ObjCD.GetFirstDateofTheMonth(Me.DtpFrom.Value)
                finishPeriodMax = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value)
                EndPeriod = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpTo.Value)
                Me.objPlanning.GetHistoricalByPrimary(tblHistorical, -1, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax))
                Me.objPlanning.GetListSales(LSales, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexSales)
                Me.objPlanning.GetListBundling(LBundling, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexBundling)
                Me.objPlanning.GetListAdj(LAdj, Me.ObjCD.GetFirstDateofTheMonth(startPeriodMin), Me.ObjCD.GetLasDateOfTheMonth(finishPeriodMax), tblIndexAdjustment)
                Me.Prg1.Visible = True
                Me.Prg1.Maximum = tblMasterProd.Rows.Count + 1
                Me.LblProgress.Visible = True

                For Each rowMasterProd In tblMasterProd
                    Me.Prg1.Value = tblMasterProd.Rows.IndexOf(rowMasterProd) + 1
                    Me.LblProgress.Text = "Updating data " + Me.Prg1.Value.ToString + " of " + Me.Prg1.Maximum.ToString + " : " _
                    & rowMasterProd.Product_Code + " - " + rowMasterProd.Description

                    With rowMasterProd
                        minStock = Me.ObjCD.DigitGrouping(.Minimum_Stock, ._Decimal)
                        'optStock = Me.ObjCD.DigitGrouping(.optimum_stock_qty, ._Decimal)
                        maxStock = Me.ObjCD.DigitGrouping(.maximum_stock, ._Decimal)
                        safetyStockTime = .safety_stock_days
                        leadTime = .lead_time_days
                        optimumStockTime = .optimum_stock_days

                        includeSales = Me.chkAllIncludeSales.Checked
                        includeBundling = Me.chkAllIncludeBundling.Checked
                        includeAdj = Me.chkAllIncludeAdjustment.Checked

                        ''hitung avg, safety stock, dan lead stock serta explanation nya
                        Me.objPlanning.CalculateAverageMonthlyDemandAll(.Primary, startPeriodMin, EndPeriod, DaysPerMonth, includeSales, includeBundling, includeAdj, _
                        avgMonthly, explainAvgMonthly, tblHistorical, LSales, LBundling, LAdj, tblIndexSales, tblIndexBundling, tblIndexAdjustment, ._Decimal)

                        If .safety_stock_strategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            Me.objPlanning.CalculateSafetyStock(avgMonthly, .safety_stock_days, DaysPerMonth, safetyStockQty, explainSafetyStock, ._Decimal)
                        ElseIf .safety_stock_strategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            safetyStockQty = .safety_stock_qty
                            Me.objPlanning.CalculateSafetyStockTime(avgMonthly, .safety_stock_qty, DaysPerMonth, safetyStockTime, explainSafetyStock, ._Decimal)
                        End If
                        If .lead_time_strategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            Me.objPlanning.CalculateLeadTimeQty(avgMonthly, .lead_time_days, DaysPerMonth, leadStockQty, explainLeadTime, ._Decimal)
                        ElseIf .lead_time_strategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            leadStockQty = .lead_time_qty
                            Me.objPlanning.CalculatedLeadTimeDays(avgMonthly, .lead_time_qty, DaysPerMonth, leadTime, explainLeadTime, ._Decimal)
                        End If
                        If .optimum_stock_strategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                            Me.objPlanning.CalculateOptimumStockQty(avgMonthly, .optimum_stock_days, DaysPerMonth, optimumStockQty, explainOptimumStock, ._Decimal)
                        ElseIf .optimum_stock_strategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined Then
                            optimumStockQty = .optimum_stock_qty
                            Me.objPlanning.CalculatedOptimumStockDays(avgMonthly, .optimum_stock_qty, DaysPerMonth, optimumStockTime, explainOptimumStock, ._Decimal)
                        End If

                        ''betulkan qty2 yang kurang dari safety stock
                        If minStock < safetyStockQty Then
                            minStock = safetyStockQty
                        End If
                        If optimumStockQty < safetyStockQty Then
                            optimumStockQty = safetyStockQty
                            Me.objPlanning.CalculatedOptimumStockDays(avgMonthly, optimumStockQty, DaysPerMonth, optimumStockTime, explainOptimumStock, ._Decimal)
                        End If
                        If maxStock < safetyStockQty Then
                            maxStock = safetyStockQty
                        End If

                        ''simpan perubahan ke db
                        dtaMasterProd.UpdateQuery(minStock, optimumStockQty, maxStock, avgMonthly, safetyStockTime, safetyStockQty, _
                        leadTime, leadStockQty, explainAvgMonthly, explainSafetyStock, explainLeadTime, Me.UserCode, Now, Now, optimumStockTime, _
                        explainOptimumStock, Me.DtpFrom.Value, Me.dtpTo.Value, IncludeSalesInt, IncludeBundlingInt, IncludeAdjustmentInt, .safety_stock_strategy, _
                        .lead_time_strategy, .optimum_stock_strategy, .procurement_strategy, .Primary)
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Prg1.Visible = False
            'Me.Cursor = Cursors.Default
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub BackgroundWorker7_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker7.RunWorkerCompleted
        Try
            Me.AktifkanForm()
            Me.Cursor = Cursors.Arrow
            MsgBox("Data update completed.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        Try
            If Me.txtDocumentNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            'ketika apply ke product, pastikan 
            '1. optimum stock > safety stock
            '2. jika user defined walaupun decimal yang dispesifikasikan lebih besar dari decimal product, hasilnya tetap harus decimal product
            If Me.txtDocumentNo.Text = "" Then
                Exit Sub
            End If

            If MsgBox("Apply calculation?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan) = "" Then
                    MsgBox("Please specify work days per month.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Control.CheckForIllegalCrossThreadCalls = False
                If Me.BackgroundWorker7.IsBusy = False Then
                    Me.NonAktifkanForm()
                    Me.BackgroundWorker7.RunWorkerAsync()
                End If
                'Me.NonAktifkanForm()
                'Me.UpdateData()
                'Me.AktifkanForm()
                'Me.Cursor = Cursors.Arrow
                'MsgBox("Data update completed.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanForm()
        Try
            Me.Prg1.Visible = True
            Me.LblProgress.Visible = True
            Me.LblProgress.Text = ""

            Me.ControlBox = False
            'ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.btnApply.Visible = False
            Me.btnDelete.Visible = False
            Me.btnExit.Visible = False
            Me.btnFind.Visible = False
            Me.btnNew.Visible = False
            Me.btnPrint.Visible = False
            Me.btnEdit.Visible = False
            Me.LblProgress.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanForm()
        Try
            Me.Prg1.Visible = False
            Me.LblProgress.Visible = False

            Me.ControlBox = True
            Me.btnApply.Visible = True
            Me.btnDelete.Visible = True
            Me.btnExit.Visible = True
            Me.btnFind.Visible = True
            Me.btnNew.Visible = True
            Me.btnPrint.Visible = True
            Me.btnEdit.Visible = True
            Me.LblProgress.Visible = False
            'ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowDetail_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDetail.Leave
        Try
            If Me.chkShowDetail.Enabled = True Then
                'simpan datanya ke dalam software untuk diingat settingnya
                If Me.chkShowDetail.Checked = True Then
                    Me.ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanningGrouping, _
                    ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanningGrouping_ShowDetail, "1", Me.UserCode)
                Else
                    Me.ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanningGrouping, _
                    ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanningGrouping_ShowDetail, "0", Me.UserCode)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class