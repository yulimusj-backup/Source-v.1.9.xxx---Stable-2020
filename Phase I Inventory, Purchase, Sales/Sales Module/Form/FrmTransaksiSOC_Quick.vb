Option Explicit On
Option Strict Off

Public Class FrmTransaksiSOC_Quick
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public UserCode As String = ""

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim ObjSalesPerson As New Core_Module.ClsMasterSalesPerson
    Dim ObjSO As New Sales_Module.ClsTransaksiSalesOrder
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim ObjCustoemr As New Core_Module.clsMasterCustomer
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim EraseRowIndex As Integer = -1
    Dim EraseIndex_C As Integer = -1
    Dim EraseIndex_SG As Integer = -1
    Dim EraseIndex_SA As Integer = -1
    Dim EraseIndex_MS As Integer = -1
    Dim EraseIndex_CT As Integer = -1
    Dim EraseIndex_SP As Integer = -1
    Dim EraseIndex_ST As Integer = -1
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1 '
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim DummyIndex As Integer = -1
    Dim index As Integer = -1
    Dim buffer As String = ""

    Dim DTAUser As New DseTransaksiSOC_QuickTableAdapters.tbltrans_soc_setting_userTableAdapter
    Dim TBLUser As New DseTransaksiSOC_Quick.tbltrans_soc_setting_userDataTable
    Dim rowUser As DseTransaksiSOC_Quick.tbltrans_soc_setting_userRow
    Dim DTA As New DseTransaksiSOC_QuickTableAdapters.DataTable1TableAdapter
    Dim TBL As New DseTransaksiSOC_Quick.DataTable1DataTable
    Dim Row As DseTransaksiSOC_Quick.DataTable1Row
    Dim DTAUnit As New Core_Module.dseMasterProductTableAdapters.tblmaster_product_item_multiunitTableAdapter
    Dim TBLUnit As New Core_Module.dseMasterProduct.tblmaster_product_item_multiunitDataTable
    Dim rowUnit As Core_Module.dseMasterProduct.tblmaster_product_item_multiunitRow

    Private Sub FrmTransaksiSOC_Quick_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim rowDGV As DataGridViewRow
        Try
            'simpan posisi terakhir format supplier
            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SOC_Quick, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SOC_Quick_CustomerFormat, Me.cboCustomerFormat.SelectedIndex, Me.UserCode)
            'simpan data ref user
            Me.DgvUser.AllowUserToAddRows = False
            If Me.DgvUser.Rows.Count > 0 Then
                Me.DTAUser.Delete(Me.UserCode)
                For Each rowDGV In Me.DgvUser.Rows
                    If IsNothing(rowDGV.Cells(Me.ColCreatedBy.Index).Value) = False Then
                        Me.DTAUser.Insert(Me.UserCode, rowDGV.Cells(Me.ColCreatedBy.Index).Value)
                    End If
                Next
            End If
            Me.DgvUser.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiSOC_Quick_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DummyVal As String = ""
        Dim rowDGV As DataGridViewRow

        Try
            Me.ShowIcon = False

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvSalesPerson.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvSalesArea.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCustomer.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCustomerGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvMarketSegmentation.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvCustomerType.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvShipto.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.cboCustomerFormat.SelectedIndex = 0
            DummyVal = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.POC_Quick, _
                      Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.POC_Quick_SupplierFormat, Me.UserCode)
            If DummyVal <> "" Then
                Me.cboCustomerFormat.SelectedIndex = DummyVal
            End If

            'load data user
            Me.DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAUser.FillBy_UserRef(Me.TBLUser, Me.UserCode)
            If Me.TBLUser.Rows.Count > 0 Then
                Me.DgvUser.AllowUserToAddRows = False
                For Each rowUser In Me.TBLUser.Rows
                    Me.DgvUser.Rows.Add()
                    rowDGV = Me.DgvUser.Rows(Me.DgvUser.Rows.Count - 1)
                    rowDGV.Cells(Me.ColCreatedBy.Index).Value = rowUser.user_code_created_by
                    rowDGV.Cells(Me.ColCreatedBy_Name.Index).Value = rowUser.user_code_ref
                    rowDGV.Cells(Me.ColCreatedBy_Primary.Index).Value = Me.ObjUser.GetPrimary(rowUser.user_code_ref)
                Next
                Me.DgvUser.AllowUserToAddRows = True
            End If
            Me.DgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dtpFrom.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


#Region "Filter datagridview"
    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
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
        Dim Level As Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

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
                    If Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
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
                    If Trim(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.objCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.objCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objCategory.GetDesc(ij.PrimaryCategory)
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
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.objCategory.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.objProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.objProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.objProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(ij.PrimaryProduct)
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
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.objProduct.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.objProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError

    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemProductGroup.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemProductGroup.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_CellContextMenuStripChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellContextMenuStripChanged

    End Sub

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim KodeUser As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            KodeUser = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(KodeUser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.bgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.bgwUser.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.bgwUser.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.bgwUser.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vKodeUser Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BgwUser_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwUser.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            ' If Me.TxtNoPO.Enabled = True Then
            rowDGV = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)
            With rowDGV
                .Cells(Me.ColCreatedBy_No.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColCreatedBy.Index).Value = ""
                .Cells(Me.ColCreatedBy_Name.Index).Value = ""
                .Cells(Me.ColCreatedBy_Primary.Index).Value = "-1"
            End With
            ' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomer.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColC_Code.Index Then
                Me.buffer = Me.DgvCustomer(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomer.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomer.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjRef As New Core_Module.clsMasterCustomer
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColC_Code.Index Then
                        .Cells(Me.ColC_Code.Index).Value = .Cells(Me.ColC_Code.Index).Value.ToString.ToUpper
                        If ObjRef.ApakahSudahAda(.Cells(Me.ColC_Code.Index).Value) = True Then
                            primary = ObjRef.GetPrimary(.Cells(Me.ColC_Code.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_C(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColC_Primary.Index).Value = primary
                                .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(primary)
                                .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_C = e.RowIndex
                                        Me.bgwCustomer.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomer)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomer
                            ij.txtCari.Text = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomer <> -1 Then
                                index = Me.GetIndexItem_C(ij.PrimaryCustomer, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColC_Primary.Index).Value = ij.PrimaryCustomer
                                    .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                    .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_C = e.RowIndex
                                            Me.bgwCustomer.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustomer)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColC_Primary.Index).Value = ij.PrimaryCustomer
                                        .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                        .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomer = -1 Then
                                If Me.DgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_C = e.RowIndex
                                        Me.bgwCustomer.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomer)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomer)
                        Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value = Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColC_Code.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_C = e.RowIndex
                                Me.bgwCustomer.RunWorkerAsync()
                            Else
                                Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomer)
                            Exit Sub
                        Else
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColC_Code.Index Then
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_C(ByVal vPrimary As Double, _
      ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomer.Rows.Count - 1
                RowDGV = Me.DgvCustomer.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColC_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Sub DgvCustomer_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvCustomer.DataError

    End Sub

    Private Sub DgvCustomer_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvCustomer.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvCustomer.Rows(Me.DgvCustomer.Rows.Count - 1)
            dr.Cells(Me.ColC_No.Index).Value = Me.DgvCustomer.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomer.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomer.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvCustomer, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustomer_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustomer.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustomer_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustomer.RunWorkerCompleted
        Try
            Me.DgvCustomer.Rows.RemoveAt(EraseIndex_C)
            Me.ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCustomerGroup.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSG_SalesGroup.Index Then
                Me.buffer = Me.dgvCustomerGroup(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSalesGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCustomerGroup.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvCustomerGroup.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Core_Module.clsMasterCustomerGroup
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSG_SalesGroup.Index Then
                        .Cells(Me.ColSG_SalesGroup.Index).Value = .Cells(Me.ColSG_SalesGroup.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColSG_SalesGroup.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColSG_SalesGroup.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_CG(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColSG_Primary.Index).Value = primary
                                .Cells(Me.ColSG_SalesGroup.Index).Value = ObjSA.GetGroupCode(primary)
                                .Cells(Me.ColSG_Desc.Index).Value = ObjSA.GetGroupDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvCustomerGroup.Item(Me.ColSG_SalesGroup.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SG = e.RowIndex
                                        Me.bgwCustGroup.RunWorkerAsync()
                                    Else
                                        Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomerGroup
                            ij.txtCari.Text = Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerGroup <> -1 Then
                                index = Me.GetIndexItem_CG(ij.PrimaryCustomerGroup, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColSG_Primary.Index).Value = ij.PrimaryCustomerGroup
                                    .Cells(Me.ColSG_SalesGroup.Index).Value = ObjSA.GetGroupCode(ij.PrimaryCustomerGroup)
                                    .Cells(Me.ColSG_Desc.Index).Value = ObjSA.GetGroupDesc(ij.PrimaryCustomerGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_SG = e.RowIndex
                                            Me.bgwCustGroup.RunWorkerAsync()
                                        Else
                                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColSG_Primary.Index).Value = ij.PrimaryCustomerGroup
                                        .Cells(Me.ColSG_SalesGroup.Index).Value = ObjSA.GetGroupCode(ij.PrimaryCustomerGroup)
                                        .Cells(Me.ColSG_Desc.Index).Value = ObjSA.GetGroupDesc(ij.PrimaryCustomerGroup)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerGroup = -1 Then
                                If Me.dgvCustomerGroup.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SG = e.RowIndex
                                        Me.bgwCustGroup.RunWorkerAsync()
                                    Else
                                        Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                        Me.dgvCustomerGroup.Item(Me.ColSG_SalesGroup.Index, e.RowIndex).Value = Me.dgvCustomerGroup.Item(Me.ColSG_SalesGroup.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.dgvCustomerGroup.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColSG_SalesGroup.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_SG = e.RowIndex
                                Me.bgwCustGroup.RunWorkerAsync()
                            Else
                                Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
                            Exit Sub
                        Else
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSG_SalesGroup.Index Then
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.dgvCustomerGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_CG(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvCustomerGroup.Rows.Count - 1
                RowDGV = Me.dgvCustomerGroup.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSG_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub bgwCustGroup_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustGroup.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustGroup_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustGroup.RunWorkerCompleted
        Try
            Me.dgvCustomerGroup.Rows.RemoveAt(EraseIndex_SG)
            Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomerGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvCustomerGroup.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.dgvCustomerGroup.Rows(Me.dgvCustomerGroup.Rows.Count - 1)
            dr.Cells(Me.ColSG_No.Index).Value = Me.dgvCustomerGroup.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomerGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvCustomerGroup.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvCustomerGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCustomerGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCustomerGroup.Sorted
        Try
            ObjCD.SetNumber(Me.dgvCustomerGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSalesArea.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                Me.buffer = Me.DgvSalesArea(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSalesArea.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSalesArea.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Core_Module.clsMasterCustomerSalesArea
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                        .Cells(Me.ColSA_SalesArea.Index).Value = .Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColSA_SalesArea.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_SA(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColSA_Primary.Index).Value = primary
                                .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(primary)
                                .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.bgwSalesArea.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterSalesArea
                            ij.txtCari.Text = Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerSalesArea <> -1 Then
                                index = Me.GetIndexItem_SA(ij.PrimaryCustomerSalesArea, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                    .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                    .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_SA = e.RowIndex
                                            Me.bgwSalesArea.RunWorkerAsync()
                                        Else
                                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                        .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                        .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerSalesArea = -1 Then
                                If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.bgwSalesArea.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                        Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_SA = e.RowIndex
                                Me.bgwSalesArea.RunWorkerAsync()
                            Else
                                Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvSalesArea)
                            Exit Sub
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSA_SalesArea.Index Then
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_SA(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSalesArea.Rows.Count - 1
                RowDGV = Me.DgvSalesArea.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSA_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvSalesArea_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvSalesArea.DataError
        Try
            ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvSalesArea.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvSalesArea.Rows(Me.DgvSalesArea.Rows.Count - 1)
            dr.Cells(Me.ColSA_No.Index).Value = Me.DgvSalesArea.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSalesArea.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSalesArea.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvSalesArea, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwSalesArea_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwSalesArea.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwSalesArea_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwSalesArea.RunWorkerCompleted
        Try
            Me.DgvSalesArea.Rows.RemoveAt(EraseIndex_SA)
            Me.ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvMarketSegmentation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                Me.buffer = Me.DgvMarketSegmentation(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvMarketSegmentation.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvMarketSegmentation.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjMS As New Core_Module.clsMasterCustomerMarketSegmentation
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                        .Cells(Me.ColMS_MarketSeg.Index).Value = .Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper
                        If ObjMS.ApakahSudahAda(.Cells(Me.ColMS_MarketSeg.Index).Value) = True Then
                            primary = ObjMS.GetPrimary(.Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_MS(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColMS_Primary.Index).Value = primary
                                .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(primary)
                                .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.bgwCustSegmen.RunWorkerAsync()
                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomerSegmentation
                            ij.txtCari.Text = Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerMarketSegmentation <> -1 Then
                                index = Me.GetIndexItem_MS(ij.PrimaryCustomerMarketSegmentation, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                    .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                    .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_MS = e.RowIndex
                                            Me.bgwCustSegmen.RunWorkerAsync()
                                        Else
                                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                        .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                        .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerMarketSegmentation = -1 Then
                                If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.bgwCustSegmen.RunWorkerAsync()

                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                        Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_MS = e.RowIndex
                                Me.bgwCustSegmen.RunWorkerAsync()
                            Else
                                Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                            Exit Sub
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColMS_MarketSeg.Index Then
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_MS(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvMarketSegmentation.Rows.Count - 1
                RowDGV = Me.DgvMarketSegmentation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColMS_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvMarketSegmentation_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvMarketSegmentation.DataError

    End Sub

    Private Sub DgvMarketSegmentation_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvMarketSegmentation.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvMarketSegmentation.Rows(Me.DgvMarketSegmentation.Rows.Count - 1)
            dr.Cells(Me.ColMS_No.Index).Value = Me.DgvMarketSegmentation.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvMarketSegmentation.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustSegmen_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustSegmen.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustSegmen_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustSegmen.RunWorkerCompleted
        Try
            Me.DgvMarketSegmentation.Rows.RemoveAt(EraseIndex_MS)
            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomerType.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCT_CustType.Index Then
                Me.buffer = Me.DgvCustomerType(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomerType.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomerType.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjCT As New Core_Module.clsMasterCustomerType
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCT_CustType.Index Then
                        .Cells(Me.ColCT_CustType.Index).Value = .Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper
                        If ObjCT.ApakahSudahAda(.Cells(Me.ColCT_CustType.Index).Value) = True Then
                            primary = ObjCT.GetPrimary(.Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_CT(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColCT_Primary.Index).Value = primary
                                .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(primary)
                                .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.bgwCustType.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterCustomerType
                            ij.txtCari.Text = Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerType <> -1 Then
                                index = Me.GetIndexItem_CT(ij.PrimaryCustomerType, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                    .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                    .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_CT = e.RowIndex
                                            Me.bgwCustType.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                        .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                        .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerType = -1 Then
                                If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.bgwCustType.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                        Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColCT_CustType.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_CT = e.RowIndex
                                Me.bgwCustType.RunWorkerAsync()
                            Else
                                Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomerType)
                            Exit Sub
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColCT_CustType.Index Then
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_CT(ByVal vPrimary As Double, _
       ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomerType.Rows.Count - 1
                RowDGV = Me.DgvCustomerType.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCT_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvCustomerType_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvCustomerType.DataError

    End Sub

    Private Sub DgvCustomerType_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvCustomerType.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvCustomerType.Rows(Me.DgvCustomerType.Rows.Count - 1)
            dr.Cells(Me.ColCT_No.Index).Value = Me.DgvCustomerType.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomerType.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomerType.Sorted
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustType_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwCustType.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwCustType_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwCustType.RunWorkerCompleted
        Try
            Me.DgvCustomerType.Rows.RemoveAt(EraseIndex_CT)
            Me.ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesPerson_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSalesPerson.CellBeginEdit
        Try
            If e.ColumnIndex = Me.SalesPerson_ColCode.Index Then
                Me.buffer = Me.DgvSalesPerson(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesPerson_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSalesPerson.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSalesPerson.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSP As New Core_Module.ClsMasterSalesPerson
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.SalesPerson_ColCode.Index Then
                        .Cells(Me.SalesPerson_ColCode.Index).Value = .Cells(Me.SalesPerson_ColCode.Index).Value.ToString.ToUpper
                        If ObjSP.ApakahSudahAda(.Cells(Me.SalesPerson_ColCode.Index).Value) = True Then
                            primary = ObjSP.GetPrimary(.Cells(Me.SalesPerson_ColCode.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_SP(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColSA_Primary.Index).Value = primary
                                .Cells(Me.SalesPerson_ColCode.Index).Value = ObjSP.GetCode(primary)
                                .Cells(Me.ColSA_Desc.Index).Value = ObjSP.GetName(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSalesPerson.Item(Me.SalesPerson_ColCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SP = e.RowIndex
                                        Me.bgwSalesPerson.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesPerson)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmMasterSalesPerson_Cari
                            ij.txtCari.Text = Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimarySalesPerson <> -1 Then
                                index = Me.GetIndexItem_SP(ij.PrimarySalesPerson, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColSA_Primary.Index).Value = ij.PrimarySalesPerson
                                    .Cells(Me.SalesPerson_ColCode.Index).Value = ObjSP.GetCode(ij.PrimarySalesPerson)
                                    .Cells(Me.SalesPerson_ColName.Index).Value = ObjSP.GetName(ij.PrimarySalesPerson)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_SP = e.RowIndex
                                            Me.bgwSalesPerson.RunWorkerAsync()
                                        Else
                                            Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvSalesPerson)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColSA_Primary.Index).Value = ij.PrimarySalesPerson
                                        .Cells(Me.SalesPerson_ColCode.Index).Value = ObjSP.GetCode(ij.PrimarySalesPerson)
                                        .Cells(Me.SalesPerson_ColName.Index).Value = ObjSP.GetName(ij.PrimarySalesPerson)
                                    End If
                                End If
                            ElseIf ij.PrimarySalesPerson = -1 Then
                                If Me.DgvSalesPerson.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_SP = e.RowIndex
                                        Me.bgwSalesPerson.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesPerson)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSalesPerson)
                        Me.DgvSalesPerson.Item(Me.SalesPerson_ColCode.Index, e.RowIndex).Value = Me.DgvSalesPerson.Item(Me.SalesPerson_ColCode.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvSalesPerson.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.SalesPerson_ColCode.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_SP = e.RowIndex
                                Me.bgwSalesPerson.RunWorkerAsync()
                            Else
                                Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvSalesPerson)
                            Exit Sub
                        Else
                            Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.SalesPerson_ColCode.Index Then
                            Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvSalesPerson.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvSalesPerson_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvSalesPerson.DataError
        Try
            ObjCD.SetNumber(Me.DgvSalesPerson)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesPerson_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvSalesPerson.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvSalesPerson.Rows(Me.DgvSalesPerson.Rows.Count - 1)
            dr.Cells(Me.SalesPerson_ColNo.Index).Value = Me.DgvSalesPerson.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesPerson_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSalesPerson.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvSalesPerson)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesPerson_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSalesPerson.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvSalesPerson, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_SP(ByVal vPrimary As Double, _
ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSalesPerson.Rows.Count - 1
                RowDGV = Me.DgvSalesPerson.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.SalesPerson_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub bgwSalesPerson_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwSalesPerson.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwSalesPerson_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwSalesPerson.RunWorkerCompleted
        Try
            Me.DgvSalesPerson.Rows.RemoveAt(EraseIndex_SP)
            Me.ObjCD.SetNumber(Me.DgvSalesPerson)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#End Region

    Private Sub btnFindSO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSO.Click
        Dim i As New FrmCariTransaksiSO
        i.BackColorAll = Me.BackColorAll

        Try
            'i.cboCriteria.SelectedIndex = 0
            If Me.TxtSONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtSONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.TxtPrimarySO.Text = i.PrimaryTransaksi
                Me.TxtSONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimarySO.Text)
            Else
                Me.TxtPrimarySO.Text = "-1"
                Me.TxtSONumber.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSONumber.GotFocus
        Try
            Me.TxtSONumber.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSONumber.KeyDown
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjSO.GetPrimary(Me.TxtSONumber.Text)
                If Primary <> -1 Then
                    Me.TxtPrimarySO.Text = Primary
                    Me.TxtSONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimarySO.Text)
                Else
                    Me.btnFindSO.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSONumber.Leave
        Try
            If Me.TxtSONumber.Text <> "" Then
                Dim Primary As Double = ObjSO.GetPrimary(Me.TxtSONumber.Text)
                If Primary <> -1 Then
                    Me.TxtPrimarySO.Text = Primary
                    Me.TxtSONumber.Text = Me.ObjSO.GetNoSO(Me.TxtPrimarySO.Text)
                Else
                    Me.btnFindSO.PerformClick()
                End If
            Else
                Me.TxtPrimarySO.Text = "-1"
                Me.TxtSONumber.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSONumber.LostFocus
        Try
            Me.TxtSONumber.BackColor = Me.TxtPrimarySO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim UserCodeList As New List(Of String)
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.DgvUser.AllowUserToAddRows = False
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'kosongkan dgv item
            Me.dgvItem.Rows.Clear()

            If Me.TxtPrimarySO.Text = "" Or Me.TxtPrimarySO.Text = "-1" Then
                If Me.ChkIgnoreLowTolerance.Checked = True Then
                    DTA.FillBy_All_UpTolerance(TBL, Me.dtpFrom.Value, Me.dtpTo.Value)
                Else
                    DTA.FillBy_All(TBL, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                End If
            Else
                If Me.ChkIgnoreLowTolerance.Checked = True Then
                    DTA.FillBy_PrimarySO_UpTolerance(TBL, Me.dtpFrom.Value, Me.dtpTo.Value, Me.TxtPrimarySO.Text)
                Else
                    DTA.FillBy_PrimarySO(TBL, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.TxtPrimarySO.Text)
                End If
            End If
            Me.GenerateFiltering()

            'pindahkan ke dgv
            For Each Row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Row.Amount, Row.Default_Precision)
                rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.FalseString
                rowDGV.Cells(Me.ColCodeNo.Index).Value = Row.Product_Code
                rowDGV.Cells(Me.ColCurrency.Index).Value = Row.Currency_Code
                rowDGV.Cells(Me.ColDate.Index).Value = Row.SO_Date
                rowDGV.Cells(Me.ColDesc.Index).Value = Row.Product_Description
                rowDGV.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Row.Disc_Amount, Row.Default_Precision)
                rowDGV.Cells(Me.ColDiscPerct.Index).Value = Me.ObjCD.DigitGrouping(Row.Discount, 2)
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                rowDGV.Cells(Me.ColSONo.Index).Value = Row.SO_No
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Row.SOB_Multi, Row.Decimal_Multi)
                rowDGV.Cells(Me.ColQtyCancel.Index).Value = 0
                If Me.cboCustomerFormat.SelectedIndex = 0 Then
                    rowDGV.Cells(Me.ColCustomer.Index).Value = Row.Customer_Code
                ElseIf Me.cboCustomerFormat.SelectedIndex = 1 Then
                    rowDGV.Cells(Me.ColCustomer.Index).Value = Row.Customer_Name
                Else
                    rowDGV.Cells(Me.ColCustomer.Index).Value = Row.Customer_Code + " - " + Row.Customer_Name
                End If

                If Row.Ship_to_Customer_Primary <> -1 Then
                    rowDGV.Cells(Me.ColCustShipto.Index).Value = Row.Ship_to_Customer_Code + " - " + Row.Ship_to_Customer_Name
                End If
                rowDGV.Cells(Me.ColUnit.Index).Value = Row.Unit_Multi
                rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Row.Unit_Price, Row.Default_Precision)
                rowDGV.Cells(Me.ColPrimaryPoduct.Index).Value = Row.Primary_Product
                rowDGV.Cells(Me.ColPrimaryUnitMulti.Index).Value = Row.Primary_Units_Multi
                rowDGV.Cells(Me.ColPrimarySO.Index).Value = Row.Primary_SO
                rowDGV.Cells(Me.ColRemarksItem.Index).Value = ""
                rowDGV.Cells(Me.ColPrimaruUnitDefault.Index).Value = Row.Primary_Unit_Default
                rowDGV.Cells(Me.ColQtyDefault.Index).Value = Me.ObjCD.DigitGrouping(Row.SOB_Multi * Row.Ratio_Default / Row.Ratio_Reference, Row._Decimal)
                rowDGV.Cells(Me.ColRatioDefault.Index).Value = Row.Ratio_Default
                rowDGV.Cells(Me.ColRatioReference.Index).Value = Row.Ratio_Reference
                rowDGV.Cells(Me.ColDefPrec_QtyDefault.Index).Value = Row._Decimal
                rowDGV.Cells(Me.ColUnitDefault.Index).Value = Row.Unit_Default
                rowDGV.Cells(Me.ColDecimalMulti.Index).Value = Row.Decimal_Multi
            Next

            Me.DgvUser.AllowUserToAddRows = True
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GenerateFiltering()
        Dim FilterProduct As String = ""
        Dim FilterSalesArea As String = ""
        Dim FilterSegment As String = ""
        Dim FilterCustomer As String = ""
        Dim FilterCustomerType As String = ""
        Dim FilterCustomerGroup As String = ""
        Dim FilterGroup As String = ""
        Dim FIlterSalesPerson As String = ""
        Dim FilterUser As String = ""
        Dim FilterShipto As String = ""
        Dim rowDGV As New DataGridViewRow
        Dim Filter As String = ""
        Dim row As DseTransaksiSOC_Quick.DataTable1Row
        Dim TBLCopyan As New DseTransaksiSOC_Quick.DataTable1DataTable
        Dim rowCopyan As DseTransaksiSOC_Quick.DataTable1Row
        Try
            If Me.DgvUser.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvUser.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCreatedBy.Index).Value) = False Then
                        If FilterUser = "" Then
                            FilterUser = "[Primary User] = " & rowDGV.Cells(Me.ColCreatedBy_Primary.Index).Value
                        Else
                            FilterUser &= " Or [Primary User] = " & rowDGV.Cells(Me.ColCreatedBy_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvMarketSegmentation.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvMarketSegmentation.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColMS_MarketSeg.Index).Value) = False Then
                        If FilterSegment = "" Then
                            FilterSegment = "[Primary Customer Segmentasi] = " & rowDGV.Cells(Me.ColMS_Primary.Index).Value
                        Else
                            FilterSegment &= " Or [Primary Customer Segmentasi] = " & rowDGV.Cells(Me.ColMS_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvSalesArea.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvSalesArea.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColSA_SalesArea.Index).Value) = False Then
                        If FilterSalesArea = "" Then
                            FilterSalesArea = "[Primary Sales Area] = " & rowDGV.Cells(Me.ColSA_Primary.Index).Value
                        Else
                            FilterSalesArea &= " Or [Primary Sales Area] = " & rowDGV.Cells(Me.ColSA_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvCustomer.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvCustomer.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColC_Code.Index).Value) = False Then
                        If FilterCustomer = "" Then
                            FilterCustomer = "[Primary Customer] = " & rowDGV.Cells(Me.ColC_Primary.Index).Value
                        Else
                            FilterCustomer &= " Or [Primary Customer] = " & rowDGV.Cells(Me.ColC_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvCustomerType.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvCustomerType.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColCT_CustType.Index).Value) = False Then
                        If FilterCustomerType = "" Then
                            FilterCustomerType = "[Primary Customer Type] = " & rowDGV.Cells(Me.ColCT_Primary.Index).Value
                        Else
                            FilterCustomerType &= " Or [Primary Customer Type] = " & rowDGV.Cells(Me.ColCT_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.dgvCustomerGroup.Rows.Count > 0 Then
                For Each rowDGV In Me.dgvCustomerGroup.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColSG_SalesGroup.Index).Value) = False Then
                        If FilterCustomerGroup = "" Then
                            FilterCustomerGroup = "[Primary Customer Group] = " & rowDGV.Cells(Me.ColSG_Primary.Index).Value
                        Else
                            FilterCustomerGroup &= " Or [Primary Customer Group] = " & rowDGV.Cells(Me.ColSG_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvSalesPerson.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvSalesPerson.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.SalesPerson_ColCode.Index).Value) = False Then
                        If FIlterSalesPerson = "" Then
                            FIlterSalesPerson = "[Primary Sales Person] = " & rowDGV.Cells(Me.SalesPerson_Primary.Index).Value
                        Else
                            FIlterSalesPerson &= " Or [Primary Sales Person] = " & rowDGV.Cells(Me.SalesPerson_Primary.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvShipto.Rows.Count > 0 Then
                For Each rowDGV In Me.DgvShipto.Rows
                    If String.IsNullOrEmpty(rowDGV.Cells(Me.ColShiptoCustCode.Index).Value) = False Then
                        If FilterShipto = "" Then
                            FilterShipto = "[Ship to Customer Primary] = " & rowDGV.Cells(Me.ColShiptoCustPrimary.Index).Value
                        Else
                            FilterShipto &= " Or [Ship to Customer Primary] = " & rowDGV.Cells(Me.ColShiptoCustPrimary.Index).Value
                        End If
                    End If
                Next
            End If

            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                    If FilterGroup = "" Then
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    Else
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " OR [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    End If
                End If
            Next

            If FilterSegment <> "" Then
                Filter = FilterSegment
            End If
            If FilterSalesArea <> "" Then
                If Filter = "" Then
                    Filter = FilterSalesArea
                Else
                    Filter = "(" & Filter & ") AND (" & FilterSalesArea & ") "
                End If
            End If
            If FilterCustomer <> "" Then
                If Filter = "" Then
                    Filter = FilterCustomer
                Else
                    Filter = "(" & Filter & ") AND (" & FilterCustomer & ") "
                End If
            End If
            If FilterCustomerType <> "" Then
                If Filter = "" Then
                    Filter = FilterCustomerType
                Else
                    Filter = "(" & Filter & ") AND (" & FilterCustomerType & ") "
                End If
            End If

            If FilterCustomerGroup <> "" Then
                If Filter = "" Then
                    Filter = FilterCustomerGroup
                Else
                    Filter = "(" & Filter & ") AND (" & FilterCustomerGroup & ") "
                End If
            End If

            If FIlterSalesPerson <> "" Then
                If Filter = "" Then
                    Filter = FIlterSalesPerson
                Else
                    Filter = "(" & Filter & ") AND (" & FIlterSalesPerson & ") "
                End If
            End If

            If FilterShipto <> "" Then
                If Filter = "" Then
                    Filter = FilterShipto
                Else
                    Filter = "(" & Filter & ") AND (" & FilterShipto & ") "
                End If
            End If

            If FilterGroup <> "" Then
                If Filter = "" Then
                    Filter = FilterGroup
                Else
                    Filter = "(" & Filter & ") AND (" & FilterGroup & ") "
                End If
            End If

            If FilterUser <> "" Then
                If Filter = "" Then
                    Filter = FilterUser
                Else
                    Filter = "(" & Filter & ") AND (" & FilterUser & ") "
                End If
            End If

            If Filter <> "" Then
                Me.TBL.DefaultView.RowFilter = Filter
                For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                    rowCopyan = Me.TBL.DefaultView(i).Row
                    TBLCopyan.ImportRow(rowCopyan)
                Next

                Me.TBL.Rows.Clear()
                For Each row In TBLCopyan.Rows
                    TBL.ImportRow(row)
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yg bisa diedit:
            '1. chk cancel
            '2. qty cancel+remarks item (jika chk cancel di check)
            If e.ColumnIndex = Me.ColCancel.Index Then
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            ElseIf e.ColumnIndex = Me.ColQtyCancel.Index Or e.ColumnIndex = Me.ColRemarksItem.Index Then
                If Me.dgvItem.Item(Me.ColCancel.Index, e.RowIndex).Value = Boolean.TrueString Then
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColQtyCancel.Index Then
                    'kembalikan ke nilai buffer
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColCancel.Index Then
                    'kemungkinan
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                        'di cek
                        Me.dgvItem.Item(Me.ColQtyCancel.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                        If ObjCD.CekBilanganPositif(Me.dgvItem.Item(Me.ColQtyCancel.Index, e.RowIndex).Value) = False Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                        End If
                    Else
                        'di un check
                        Me.dgvItem.Item(Me.ColQtyCancel.Index, e.RowIndex).Value = 0
                        Me.dgvItem.Item(Me.ColRemarksItem.Index, e.RowIndex).Value = ""
                    End If
                ElseIf e.ColumnIndex = Me.ColQtyCancel.Index Then
                    'pastikan non negatif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'pastikan nilainya lebih kecil dari qty
                    If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) Then
                        MsgBox("Maximum value your can enter is: " + Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value + " " + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value, _
                        MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'digit grouping
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem.Item(Me.ColDecimalMulti.Index, e.RowIndex).Value)

                    'hitung ulang qty default
                    Me.dgvItem.Item(Me.ColQtyDefault.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value * _
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value / Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColDefPrec_QtyDefault.Index, e.RowIndex).Value)
                ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Then
                    'trim
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    'truncate jika lebih panjang dr 150
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 150 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 150)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.ObjCD.setnumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim vPrimary As Double = -1

        Try
            If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColSONo.Index Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPO") = False Then
                    MsgBox("Insufficient acces.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    vPrimary = Me.dgvItem.Item(Me.ColPrimarySO.Index, e.RowIndex).Value
                    Dim ij As Sales_Module.FrmTransaksiSO
                    Dim j As Windows.Forms.Form

                    For Each j In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Sales_Module.FrmTransaksiSO)
                        If IsNothing(ij) = False Then
                            ij.BringToFront()
                            ij.Focus()
                            ij.ShowTransaction(vPrimary)
                            Exit Sub
                        End If
                    Next

                    Dim i As New Sales_Module.FrmTransaksiSO
                    ' Increment the caption counter.
                    i.UserCode = Me.UserCode.ToString
                    ' Set the caption to be unique.
                    i.BAckColorAll = Core_Module.GetSMESSBackColorAll
                    i.MdiParent = Me.ParentForm
                    i.Show()
                    i.ShowTransaction(vPrimary)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim AdaYgDicancel As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim ListPrimarySO As New List(Of Double)
        Dim PrimarySO As Double = -1
        Dim ij As New FrmTransaksiSOC
        Dim CustName, CustCode, CustPONo, ContactCP, PhoneNo As String

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.TrueString Then
                    AdaYgDicancel = True
                    Exit For
                End If
            Next
            If AdaYgDicancel = True Then
                'cancellation reason harus ada nilainya
                Me.TxtCancellationReason.Text = Me.TxtCancellationReason.Text.Trim
                If Me.TxtCancellationReason.Text = "" Then
                    MsgBox("Please specify cancellation reason.", MsgBoxStyle.Information, "Information")
                    Me.TxtCancellationReason.Focus()
                    Exit Sub
                End If

                If MsgBox("Start cancellation process?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    '1. isikan list po
                    For Each rowDGV In Me.dgvItem.Rows
                        If rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.TrueString Then
                            If ListPrimarySO.Contains(rowDGV.Cells(Me.ColPrimarySO.Index).Value) = False Then
                                ListPrimarySO.Add(rowDGV.Cells(Me.ColPrimarySO.Index).Value)
                            End If
                        End If
                    Next

                    Me.prg01.Visible = True
                    Me.prg01.Value = 0
                    Me.prg01.Maximum = ListPrimarySO.Count
                    ij.prgProgress.Value = 0
                    ij.prgProgress.Maximum = Me.prg01.Maximum
                    Me.LblProgress.Text = ""
                    Me.LblProgress.Visible = True

                    '2. untuk setiap list po, buat poc, lengkapi item nya lalu simpan
                    ij.vFromQuick = True
                    ij.vDirectlySave = False
                    ij.UserCode = Me.UserCode
                    ij.Show()
                    ij.Location = New Point((ij.Width * -1) - 5, (ij.Height * -1) - 5)
                    For Each PrimarySO In ListPrimarySO
                        ij.SetBaru()
                        ij.txtPrimarySO.Text = PrimarySO
                        ij.txtSONO.Text = Me.ObjSO.GetNoSO(PrimarySO)
                        ij.vShowDialog = False

                        Me.prg01.Value += 1
                        Me.LblProgress.Text = "Data " + Me.prg01.Value.ToString + " of " + Me.prg01.Maximum.ToString + " - SO Number: " + ij.txtSONO.Text
                        ij.prgProgress.Value = Me.prg01.Value
                        ij.lblProgress.Visible = True
                        ij.lblProgress.Text = Me.LblProgress.Text

                        ij.TxtPrimaryCustomer.Text = ObjSO.GetPrimaryCustomer(PrimarySO)
                        Me.ObjSO.GetCustCode_CustName_CustPONO_ContactPerson_PhoneNo(PrimarySO, CustName, CustCode, CustPONo, ContactCP, PhoneNo)
                        ij.txtCustomerCode.Text = CustCode
                        ij.TxtCustomerName.Text = CustName
                        ij.txtContactPerson.Text = ContactCP
                        ij.txtPhoneNo.Text = PhoneNo
                        ij.txtCustPONO.Text = CustPONo

                        ij.txtCancellationReason.Text = Me.TxtCancellationReason.Text.Trim
                        ij.TxtRemark.Text = "Quick SO Cancellation"

                        For Each rowDGV1 As DataGridViewRow In Me.dgvItem.Rows
                            If rowDGV1.Cells(Me.ColPrimarySO.Index).Value = PrimarySO Then
                                ij.dgvItem.Rows.Add()
                                rowDGV = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                                With rowDGV
                                    .Cells(ij.ColNo.Index).Value = ij.dgvItem.Rows.Count
                                    .Cells(ij.ColCodeNo.Index).Value = rowDGV1.Cells(Me.ColCodeNo.Index).Value
                                    .Cells(ij.ColDescription.Index).Value = rowDGV1.Cells(Me.ColDesc.Index).Value
                                    .Cells(ij.ColQtyMulti.Index).Value = rowDGV1.Cells(Me.ColQty.Index).Value
                                    .Cells(ij.ColQty.Index).Value = rowDGV1.Cells(Me.ColQtyDefault.Index).Value
                                    .Cells(ij.ColUnit.Index).Value = rowDGV1.Cells(Me.ColUnitDefault.Index).Value
                                    .Cells(ij.ColRemark.Index).Value = rowDGV1.Cells(Me.ColRemarksItem.Index).Value
                                    .Cells(ij.ColPrimaryCodeNo.Index).Value = rowDGV1.Cells(Me.ColPrimaryPoduct.Index).Value
                                    .Cells(ij.ColPrimaryUnits.Index).Value = rowDGV1.Cells(Me.ColPrimaruUnitDefault.Index).Value
                                    .Cells(ij.ColQtyMulti.Index).Value = rowDGV1.Cells(Me.ColQtyCancel.Index).Value

                                    .Cells(ij.ColUnitsMulti.Index).Value = rowDGV1.Cells(Me.ColUnit.Index).Value
                                    If .Cells(ij.ColUnitsMulti.Index).Value <> .Cells(ij.ColUnit.Index).Value Then
                                        .Cells(ij.ColRatio.Index).Value = Me.ObjCD.DigitGrouping(rowDGV1.Cells(Me.ColRatioReference.Index).Value, rowDGV1.Cells(Me.ColDecimalMulti.Index).Value) _
                                                & " " & rowDGV1.Cells(Me.ColUnit.Index).Value & " = " & Me.ObjCD.DigitGrouping(rowDGV1.Cells(Me.ColRatioDefault.Index).Value, _
                                                rowDGV1.Cells(Me.ColDefPrec_QtyDefault.Index).Value) & " " & rowDGV1.Cells(Me.ColUnitDefault.Index).Value
                                    Else
                                        .Cells(ij.ColRatio.Index).Value = "-"
                                    End If
                                    .Cells(ij.ColRatioReference.Index).Value = rowDGV1.Cells(Me.ColRatioReference.Index).Value
                                    .Cells(ij.ColRatioDefault.Index).Value = rowDGV1.Cells(Me.ColRatioDefault.Index).Value
                                    .Cells(ij.ColPrimaryUnitsMulti.Index).Value = rowDGV1.Cells(Me.ColPrimaryUnitMulti.Index).Value
                                End With
                            End If
                        Next
                        ij.Hide()
                        ij.vDirectlySave = Me.chkDirectSave.Checked
                        ij.vShowDialog = True
                        ij.ShowDialog()
                        If ij.vStopCancellation = True Then
                            Exit For
                        End If
                    Next
                    '3. 
                End If

                Me.btnFilter.PerformClick()
            Else
                MsgBox("Please specify at least 1 item to cancel.", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimarySO = Nothing
            ij = Nothing
            Me.prg01.Value = 0
            Me.prg01.Visible = False
            Me.LblProgress.Visible = False
        End Try
    End Sub

    Private Sub llbCheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbCheckAll.Click
        Dim rowDGV As DataGridViewRow
        Try
            If MsgBox("Check all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvItem.Rows
                    rowDGV.Cells(Me.ColQtyCancel.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimalMulti.Index).Value)
                    If ObjCD.CekBilanganPositif(rowDGV.Cells(Me.ColQtyCancel.Index).Value) = True Then
                        rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.TrueString
                    End If
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub llbUncheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbUncheckAll.Click
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Uncheck all data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                For Each rowDGV In Me.dgvItem.Rows
                    rowDGV.Cells(Me.ColCancel.Index).Value = Boolean.FalseString
                    rowDGV.Cells(Me.ColQtyCancel.Index).Value = 0
                    rowDGV.Cells(Me.ColRemarksItem.Index).Value = ""
                Next
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub LlbClearFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LlbClearFilter.Click
        Try
            Me.DgvShipto.Rows.Clear()
            Me.DgvUser.Rows.Clear()
            Me.DgvCustomer.Rows.Clear()
            Me.dgvCustomerGroup.Rows.Clear()
            Me.DgvCustomerType.Rows.Clear()
            Me.dgvItemProductGroup.Rows.Clear()
            Me.DgvSalesPerson.Rows.Clear()
            Me.DgvMarketSegmentation.Rows.Clear()
            Me.DgvSalesArea.Rows.Clear()
            Me.dtpFrom.Value = Now
            Me.dtpTo.Value = Now
            Me.ChkIgnoreLowTolerance.Checked = False
            Me.chkDirectSave.Checked = False
            Me.TxtPrimarySO.Text = "-1"
            Me.TxtSONumber.Clear()
            Me.cboCustomerFormat.SelectedIndex = 0
            Me.btnFilter.Focus()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCancellationReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCancellationReason.GotFocus
        Try
            Me.TxtCancellationReason.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCancellationReason_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCancellationReason.LostFocus
        Try
            Me.TxtCancellationReason.BackColor = Me.TxtPrimarySO.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvUser.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvUser.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub bgwShipto_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwShipto.DoWork, bgwCustomer.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwShipto_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwShipto.RunWorkerCompleted, bgwCustomer.RunWorkerCompleted
        Try
            Me.DgvShipto.Rows.RemoveAt(EraseIndex_ST)
            Me.ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvShipto_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvShipto.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColShiptoCustCode.Index Then
                Me.buffer = Me.DgvShipto(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvShipto_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvShipto.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvShipto.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjRef As New Core_Module.clsMasterCustomer
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColShiptoCustCode.Index Then
                        .Cells(Me.ColShiptoCustCode.Index).Value = .Cells(Me.ColShiptoCustCode.Index).Value.ToString.ToUpper
                        If ObjRef.ApakahSudahAda(.Cells(Me.ColShiptoCustCode.Index).Value) = True Then
                            primary = ObjRef.GetPrimary(.Cells(Me.ColShiptoCustCode.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_ST(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColShiptoCustPrimary.Index).Value = primary
                                .Cells(Me.ColShiptoCustCode.Index).Value = ObjRef.GetCustomerCode(primary)
                                .Cells(Me.ColShiptoCustName.Index).Value = ObjRef.GetCustName(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvShipto.Item(Me.ColShiptoCustCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_ST = e.RowIndex
                                        Me.bgwShipto.RunWorkerAsync()
                                    Else
                                        Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvShipto)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomer
                            ij.txtCari.Text = Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomer <> -1 Then
                                index = Me.GetIndexItem_ST(ij.PrimaryCustomer, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColShiptoCustPrimary.Index).Value = ij.PrimaryCustomer
                                    .Cells(Me.ColShiptoCustCode.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                    .Cells(Me.ColShiptoCustName.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseIndex_ST = e.RowIndex
                                            Me.bgwShipto.RunWorkerAsync()
                                        Else
                                            Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvShipto)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColShiptoCustPrimary.Index).Value = ij.PrimaryCustomer
                                        .Cells(Me.ColShiptoCustCode.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                        .Cells(Me.ColShiptoCustName.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomer = -1 Then
                                If Me.DgvShipto.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseIndex_ST = e.RowIndex
                                        Me.bgwShipto.RunWorkerAsync()
                                    Else
                                        Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvShipto)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvShipto)
                        Me.DgvShipto.Item(Me.ColShiptoCustCode.Index, e.RowIndex).Value = Me.DgvShipto.Item(Me.ColShiptoCustCode.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvShipto.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColShiptoCustCode.Index Then
                            If String.IsNullOrEmpty(Me.buffer) = True Then
                                Me.EraseIndex_ST = e.RowIndex
                                Me.bgwShipto.RunWorkerAsync()
                            Else
                                Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvShipto)
                            Exit Sub
                        Else
                            Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColShiptoCustCode.Index Then
                            Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Else
                            Me.DgvShipto.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvShipto_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvShipto.RowsAdded
        Dim dr As DataGridViewRow
        Try
            dr = Me.DgvShipto.Rows(Me.DgvShipto.Rows.Count - 1)
            dr.Cells(Me.ColShiptoNo.Index).Value = Me.DgvShipto.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvShipto_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvShipto.RowsRemoved
        Me.ObjCD.SetNumber(Me.DgvShipto)
    End Sub

    Private Sub DgvShipto_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvShipto.Sorted
        Me.ObjCD.SetNumber(Me.DgvShipto)
    End Sub

    Private Function GetIndexItem_ST(ByVal vPrimary As Double, _
     ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvShipto.Rows.Count - 1
                RowDGV = Me.DgvShipto.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColShiptoCustPrimary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class