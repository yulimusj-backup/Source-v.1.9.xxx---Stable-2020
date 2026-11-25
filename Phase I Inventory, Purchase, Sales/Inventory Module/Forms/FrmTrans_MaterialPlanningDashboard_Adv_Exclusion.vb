Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_Adv_Exclusion
    Public UserCode As String = ""
    Public TipeTransaksiExclusion As ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion = ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion.MaterialPlanning
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ShowWarehouseExclusion As Boolean = True
    Public ModifySupplierSelectionMode As Boolean = False
    Public SupplierSelectionMode As Core_Module.clsMasterProduct.EnumSupplierSelectionMode _
    = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order

    Public Enum EnumExcludeInclude As Integer
        Include
        Exclude
    End Enum

    Public Enum EnumLocationLevel As Integer
        Warehouse
        Location
    End Enum

    Dim buffer As String = ""
    Dim DummyIndex As Integer = -1
    Dim objItem As New ClsTransaksiMaterialPlanningGroupItem
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1
    Dim Filter As String = ""
    Dim ObjWarehouse As New Core_Module.clsMasterWarehouse
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Dim LangsungTutup As Boolean = False
    Dim TipeExclusion As ClsTransaksiMaterialPlanning.EnumTipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude

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
                If Buffer = "" Then
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
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
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
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
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

    Public Sub IsiDGV_ItemGroup()
        Dim DTAGroup As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_adv_exclusionTableAdapter
        Dim TBLGroup As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable
        Dim rowGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionRow
        Dim AllowToAddRows_ItemGroup As Boolean
        Dim rowDGV As DataGridViewRow
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim DTAProduct2 As New DseTransaksiMaterialPlanningGroupTableAdapters.TblProductTableAdapter
        Dim TBLProduct2 As New DseTransaksiMaterialPlanningGroup.TblProductDataTable
        Dim rowProduct2 As DseTransaksiMaterialPlanningGroup.TblProductRow
        Dim DTACategory As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_categoryTableAdapter
        Dim TBLCategory As New DseTransaksiMaterialPlanningGroup.tblmaster_categoryDataTable
        Dim rowCategory As DseTransaksiMaterialPlanningGroup.tblmaster_categoryRow
        Dim DTAProductGroup2 As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_product_groupTableAdapter
        Dim TblProductGroup2 As New DseTransaksiMaterialPlanningGroup.tblmaster_product_groupDataTable
        Dim rowProductGroup2 As DseTransaksiMaterialPlanningGroup.tblmaster_product_groupRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAGroup.Connection.ConnectionString = Core_Module.GetConstring
            DTAProduct2.Connection.ConnectionString = Core_Module.GetConstring
            DTACategory.Connection.ConnectionString = Core_Module.GetConstring
            DTAProductGroup2.Connection.ConnectionString = Core_Module.GetConstring

            Me.dgvItemProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemProductGroup.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItemProductGroup.Rows.Clear()
            AllowToAddRows_ItemGroup = Me.dgvItemProductGroup.AllowUserToAddRows
            Me.dgvItemProductGroup.AllowUserToAddRows = False


            DTAGroup.Fill(TBLGroup, Me.UserCode, Me.TipeTransaksiExclusion)
            DTAProduct2.Fill(TBLProduct2)
            For Each rowGroup In TBLGroup.Rows
                Me.dgvItemProductGroup.Rows.Add()
                rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
                With rowGroup
                    rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = .remarks
                    rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = .primary_acuan
                    ProductLevel = .level
                    rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(ProductLevel)
                    If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Filter = "[Primary Brand] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Filter = "[Primary Category] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                        Filter = "[Primary] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                    End If

                    TBLProduct2.DefaultView.RowFilter = Filter
                    If TBLProduct2.DefaultView.Count = 0 Then
                        If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                            If TBLCategory.Rows.Count = 0 Then
                                DTACategory.Fill(TBLCategory)
                            End If
                            TBLCategory.DefaultView.RowFilter = "Primary = " + .primary_acuan.ToString
                            If TBLCategory.DefaultView.Count > 0 Then
                                rowCategory = TBLCategory.DefaultView(0).Row
                                rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = rowCategory.Category_Code
                                rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = rowCategory.Description
                            Else
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Or _
                        ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Or _
                        ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                            If TblProductGroup2.Rows.Count = 0 Then
                                DTAProductGroup2.Fill(TblProductGroup2)
                            End If
                            TblProductGroup2.DefaultView.RowFilter = "Primary = " + .primary_acuan.ToString
                            If TblProductGroup2.DefaultView.Count > 0 Then
                                rowProductGroup2 = TblProductGroup2.DefaultView(0).Row
                                rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = rowProductGroup2.Group_Code
                                rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = rowProductGroup2.Description
                            Else
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                    Else
                        rowProduct2 = TBLProduct2.DefaultView(0).Row
                        If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = rowProduct2.Brand_Code
                            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = rowProduct2.Brand_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = rowProduct2.Category_Code
                            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = rowProduct2.Category_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = rowProduct2.Product_Code
                            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = rowProduct2.Product_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = rowProduct2.Product_Group_Code
                            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = rowProduct2.Product_Group_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = rowProduct2.Sub_Brand_Code
                            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = rowProduct2.Sub_Brand_Desc
                        End If
                    End If
                End With
            Next

            Me.dgvItemProductGroup.AllowUserToAddRows = AllowToAddRows_ItemGroup
            Me.dgvItemProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If TBLGroup.Rows.Count = 0 Then
                Me.rboTipe_Exclusion.Checked = True
            Else
                rowGroup = TBLGroup.Rows(0)
                TipeExclusion = rowGroup.tipe_exclusion
                If TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude Then
                    Me.rboTipe_Exclusion.Checked = True
                Else
                    Me.rboTipe_Inclusion.Checked = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGroup.Dispose()
            TBLGroup.Dispose()
            DTAProduct2.Dispose()
            TBLProduct2.Dispose()
            DTACategory.Dispose()
            TBLCategory.Dispose()
            DTAProductGroup2.Dispose()
            TblProductGroup2.Dispose()
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Adv_Exclusion_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_adv_exclusionTableAdapter
        Dim DummyVal As String = ""

        Try
            'hapuskan data yang primary acuannya -1
            DTA.DeleteQuery_PrimaryAcuanM1()
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvWarehouse.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.IsiDGV_ItemGroup()
            If Me.ShowWarehouseExclusion = True Then
                'load data exclude_inculde warehouse
                Me.LoadDataIncldeExcludeWarehouse()
            Else
                Me.TabControl1.TabPages.RemoveAt(1)
            End If

            DummyVal = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_SupplierSelectionMode, Me.UserCode)
            If DummyVal = "" Then
                Me.cboSupplierSelectionMode.SelectedIndex = 0
            Else
                Me.cboSupplierSelectionMode.SelectedIndex = DummyVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Adv_Exclusion_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim RetVal As MsgBoxResult

        Try
            If Me.LangsungTutup = False Then
                RetVal = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm") = MsgBoxResult.Yes

                If RetVal = MsgBoxResult.Yes Then
                    If Me.ShowWarehouseExclusion = True Then
                        If Me.ApakahDataInculdeExcludeWarehouseBenar = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                    Me.SimpanData()
                    If Me.ShowWarehouseExclusion = True Then
                        Me.SimpanDataIncludeExcludeWarehouse()
                    End If

                    'If Me.cboSupplierSelectionMode.SelectedIndex = 0 Then
                    '    Me.ModifySupplierSelectionMode = False
                    'Else
                    '    Me.ModifySupplierSelectionMode = True
                    '    If Me.cboSupplierSelectionMode.SelectedIndex = 1 Then
                    '        Me.SupplierSelectionMode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order
                    '    ElseIf Me.cboSupplierSelectionMode.SelectedIndex = 2 Then
                    '        Me.SupplierSelectionMode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier
                    '    ElseIf Me.cboSupplierSelectionMode.SelectedIndex = 3 Then
                    '        Me.SupplierSelectionMode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO
                    '    End If
                    'End If
                ElseIf RetVal = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanData()
        Dim DTA As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_adv_exclusionTableAdapter

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTA.Delete(Me.UserCode, Me.TipeTransaksiExclusion)
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            If Me.rboTipe_Exclusion.Checked = True Then
                Me.TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude
            Else
                Me.TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Include
            End If
            For Each rowdgv As DataGridViewRow In Me.dgvItemProductGroup.Rows
                If rowdgv.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                    DTA.Insert(-1, Me.objItem.GetProductLevel(rowdgv.Cells(Me.ColProductGroup_Level.Index).Value), rowdgv.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value, _
                    rowdgv.Cells(Me.ColProductGroup_Remarks.Index).Value, Me.UserCode, Me.TipeExclusion, Me.TipeTransaksiExclusion)
                End If
            Next
            Me.dgvItemProductGroup.AllowUserToAddRows = True

            Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_SupplierSelectionMode, Me.cboSupplierSelectionMode.SelectedIndex, Me.UserCode)

            If Me.cboSupplierSelectionMode.SelectedIndex = 0 Then
                Me.ModifySupplierSelectionMode = False
            Else
                Me.ModifySupplierSelectionMode = True
                If Me.cboSupplierSelectionMode.SelectedIndex = 1 Then
                    Me.SupplierSelectionMode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order
                ElseIf Me.cboSupplierSelectionMode.SelectedIndex = 2 Then
                    Me.SupplierSelectionMode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier
                ElseIf Me.cboSupplierSelectionMode.SelectedIndex = 3 Then
                    Me.SupplierSelectionMode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ShowWarehouseExclusion = True Then
                If Me.ApakahDataInculdeExcludeWarehouseBenar = False Then
                    Exit Sub
                End If
            End If
            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.SimpanData()
                If Me.ShowWarehouseExclusion = True Then
                    Me.SimpanDataIncludeExcludeWarehouse()
                End If
                Me.LangsungTutup = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Close and discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.LangsungTutup = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGVWarehouse()
        Dim Dummy As String = ""
        Dim DTA As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_setting_warehouse_exclusionTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_setting_warehouse_exclusionDataTable
        Dim row As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_setting_warehouse_exclusionRow
        Dim rowDGV As DataGridViewRow
        Dim DTA_W As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_warehouseTableAdapter
        Dim TBL_W As New DseTransaksiMaterialPlanning2.tblmaster_warehouseDataTable
        Dim row_W As DseTransaksiMaterialPlanning2.tblmaster_warehouseRow
        Dim DTA_L As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_storage_areaTableAdapter
        Dim TBL_L As New DseTransaksiMaterialPlanning2.tblmaster_storage_areaDataTable
        Dim row_L As DseTransaksiMaterialPlanning2.tblmaster_storage_areaRow
        Dim LocationLevel1 As EnumLocationLevel = EnumLocationLevel.Location
        Dim Inc_Exc As EnumExcludeInclude = EnumExcludeInclude.Include

        Try
            Me.chkUseWarehouseExclusion.Checked = True
            Me.chkUseWarehouseExclusion.Checked = False

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_W.Connection.ConnectionString = Core_Module.GetConstring
            DTA_L.Connection.ConnectionString = Core_Module.GetConstring

            Me.dgvWarehouse.AllowUserToAddRows = False
            Me.dgvWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvWarehouse.Rows.Clear()

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_UseWarehouseExclusion, Me.UserCode)
            If Dummy = "" Or Dummy = "0" Then
                Me.chkUseWarehouseExclusion.Checked = False
            Else
                Me.chkUseWarehouseExclusion.Checked = True
                Me.dgvWarehouse.AllowUserToAddRows = False
                DTA.FillBy_kode_user(TBL, Me.UserCode)
                For Each row In TBL.Rows
                    Me.dgvWarehouse.Rows.Add()
                    rowDGV = Me.dgvWarehouse.Rows(Me.dgvWarehouse.Rows.Count - 1)

                    rowDGV.Cells(Me.ColWarehouse_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value = row.primary_acuan
                    If row.level = LocationLevel1 Then
                        If TBL_L.Rows.Count = 0 Then
                            DTA_L.Fill(TBL_L)
                        End If
                        TBL_L.DefaultView.RowFilter = "Primary = " + row.primary_acuan.ToString
                        If TBL_L.DefaultView.Count = 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                        row_L = TBL_L.DefaultView(0).Row
                        'location
                        rowDGV.Cells(Me.ColWarehouse_Code.Index).Value = row_L.Storage_Area_Code
                        rowDGV.Cells(Me.ColWarehouse_Desc.Index).Value = row_L.Description
                        rowDGV.Cells(Me.ColWarehouse_Level.Index).Value = "Location"
                    Else
                        'warehouse
                        If TBL_W.Rows.Count = 0 Then
                            DTA_W.Fill(TBL_W)
                        End If
                        TBL_W.DefaultView.RowFilter = "Primary = " + row.primary_acuan.ToString
                        If TBL_W.DefaultView.Count = 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                        row_W = TBL_W.DefaultView(0).Row
                        'location
                        rowDGV.Cells(Me.ColWarehouse_Code.Index).Value = row_W.Warehouse_Code
                        rowDGV.Cells(Me.ColWarehouse_Desc.Index).Value = row_W.Description
                        rowDGV.Cells(Me.ColWarehouse_Level.Index).Value = "Warehouse"
                    End If
                    If row.include_exclude = Inc_Exc Then
                        rowDGV.Cells(Me.ColWarehouse_ExcludeInclude.Index).Value = "Include"
                    Else
                        rowDGV.Cells(Me.ColWarehouse_ExcludeInclude.Index).Value = "Exclude"
                    End If
                    rowDGV.Cells(Me.ColWarehouse_Remarks.Index).Value = row.remarks
                Next
                Me.dgvWarehouse.AllowUserToAddRows = True
            End If

            If Me.chkUseWarehouseExclusion.Checked = True Then
                Me.dgvWarehouse.AllowUserToAddRows = True
            End If
            Me.dgvWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA_W.Dispose()
            TBL_W.Dispose()
            DTA_L.Dispose()
            TBL_L.Dispose()
        End Try
    End Sub

    Private Sub LoadDataIncldeExcludeWarehouse()
        Dim Dummy As String = ""

        Try
            Me.chkUseWarehouseExclusion.Checked = True
            Me.chkUseWarehouseExclusion.Checked = False

            Me.dgvWarehouse.AllowUserToAddRows = False
            Me.dgvWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvWarehouse.Rows.Clear()

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_UseWarehouseExclusion, Me.UserCode)
            If Dummy = "" Or Dummy = "0" Then
                Me.chkUseWarehouseExclusion.Checked = False
            Else
                Me.chkUseWarehouseExclusion.Checked = True
                Me.LoadDGVWarehouse()
            End If

            If Me.chkUseWarehouseExclusion.Checked = True Then
                Me.dgvWarehouse.AllowUserToAddRows = True
            End If
            Me.dgvWarehouse.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseWarehouseExclusion_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseWarehouseExclusion.CheckedChanged
        Try
            If Me.chkUseWarehouseExclusion.Checked = True Then
                Me.dgvWarehouse.Enabled = True
                Me.dgvWarehouse.AllowUserToAddRows = True
                Me.dgvWarehouse.ReadOnly = False
                Me.dgvWarehouse.AllowUserToDeleteRows = True
            Else
                Me.dgvWarehouse.Enabled = False
                Me.dgvWarehouse.AllowUserToAddRows = False
                Me.dgvWarehouse.ReadOnly = True
                Me.dgvWarehouse.AllowUserToDeleteRows = False
                Me.dgvWarehouse.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWarehouse_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvWarehouse.CellBeginEdit
        Dim rowDGV, rowDGV_b4 As DataGridViewRow

        Try
            rowDGV = Me.dgvWarehouse.Rows(e.RowIndex)

            If rowDGV.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value = -1 Then
                If e.ColumnIndex = Me.ColWarehouse_Code.Index Or e.ColumnIndex = Me.ColWarehouse_Level.Index Then
                    'jika primary acuan =-1, hanya bisa edit bila yang di edit adalah kolom level dan primary acuan baris atasnya <> -1
                    If rowDGV.Index <> 0 Then
                        rowDGV_b4 = Me.dgvWarehouse.Rows(rowDGV.Index - 1)
                        If rowDGV_b4.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value = -1 Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    Me.buffer = Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            Else
                'jika primary acuan <> -1, hanya bisa edit kolom: remarks, kode acuan, exclude/include
                If e.ColumnIndex = Me.ColWarehouse_Remarks.Index Or _
                e.ColumnIndex = Me.ColWarehouse_Code.Index Or _
                e.ColumnIndex = Me.ColWarehouse_ExcludeInclude.Index Then
                    Me.buffer = Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexAcuan(ByVal vLevel As String, ByVal vPrimary As Double, ByVal vIndexAcuan As Integer) As Double
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvWarehouse.Rows
                If rowDGV.Cells(Me.ColWarehouse_Level.Index).Value = vLevel And _
                rowDGV.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value = vPrimary And _
                rowDGV.Index <> vIndexAcuan Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub EndEditWarehouse(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.dgvWarehouse.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryWarehouse As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row
                If IsNothing(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColWarehouse_Code.Index Then
                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        'kemungkinan
                        If Me.ObjWarehouse.ApakahSudahAda(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryWarehouse = Me.ObjWarehouse.GetPrimary(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexAcuan(Me.dgvWarehouse.Item(Me.ColWarehouse_Level.Index, e.RowIndex).Value.ToString, PrimaryWarehouse, e.RowIndex)

                            If index < 0 Then
                                Me.dgvWarehouse.Item(Me.ColWarehouse_PrimaryAcuan.Index, e.RowIndex).Value = PrimaryWarehouse
                                Me.dgvWarehouse.Item(Me.ColWarehouse_Desc.Index, e.RowIndex).Value = Me.ObjWarehouse.GetDescription(PrimaryWarehouse)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Warehouse code duplication, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) = False Then
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value = Me.buffer
                                    Else
                                        Me.DummyIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterWarehouse
                            ij.txtCari.Text = Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryWarehouse <> -1 Then
                                index = Me.GetIndexAcuan(Me.dgvWarehouse.Item(Me.ColWarehouse_Level.Index, e.RowIndex).Value, ij.PrimaryWarehouse, e.RowIndex)

                                If index < 0 Then
                                    Me.dgvWarehouse.Item(Me.ColWarehouse_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryWarehouse
                                    Me.dgvWarehouse.Item(Me.ColWarehouse_Desc.Index, e.RowIndex).Value = Me.ObjWarehouse.GetDescription(ij.PrimaryWarehouse)
                                    Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value = Me.ObjWarehouse.GetWarehouseCode(ij.PrimaryWarehouse)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Warehouse code duplication, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.DummyIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                        Exit Sub
                                    Else
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryWarehouse
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_Desc.Index, e.RowIndex).Value = Me.ObjWarehouse.GetDescription(ij.PrimaryWarehouse)
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value = Me.ObjWarehouse.GetWarehouseCode(ij.PrimaryWarehouse)
                                    End If
                                End If
                            ElseIf ij.PrimaryWarehouse = -1 Then
                                If Me.dgvWarehouse.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.DummyIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                    Exit Sub
                                End If
                                If Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.DummyIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvWarehouse)
                    End If
                Else
                    Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocation(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.dgvWarehouse.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryLocation As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row
                If IsNothing(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColWarehouse_Code.Index Then
                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        'kemungkinan
                        If Me.ObjLocation.ApakahSudahAda(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryLocation = Me.ObjLocation.GetPrimary(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexAcuan(Me.dgvWarehouse.Item(Me.ColWarehouse_Level.Index, e.RowIndex).Value.ToString, PrimaryLocation, e.RowIndex)

                            If index < 0 Then
                                Me.dgvWarehouse.Item(Me.ColWarehouse_PrimaryAcuan.Index, e.RowIndex).Value = PrimaryLocation
                                Me.dgvWarehouse.Item(Me.ColWarehouse_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(PrimaryLocation)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Location code duplication, please input another code.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) = False Then
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value = Me.buffer
                                    Else
                                        Me.DummyIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterStorageArea
                            ij.txtCari.Text = Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryStorageArea <> -1 Then
                                index = Me.GetIndexAcuan(Me.dgvWarehouse.Item(Me.ColWarehouse_Level.Index, e.RowIndex).Value, ij.PrimaryStorageArea, e.RowIndex)

                                If index < 0 Then
                                    Me.dgvWarehouse.Item(Me.ColWarehouse_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                    Me.dgvWarehouse.Item(Me.ColWarehouse_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                    Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Location code duplication, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.DummyIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                        Exit Sub
                                    Else
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_Desc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                        Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                    End If
                                End If
                            ElseIf ij.PrimaryStorageArea = -1 Then
                                If Me.dgvWarehouse.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.DummyIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                    Exit Sub
                                End If
                                If Me.dgvWarehouse.Item(Me.ColWarehouse_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.DummyIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvWarehouse)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvWarehouse)
                    End If
                Else
                    Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWarehouse_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWarehouse.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvWarehouse.Rows(e.RowIndex)

            ''' pastikan ada isinya
            With row
                If IsNothing(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    ''' Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If

                    If e.ColumnIndex = Me.ColWarehouse_Remarks.Index Then
                        If Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value <> "" Then
                            Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Trim(Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value)
                            If Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                                Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColWarehouse_Code.Index Then
                        If Me.dgvWarehouse.Item(Me.ColWarehouse_Level.Index, e.RowIndex).Value = "Warehouse" Then
                            Me.EndEditWarehouse(sender, e)
                        Else
                            Me.EndEditLocation(sender, e)
                        End If
                        Exit Sub
                    ElseIf e.ColumnIndex = Me.ColWarehouse_ExcludeInclude.Index Then
                        'do nothing
                    ElseIf e.ColumnIndex = Me.ColWarehouse_Level.Index Then
                        'do nothing
                    End If
                Else
                    If e.ColumnIndex = Me.ColWarehouse_Remarks.Index Then
                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    ElseIf e.ColumnIndex = Me.ColWarehouse_Code.Index Then
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.DummyIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If

                        Me.ObjCD.SetNumber(Me.dgvWarehouse)
                        Exit Sub
                    Else
                        Me.dgvWarehouse.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataInculdeExcludeWarehouseBenar() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            'baris yang primary acuannya = -1, dihapuskan saja
            Me.dgvWarehouse.AllowUserToAddRows = False
            If Me.dgvWarehouse.RowCount > 0 Then
                For ij As Double = Me.dgvWarehouse.Rows.Count - 1 To 0 Step 1
                    rowDGV = Me.dgvWarehouse.Rows(ij)
                    If rowDGV.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value = -1 Then
                        Me.dgvWarehouse.Rows.Remove(rowDGV)
                    End If
                Next
            End If

            'tidak boleh ada duplication code untuk level yang sama
            'ini sudah dihandle di cell end edit

            'jika exclusion warehouse aktif, maka minimal harus ada 1 item baris dgv
            If Me.chkUseWarehouseExclusion.Checked = True Then
                If Me.dgvWarehouse.Rows.Count <= 0 Then
                    Me.TabControl1.SelectedIndex = 1
                    MsgBox("Please enter at least 1 data row.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'hal sebaliknya juga berlaku benar
            'hal ini sudah di handle pada coding dgv

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            If Me.chkUseWarehouseExclusion.Checked = True Then
                Me.dgvWarehouse.AllowUserToAddRows = True
            End If
        End Try
    End Function

    Private Sub SimpanDataIncludeExcludeWarehouse()
        Dim DTA As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_setting_warehouse_exclusionTableAdapter
        Dim Inc_Exc As EnumExcludeInclude = EnumExcludeInclude.Include
        Dim Level_R As EnumLocationLevel = EnumLocationLevel.Location
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkUseWarehouseExclusion.Checked = True Then
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_UseWarehouseExclusion, _
                "1", Me.UserCode)

                'simpan detail itemnya
                DTA.Delete(Me.UserCode)
                Me.dgvWarehouse.AllowUserToAddRows = False
                For Each rowDGV In Me.dgvWarehouse.Rows
                    If rowDGV.Cells(Me.ColWarehouse_ExcludeInclude.Index).Value = "Include" Then
                        Inc_Exc = EnumExcludeInclude.Include
                    Else
                        Inc_Exc = EnumExcludeInclude.Exclude
                    End If
                    If rowDGV.Cells(Me.ColWarehouse_Level.Index).Value = "Warehouse" Then
                        Level_R = EnumLocationLevel.Warehouse
                    Else
                        Level_R = EnumLocationLevel.Location
                    End If

                    DTA.Insert(rowDGV.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value, CInt(Level_R), _
                    rowDGV.Cells(Me.ColWarehouse_Remarks.Index).Value, CInt(Inc_Exc), Me.UserCode)
                Next
                Me.dgvWarehouse.AllowUserToAddRows = True
            Else
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_UseWarehouseExclusion, _
                "0", Me.UserCode)
                DTA.Delete(Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub dgvWarehouse_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvWarehouse.DataError

    End Sub

    Private Sub dgvWarehouse_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvWarehouse.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvWarehouse.Rows(Me.dgvWarehouse.Rows.Count - 1)

            rowDGV.Cells(Me.ColWarehouse_Code.Index).Value = ""
            rowDGV.Cells(Me.ColWarehouse_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColWarehouse_ExcludeInclude.Index).Value = "Include"
            rowDGV.Cells(Me.ColWarehouse_Level.Index).Value = "Warehouse"
            rowDGV.Cells(Me.ColWarehouse_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColWarehouse_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColWarehouse_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
            If DummyIndex <> -1 Then
                Me.dgvWarehouse.Rows.RemoveAt(DummyIndex)
                DummyIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWarehouse_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWarehouse.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvWarehouse)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvWarehouse_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvWarehouse.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvWarehouse)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class