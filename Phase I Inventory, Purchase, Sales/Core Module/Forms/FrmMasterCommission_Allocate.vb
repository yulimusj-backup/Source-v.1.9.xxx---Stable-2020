Option Explicit On
Option Strict Off

Partial Public Class FrmMasterCommission

    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim DummyIndex As Integer = -1
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1

    Private Sub chkSalesPeriod_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSalesPeriod.Click
        Try
            If Me.chkSalesPeriod.Checked = True Then
                Me.dtpFrom.Enabled = True
                Me.dtpTo.Enabled = True
            ElseIf Me.chkSalesPeriod.Checked = False Then
                Me.dtpFrom.Enabled = False
                Me.dtpTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCommission_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCommission.Click
        Try
            If Me.chkCommission.Checked = True Then
                Me.dgvFilterComm.Rows.Clear()
                Me.dgvFilterComm.ReadOnly = False
                Me.dgvFilterComm.AllowUserToAddRows = True
            ElseIf Me.chkCommission.Checked = False Then
                Me.dgvFilterComm.Rows.Clear()
                Me.dgvFilterComm.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkProduct.Click
        Try
            If Me.chkProduct.Checked = True Then
                Me.dgvItemProductGroup.Rows.Clear()
                Me.dgvItemProductGroup.ReadOnly = False
                Me.dgvItemProductGroup.AllowUserToAddRows = True
            ElseIf Me.chkProduct.Checked = False Then
                Me.dgvItemProductGroup.Rows.Clear()
                Me.dgvItemProductGroup.ReadOnly = True
                Me.dgvItemProductGroup.AllowUserToAddRows = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilterProduct_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
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
                            Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwProduct_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwProduct.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwProduct_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwProduct.RunWorkerCompleted
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

    Private Sub dgvFilterProduct_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As ClsMasterCommission_Allocate.ProductLevel = ClsMasterCommission_Allocate.ProductLevel.ProductGroup

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.bgwProduct.RunWorkerAsync()
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
                    Level = Me.ObjAllocate.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = ClsMasterCommission_Allocate.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = ClsMasterCommission_Allocate.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = ClsMasterCommission_Allocate.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = ClsMasterCommission_Allocate.ProductLevel.Category Then
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
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Trim(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.bgwProduct.RunWorkerAsync()
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
                    Me.bgwProduct.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.Buffer, vProductGroup_Level)
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.bgwProduct.RunWorkerAsync()
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
                    Me.bgwProduct.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.bgwProduct.RunWorkerAsync()
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
                    Me.bgwProduct.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
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

    Private Sub dgvFilterProduct_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.ObjAllocate.GetProductLevelStr(ClsMasterCommission_Allocate.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilterProduct_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub

    Private Sub dgvFilterComm_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvFilterComm.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColFilter_CommisionCode.Index Then
                Me.Buffer = Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvFilterComm_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFilterComm.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvFilterComm.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim Primary As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColFilter_CommisionCode.Index Then
                        If Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value = Me.Buffer Then
                            Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If

                        Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjType.ApakahSudahAda(Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Primary = Me.ObjType.GetPrimary(Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.getindexitem(.Cells(Me.ColFilter_CommisionCode.Index).Value, e.RowIndex)
                            If index < 0 Then
                                Me.dgvFilterComm.Item(Me.ColFilter_PrimaryCommission.Index, e.RowIndex).Value = Me.ObjType.GetPrimary(Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value = Me.ObjType.GetCode(Me.dgvFilterComm.Item(Me.ColFilter_PrimaryCommission.Index, e.RowIndex).Value)
                                Me.dgvFilterComm.Item(Me.ColFilter_CommissionDesc.Index, e.RowIndex).Value = Me.ObjType.GetDesc(Me.dgvFilterComm.Item(Me.ColFilter_PrimaryCommission.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.bgwAllocate.RunWorkerAsync()
                                    Else
                                        Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvFilterComm )
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New FrmMasterCommission_Cari
                            ij.txtCari.Text = Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Primary <> -1 Then
                                index = Me.getindexitem(ij.Primary, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value = Me.ObjType.GetCode(ij.Primary)
                                    Me.dgvFilterComm.Item(Me.ColFilter_CommissionDesc.Index, e.RowIndex).Value = Me.ObjType.GetDesc(ij.Primary)
                                    Me.dgvFilterComm.Item(Me.ColFilter_PrimaryCommission.Index, e.RowIndex).Value = ij.Primary
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.bgwAllocate.RunWorkerAsync()
                                        Else
                                            Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        'Me.ObjCD.SetNumber(Me.dgvFilterComm )
                                        Exit Sub
                                    Else
                                        Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value = Me.ObjType.GetCode(ij.Primary)
                                        Me.dgvFilterComm.Item(Me.ColFilter_CommissionDesc.Index, e.RowIndex).Value = Me.ObjType.GetDesc(ij.Primary)
                                        Me.dgvFilterComm.Item(Me.ColFilter_PrimaryCommission.Index, e.RowIndex).Value = ij.Primary
                                    End If
                                End If
                            ElseIf ij.Primary = -1 Then
                                If Me.dgvFilterComm.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.bgwAllocate.RunWorkerAsync()
                                    Else
                                        Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvFilterComm, False)
                                    Exit Sub
                                End If
                                If Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.bgwAllocate.RunWorkerAsync()
                                    Else
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    'Me.ObjCD.SetNumber(Me.dgvItem)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.dgvFilterComm)
                    End If
                Else
                    Me.dgvFilterComm.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getindexitem(ByVal vCode As String, _
  ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvFilterComm.Rows.Count - 1
                RowDGV = Me.dgvFilterComm.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColFilter_CommisionCode.Index).Value = vCode Then
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

    Private Function getindexitemAllocate(ByVal vCode As String, _
 ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvAllocate.Rows.Count - 1
                RowDGV = Me.dgvAllocate.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColAllocate_CommisionnCode.Index).Value = vCode Then
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

    Private Sub bgwAllocate_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwAllocate.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwAllocate_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwAllocate.RunWorkerCompleted
        Try
            Me.dgvFilterComm.Rows.RemoveAt(EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahFilteranSudahBenar() As Boolean
        Try

            If Me.chkSalesPeriod.Checked = True Then
                If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                    MsgBox("Please choose sales period with correctly.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            If Me.chkCommission.Checked = True Then
                If Me.dgvFilterComm.Rows.Count = 1 Then
                    If String.IsNullOrEmpty(Me.dgvFilterComm.Item(Me.ColFilter_CommisionCode.Index, 0).Value) = True Then
                        MsgBox("Please enter item for data commission if filter cheklist", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            If Me.chkProduct.Checked = True Then
                If Me.dgvItemProductGroup.Rows.Count = 1 Then
                    If String.IsNullOrEmpty(Me.dgvItemProductGroup.Item(Me.ColProductGroup_CodeNo.Index, 0).Value) = True Then
                        MsgBox("Please enter item for data product level if filter cheklist", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If ApakahFilteranSudahBenar() = True Then
                Me.ObjAllocate.LoadData_ToDGV(Me.dtpFrom.Value, Me.dtpTo.Value, Me, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAllocate_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAllocate.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColAllocate_CommisionnCode.Index Or e.ColumnIndex = Me.ColAllocate_Set.Index Then
                If e.ColumnIndex = Me.ColAllocate_CommisionnCode.Index Then
                    Me.Buffer = Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value
                ElseIf e.ColumnIndex = Me.ColAllocate_Set.Index Then
                    Me.Buffer = CStr(Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAllocate_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAllocate.CellEndEdit
        Dim vRow As DataGridViewRow = Me.dgvAllocate.Rows(e.RowIndex)
        Dim vRowIdx As Integer = -1
        Dim vPrimary As Double = -1
        Try
            With vRow
                If IsNothing(Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColAllocate_CommisionnCode.Index Then
                        If Me.Buffer = Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value Then
                            Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = Me.ColAllocate_Set.Index Then
                        If Me.Buffer = CStr(Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value = CBool(Me.Buffer)
                            Exit Sub
                        End If
                    End If

                    If e.ColumnIndex = Me.ColAllocate_CommisionnCode.Index Then
                        Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If ObjType.ApakahSudahAda(.Cells(Me.ColAllocate_CommisionnCode.Index).Value) = True Then
                            vPrimary = Me.ObjType.GetPrimary(Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvAllocate.Item(Me.ColAllocate_PrimaryCommision.Index, e.RowIndex).Value = vPrimary
                            Me.dgvAllocate.Item(Me.ColAllocate_CommisionnCode.Index, e.RowIndex).Value = Me.ObjType.GetCode(vPrimary)
                            Me.ObjAllocate.SimpanItem(e.RowIndex, Me)
                        Else
                            Dim ij As New Core_Module.FrmMasterCommission_Cari
                            ij.txtCari.Text = Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Primary <> -1 Then
                                Me.dgvAllocate.Item(Me.ColAllocate_PrimaryCommision.Index, e.RowIndex).Value = ij.Primary
                                Me.dgvAllocate.Item(Me.ColAllocate_CommisionnCode.Index, e.RowIndex).Value = Me.ObjType.GetCode(ij.Primary)
                                Me.ObjAllocate.SimpanItem(e.RowIndex, Me)

                            ElseIf ij.Primary = -1 Then
                                If Me.dgvAllocate.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.dgvAllocate.Item(Me.ColAllocate_PrimaryCommision.Index, e.RowIndex).Value = -1
                                        Me.dgvAllocate.Item(Me.ColAllocate_CommisionnCode.Index, e.RowIndex).Value = ""
                                    Else
                                        Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvAllocate, False)
                                    Exit Sub
                                End If
                                If Me.dgvAllocate.Item(Me.ColAllocate_CommisionnCode.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.dgvAllocate.Item(Me.ColAllocate_PrimaryCommision.Index, e.RowIndex).Value = -1
                                        Me.dgvAllocate.Item(Me.ColAllocate_CommisionnCode.Index, e.RowIndex).Value = ""
                                    Else
                                        Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvAllocate, False)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If

                        Me.ObjCD.SetNumber(Me.dgvAllocate, False)
                    Else

                    End If
                Else
                    If e.ColumnIndex = Me.ColAllocate_CommisionnCode.Index Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.dgvAllocate.Item(Me.ColAllocate_PrimaryCommision.Index, e.RowIndex).Value = -1
                            Me.dgvAllocate.Item(Me.ColAllocate_CommisionnCode.Index, e.RowIndex).Value = ""
                        Else
                            Me.dgvAllocate.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        Me.ObjCD.SetNumber(Me.dgvAllocate, False)
                        Exit Sub
                    End If
                    Exit Sub
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSet.Click
        Dim hj As New FrmMasterCommission_Cari
        Dim OK As Boolean = False
        Try
            For Each row As DataGridViewRow In Me.dgvAllocate.Rows
                If row.Cells(Me.ColAllocate_Set.Index).Value = True Then
                    OK = True
                End If
            Next

            If OK = True Then
                hj.BackColorAll = Me.BackColorAll
                hj.ShowDialog()

                If hj.Primary <> -1 Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.ObjAllocate.GenerateQuery(Me.dgvAllocate, Me, hj.Primary)
                    Me.ObjAllocate.IsiDGV(Me.dgvAllocate, Me)
                    Windows.Forms.Cursor.Current = Cursors.Default
                End If
            Else
                MsgBox("Please check at least one item.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
          
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub llbCheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbCheckAll.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvAllocate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAllocate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each RowDgv As DataGridViewRow In Me.dgvAllocate.Rows
                With RowDgv
                    .Cells(Me.ColAllocate_Set.Index).Value = True
                End With
            Next
            Me.dgvAllocate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAllocate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbUncheckAll.Click
        Try
            Me.dgvAllocate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAllocate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each RowDgv As DataGridViewRow In Me.dgvAllocate.Rows
                With RowDgv
                    .Cells(Me.ColAllocate_Set.Index).Value = False
                End With
            Next
            Me.dgvAllocate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAllocate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAllocate_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAllocate.Sorted
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ObjCD.SetNumber(Me.dgvAllocate, False)
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
