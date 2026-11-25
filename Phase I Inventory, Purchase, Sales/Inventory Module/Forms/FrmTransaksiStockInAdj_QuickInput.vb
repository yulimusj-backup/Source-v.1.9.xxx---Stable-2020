Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiStockInAdj_QuickInput
    Dim ObjCD As New Core_Module.ClsCekData
    Dim buffer As String = ""
    Dim DummyIndex As Integer = -1
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1
    Dim objItem As New ClsTransaksiMaterialPlanningGroupItem
    Dim ObjProduct As New clsMasterProduct
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public DefQty As Decimal = 0

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.chkCoverMinus.Checked = False Then
                Me.dgvItemProductGroup.AllowUserToAddRows = False
                If Me.dgvItemProductGroup.RowCount > 0 Then
                    'pastikan qty > 0
                    If CDec(Me.TxtQty.Text) <= 0 And Me.chkImportFromDB.Checked = False And Me.ChkMirror.Checked = False Then
                        MsgBox("Please specify a positive number for qty data.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    If MsgBox("Close form and start quick input?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Else
                    MsgBox("Please specify at least 1 product group.", MsgBoxStyle.Information, "Information")
                End If
            Else
                If Me.ChkMirror.Checked = True Or Me.chkImportFromDB.Checked = True Then
                    MsgBox("Please make sure your option", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    'pastikan limit baris > 0
                    If Me.TxtCoverMinus_LimitData.Text <= 0 Then
                        MsgBox("Please specify a valid data for data limit", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    'pastikan tanggal yang dipilih > smess covertion date -> di pass aj, asumsi ini kita sendiri yang bikin
                    If MsgBox("Close form and start quick input?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvItemProductGroup.AllowUserToAddRows = True
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel quick input?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
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

    Private Sub TxtQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQty, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.GotFocus
        Try
            Me.TxtQty.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQty, sender, e, 10 ^ 9)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQty, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.Leave
        Try
            Me.TxtQty.Text = Me.ObjCD.DigitGrouping(Math.Round(CDec(Me.ObjCD.BuangNolDibelakangDecimal(Me.TxtQty.Text)), 6, MidpointRounding.AwayFromZero))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty.LostFocus
        Try
            Me.TxtQty.BackColor = Me.TxtRef.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInAdj_QuickInput_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.TxtQty.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.BuangNolDibelakangDecimal(Me.DefQty.ToString))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkImportFromDB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkImportFromDB.CheckedChanged
        Try
            If Me.chkImportFromDB.Checked = True Then
                Me.TxtQty.Text = 0
                Me.ChkMirror.Checked = False
            Else
                If Me.ChkMirror.Checked = True Or Me.chkImportFromDB.Checked = True Then
                    Me.TxtQty.Text = 0
                Else
                    Me.TxtQty.Text = Me.ObjCD.DigitGrouping(1000000)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkMirror_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkMirror.CheckedChanged
        Try
            If Me.ChkMirror.Checked = True Then
                Me.TxtQty.Text = 0
                Me.chkImportFromDB.Checked = False
            Else
                If Me.ChkMirror.Checked = True Or Me.chkImportFromDB.Checked = True Then
                    Me.TxtQty.Text = 0
                Else
                    Me.TxtQty.Text = Me.ObjCD.DigitGrouping(1000000)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCoverMinus_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCoverMinus.CheckedChanged
        Try
            If Me.chkCoverMinus.Checked = True Then
                Me.dtpCoverMinus_RefDate.Enabled = True
                Me.TxtCoverMinus_LimitData.Enabled = True
            Else
                Me.dtpCoverMinus_RefDate.Enabled = False
                Me.TxtCoverMinus_LimitData.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCoverMinus_LimitData_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCoverMinus_LimitData.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtCoverMinus_LimitData, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCoverMinus_LimitData_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCoverMinus_LimitData.GotFocus
        Try
            If Me.TxtCoverMinus_LimitData.Enabled = True Then
                Me.TxtCoverMinus_LimitData.BackColor = Me.BackColorAll
            Else
                Me.TxtCoverMinus_LimitData.BackColor = Me.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCoverMinus_LimitData_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCoverMinus_LimitData.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtCoverMinus_LimitData, sender, e, 5000)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCoverMinus_LimitData_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCoverMinus_LimitData.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtCoverMinus_LimitData, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCoverMinus_LimitData_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCoverMinus_LimitData.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtCoverMinus_LimitData, sender, e, 0, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCoverMinus_LimitData_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCoverMinus_LimitData.LostFocus
        Try
            If Me.TxtCoverMinus_LimitData.Enabled = True Then
                Me.TxtCoverMinus_LimitData.BackColor = Me.TextBox1.BackColor
            Else
                Me.TxtCoverMinus_LimitData.BackColor = Me.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class