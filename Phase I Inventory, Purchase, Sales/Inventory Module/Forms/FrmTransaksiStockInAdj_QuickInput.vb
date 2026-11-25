Option Explicit On
Option Strict Off

Imports Core_Module
Imports System.IO

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

            '2025-09-02 IM, untuk keeprluan WDI
            Me.dgvImport.AllowUserToAddRows = False
            Me.dgvImport.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.FormatDGV()
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

    '2025-09-02 IM, coding Tab IMPORT
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Dim hasil As MsgBoxResult
        If Me.dgvImport.Rows.Count > 0 Then
            hasil = MsgBox("Cancel save and discard data?", MsgBoxStyle.YesNo, "Confirm")
            If hasil = MsgBoxResult.Yes Then
                Me.Close()
            End If
        Else
            Me.Close()
        End If
    End Sub

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click
        Dim ofd As New OpenFileDialog()
        Try
            ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            ofd.Title = "Pilih File CSV"

            If ofd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim csvPath As String = ofd.FileName
                ' Tampilkan nama file ke TextBox (full path)
                txtFilename.Text = csvPath
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Try
            Me.dgvImport.Rows.Clear()
            'If Me.chkShowErrorOnly.Checked Then
            '    loadDataDummyErrorOnly()
            'Else
            '    Me.loadDataDummy()
            'End If
            Me.ObjCD.UpdateThreading()
            If Me.dgvImport.Rows.Count > 0 AndAlso _
            MsgBox("There is already item exist, are you sure to replace it?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            Me.dgvImport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvImport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvImport.AllowUserToAddRows = False

            Me.loadCSV(Me.txtFilename.Text)

            'cek jumlah hasil load
            If Me.dgvImport.RowCount = 0 Then
                MsgBox("No data found.", MsgBoxStyle.Information, "Inforation")
                Me.txtFilename.Text = ""
                Exit Sub
            End If

            If Me.chkShowErrorOnly.Checked Then
                Me.removeNonInvalid()
            End If

            Me.FormatDGV()

            If Me.ChkShowInvalidOnly = False Then
                MsgBox("Data load completed.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Data load completed. Some invalid data was found.", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            Me.ObjCD.SetNumber(Me.dgvImport)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                'Me.SalinKeObjek()
                hasil = MsgBox("Import Data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    'If Me.ObjAsset.SimpanPerubahan(Me.UserCode) Then
                    MsgBox("Data Imported", MsgBoxStyle.Information, "Information")
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    '    Me.SalinKeField()
                    Me.Close()
                    '    Me.btnEdit.Focus()
                    'End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim rowdgv As DataGridViewRow
        Try
            If Me.dgvImport.Rows.Count = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            For Each rowdgv In Me.dgvImport.Rows
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_CodeNo.Index).Value) Then
                    MsgBox("Please enter valid product code.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                If CDec(rowdgv.Cells(Me.Col_Import_Qty.Index).Value) < 0 Or String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Qty.Index).Value) Then
                    MsgBox("Qty must be positive number or zero.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_UnitCost.Index).Value) Then
                    MsgBox("Please enter valid unit code.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If CDec(rowdgv.Cells(Me.Col_Import_Amount.Index).Value) < 0 Or String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Amount.Index).Value) Then
                    MsgBox("Amount must be positive number or zero.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If CDec(rowdgv.Cells(Me.col_Import_Qty.Index).Value) > 0 AndAlso _
                CDec(rowdgv.Cells(Me.col_Import_Amount.Index).Value) / CDec(rowdgv.Cells(Me.col_Import_Qty.Index).Value) <= 0 Then
                    MsgBox("Unit cost must .", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) = False Then
                    MsgBox("Data invalid cannot be imported.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub loadDataDummy()
    '    Dim rowdgv As DataGridViewRow
    '    Try
    '        Me.dgvImport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        Me.dgvImport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

    '        'T100 X 2.5
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "T100 X 2.5"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "T100 x 2.5"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "4,500.00"
    '            .Cells(Me.col_Import_Qty.Index).Value = "5,000"
    '            .Cells(Me.col_Import_Amount.Index).Value = "22,500,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'T100 X 3.5
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "T100 X 3.5"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "T100 x 3.5"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "4,500.00"
    '            .Cells(Me.col_Import_Qty.Index).Value = "5,000"
    '            .Cells(Me.col_Import_Amount.Index).Value = "22,500,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'T355 X 3
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "T355 X 3"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "T355 x 3"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "4,500.00"
    '            .Cells(Me.col_Import_Qty.Index).Value = "5,000"
    '            .Cells(Me.col_Import_Amount.Index).Value = "22,500,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'ALUMINIUM OKSIDA
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "ALUMINIUM OKSIDA"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Aluminium Oksida"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "600.00"
    '            .Cells(Me.col_Import_Qty.Index).Value = "125,000"
    '            .Cells(Me.col_Import_Amount.Index).Value = "75,000,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'RESIN LIQUID	
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "RESIN LIQUID"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Resin Liquid"
    '            .Cells(Me.col_Import_Qty.Index).Value = "120,000"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "1,000.00"
    '            .Cells(Me.col_Import_Amount.Index).Value = "120,000,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'SILIKON KARBIDA
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "SILIKON KARBIDA"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Silikon Karbida"
    '            .Cells(Me.col_Import_Qty.Index).Value = "212,111"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "700.00"
    '            .Cells(Me.col_Import_Amount.Index).Value = "148,477,700.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'RING	
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "RING"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Ring"
    '            .Cells(Me.col_Import_Qty.Index).Value = "250,000"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "250.00"
    '            .Cells(Me.col_Import_Amount.Index).Value = "62,500,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'RING LABEL	
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Ready"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "RING LABEL"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Ring Label"
    '            .Cells(Me.col_Import_Qty.Index).Value = "250,000"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "300.00"
    '            .Cells(Me.col_Import_Amount.Index).Value = "75,000,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = ""
    '        End With

    '        'CERAMIC BOND
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Error"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "CERAMIC BOND"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Ceramic Bond"
    '            .Cells(Me.col_Import_Qty.Index).Value = "10,000"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "0"
    '            .Cells(Me.col_Import_Amount.Index).Value = "62,500,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = "Unit Cost must not be empty or zero"
    '        End With
    '        rowdgv.DefaultCellStyle.BackColor = Color.PaleVioletRed

    '        'RESINOID BOND
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Error"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "RESINOID BOND"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Resinoid Bond"
    '            .Cells(Me.col_Import_Qty.Index).Value = "-10,000"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "200.00"
    '            .Cells(Me.col_Import_Amount.Index).Value = "-2,000,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = "Quantity must be a positive number, Amount must be positive number"
    '        End With
    '        rowdgv.DefaultCellStyle.BackColor = Color.PaleVioletRed


    '        Me.dgvImport.Rows(0).Selected = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Me.dgvImport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        Me.dgvImport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '        Me.ObjCD.SetNumber(Me.dgvImport, True)
    '    End Try
    'End Sub

    'Private Sub loadDataDummyErrorOnly()
    '    Dim rowdgv As DataGridViewRow
    '    Try
    '        Me.dgvImport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        Me.dgvImport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

    '        'CERAMIC BOND
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Error"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "CERAMIC BOND"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Ceramic Bond"
    '            .Cells(Me.col_Import_Qty.Index).Value = "10,000"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "0"
    '            .Cells(Me.col_Import_Amount.Index).Value = "62,500,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = "Unit Cost must not be empty or zero"
    '        End With
    '        rowdgv.DefaultCellStyle.BackColor = Color.PaleVioletRed

    '        'RESINOID BOND
    '        Me.dgvImport.Rows.Add()
    '        rowdgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Count - 1)
    '        With rowdgv
    '            .Cells(Me.col_Import_Status.Index).Value = "Error"
    '            .Cells(Me.col_Import_ProductCode.Index).Value = "RESINOID BOND"
    '            .Cells(Me.col_Import_ProductDesc.Index).Value = "Resinoid Bond"
    '            .Cells(Me.col_Import_Qty.Index).Value = "-10,000"
    '            .Cells(Me.col_Import_UnitCost.Index).Value = "200.00"
    '            .Cells(Me.col_Import_Amount.Index).Value = "-2,000,000.00"
    '            .Cells(Me.col_Import_Remarks.Index).Value = "Quantity must be a positive number, Amount must be positive number"
    '        End With
    '        rowdgv.DefaultCellStyle.BackColor = Color.PaleVioletRed


    '        Me.dgvImport.Rows(0).Selected = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Me.dgvImport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        Me.dgvImport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '        Me.ObjCD.SetNumber(Me.dgvImport, True)
    '    End Try
    'End Sub

    '2025-12-09 IM, import karyawan ke tbl filter
    Private Enum EnumColumnImportProduct As Integer
        colProductCode
        colDesc
        colQty
        colUnitCode
        colBatchNo
        colExpiryDate
        colLocation
        colRemarks
        colUnitCost
        colAmount
    End Enum
    Dim ChkShowInvalidOnly As Boolean = False
    Dim Separator As String = ";"

    Private Sub loadCSV(ByVal vFileName As String)
        Dim StreamReader_ As System.IO.StreamReader
        Dim CSV_Line As String = ""
        Dim RowDgv As DataGridViewRow
        Dim IndexKetemu As Integer = -1
        Dim rDataGrid As DataGridViewRow
        Dim Idx As Integer = -1
        Dim ListIDcardYgKosong As New List(Of Integer)
        Dim ProductDataTable As New dsLaporanStockInAdjust.TblMasterProductDataTable
        Dim dta As New dsLaporanStockInAdjustTableAdapters.TblMasterProductTableAdapter
        Dim ImportData As New dsLaporanStockInAdjust.TblImportProductDataTable
        Dim unitCost As Decimal = 0
        Dim defprec As Integer = 0
        Dim defprecUnit As Integer = 0
        Dim primaryProduct As Double = -1
        Dim primaryUnit As Double = -1
        Dim dtaLocation As New dsLaporanStockInAdjustTableAdapters.TblLocationTableAdapter
        Dim TblLocation As New dsLaporanStockInAdjust.TblLocationDataTable
        Try
            StreamReader_ = New System.IO.StreamReader(vFileName)
            CSV_Line = StreamReader_.ReadLine
            Me.dgvImport.Rows.Clear()

            'dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill_ALL(ProductDataTable)
            dtaLocation.Fill(TblLocation)

            Do
                CSV_Line = StreamReader_.ReadLine

                If CSV_Line = String.Empty Then
                    Exit Do 'ketemu baris kosong, keluar dari loop
                End If

                Dim CSV_Column() As String = CSV_Line.Split(Me.Separator)
                If String.IsNullOrEmpty(CSV_Column(EnumColumnImportProduct.colProductCode).Trim.ToUpper.Replace("'", "")) = True Then
                    Exit Do 'klo sudah baris terakhir maka keluar dari loop
                End If

                RowDgv = Me.dgvImport.Rows(Me.dgvImport.Rows.Add)
                Idx = RowDgv.Index

                RowDgv.Cells(Me.Col_Import_CodeNo.Index).Value = CSV_Column(EnumColumnImportProduct.colProductCode).Trim.ToUpper.Replace("'", "")
                primaryProduct = Me.ObjProduct.GetPrimary(RowDgv.Cells(Me.Col_Import_CodeNo.Index).Value)
                If String.IsNullOrEmpty(RowDgv.Cells(Me.Col_Import_CodeNo.Index).Value) Then
                    Me.dgvImport.Rows.Remove(RowDgv)
                Else
                    If primaryProduct <> -1 Then
                        defprec = Me.ObjProduct.GetAngkaDecimal(primaryProduct)
                        primaryUnit = Me.ObjProduct.GetPrimaryUnit(primaryProduct)
                        defprecUnit = Me.ObjProduct.GetAngkaDecimal(primaryProduct, primaryUnit)
                    End If
                    RowDgv.Cells(Me.Col_Import_Description.Index).Value = CSV_Column(EnumColumnImportProduct.colDesc).Trim.Replace("'", "")
                    If Me.ObjCD.CekNumerik(CSV_Column(EnumColumnImportProduct.colQty).ToString.Replace("'", "").Replace("""", "")) Then
                        RowDgv.Cells(Me.Col_Import_Qty.Index).Value = Me.ObjCD.DigitGrouping(CSV_Column(EnumColumnImportProduct.colQty).ToString.Replace("'", "").Replace("""", ""), defprecUnit)
                    Else
                        RowDgv.Cells(Me.Col_Import_Qty.Index).Value = CSV_Column(EnumColumnImportProduct.colQty)
                    End If
                    RowDgv.Cells(Me.Col_Import_Unit.Index).Value = CSV_Column(EnumColumnImportProduct.colUnitCode).Trim.ToUpper.Replace("'", "")
                    RowDgv.Cells(Me.Col_Import_Batch.Index).Value = CSV_Column(EnumColumnImportProduct.colBatchNo).Trim.ToUpper.Replace("'", "")
                    RowDgv.Cells(Me.Col_Import_ExpiredDate.Index).Value = CSV_Column(EnumColumnImportProduct.colExpiryDate).Trim.Replace("'", "")
                    RowDgv.Cells(Me.Col_Import_Location.Index).Value = CSV_Column(EnumColumnImportProduct.colLocation).Trim.ToUpper.Replace("'", "")
                    If Me.ObjCD.CekNumerik(CSV_Column(EnumColumnImportProduct.colUnitCost).ToString.Replace("'", "").Replace("""", "")) Then
                        RowDgv.Cells(Me.Col_Import_UnitCost.Index).Value = Me.ObjCD.DigitGrouping(CSV_Column(EnumColumnImportProduct.colUnitCost).ToString.Replace("'", "").Replace("""", ""), 2)
                    Else
                        RowDgv.Cells(Me.Col_Import_UnitCost.Index).Value = CSV_Column(EnumColumnImportProduct.colUnitCost)
                    End If
                    If Me.ObjCD.CekNumerik(CSV_Column(EnumColumnImportProduct.colAmount).ToString.Replace("'", "").Replace("""", "")) Then
                        RowDgv.Cells(Me.Col_Import_Amount.Index).Value = Me.ObjCD.DigitGrouping(CSV_Column(EnumColumnImportProduct.colAmount).ToString.Replace("'", "").Replace("""", ""), 2)
                    Else
                        RowDgv.Cells(Me.Col_Import_Amount.Index).Value = CSV_Column(EnumColumnImportProduct.colAmount)
                    End If
                    RowDgv.Cells(Me.col_Import_Remarks.Index).Value = CSV_Column(EnumColumnImportProduct.colRemarks).Trim.Replace("'", "")
                    RowDgv.Cells(Me.Col_Import_Invalid.Index).Value = ""

                    'nilai selalu -1
                    RowDgv.Cells(Me.Col_Import_PrimaryPdO.Index).Value = -1
                    RowDgv.Cells(Me.Col_Import_PrimaryPOno.Index).Value = -1
                    RowDgv.Cells(Me.Col_Import_PrimarySecondaryUnits.Index).Value = -1
                    RowDgv.Cells(Me.Col_Import_SecondaryQty.Index).Value = -1

                    'nilai selalu kosong
                    RowDgv.Cells(Me.Col_Import_SuppDONo.Index).Value = ""
                    RowDgv.Cells(Me.Col_Import_PONo.Index).Value = ""
                    RowDgv.Cells(Me.Col_Import_SecondaryUnit.Index).Value = ""

                    'Nilai Sementara
                    RowDgv.Cells(Me.Col_Import_WareHouse.Index).Value = ""
                    RowDgv.Cells(Me.Col_Import_PrimaryCodeNo.Index).Value = -1
                    RowDgv.Cells(Me.Col_Import_PrimaryLocation.Index).Value = -1
                    RowDgv.Cells(Me.Col_Import_PrimaryUnits.Index).Value = -1
                End If
                'Me.FillProductData(RowDgv.Cells(Me.col_Import_ProductCode.Index).Value, RowDgv, ProductDataTable, ImportData)
            Loop

            'Cek nilai Invalid
            Me.checkInvalidData(ProductDataTable, TblLocation)
            'Me.dgvImport.Sort(Me.Col_Import_CodeNo, System.ComponentModel.ListSortDirection.Ascending)

        Catch ex As Exception
            Me.dgvImport.Rows.Clear()
            If ex.Message.ToLower.Contains("because it is being used by another process") Then
                MsgBox("Please close the CSV file that opened by another program.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Something wrong while load CSV file: " & vbNewLine & ex.Message & vbNewLine & "Please contact your software provider.", MsgBoxStyle.Information, "Information")
            End If
        Finally
            If IsNothing(StreamReader_) = False Then
                StreamReader_.Close()
                StreamReader_.Dispose()
            End If
            dta.Dispose()
        End Try

    End Sub

    Private Sub checkInvalidData(ByVal ProductDataTable As dsLaporanStockInAdjust.TblMasterProductDataTable, _
    ByVal TblLocation As dsLaporanStockInAdjust.TblLocationDataTable)
        Dim rowdgv As DataGridViewRow
        Try

            For Each rowdgv In Me.dgvImport.Rows
                'cek product data di db
                Me.CekProductData(rowdgv.Cells(Me.Col_Import_CodeNo.Index).Value, rowdgv, ProductDataTable)
                If rowdgv.Cells(Me.Col_Import_PrimaryCodeNo.Index).Value <> -1 Then
                    Me.cekLocationData(CDbl(rowdgv.Cells(Me.Col_Import_PrimaryCodeNo.Index).Value), _
                    rowdgv.Cells(Me.Col_Import_Location.Index).Value, rowdgv, TblLocation)
                    Me.cekValiditasData(CDbl(rowdgv.Cells(Me.Col_Import_PrimaryCodeNo.Index).Value), rowdgv, ProductDataTable)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CekProductData(ByVal vProductCode As String, ByVal rowdgv As DataGridViewRow, _
    ByVal ProductDataTable As dsLaporanStockInAdjust.TblMasterProductDataTable)
        Dim row As dsLaporanStockInAdjust.TblMasterProductRow
        Dim ketemuUnit As Boolean = False
        Dim ketemuSecUnit As Boolean = False
        Try
            ProductDataTable.DefaultView.RowFilter = "[Product Code] = '" + vProductCode.Trim.ToUpper + "'"
            If ProductDataTable.DefaultView.Count > 0 Then
                row = ProductDataTable.DefaultView(0).Row

                'cek product Goods atau bukan
                If row.Type = 1 Then
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Only goods products can be imported"
                    Else
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Only goods products can be imported"
                    End If
                    Exit Sub
                End If

                'fill primary data
                rowdgv.Cells(Me.Col_Import_PrimaryCodeNo.Index).Value = row.Primary
                'rowdgv.Cells(Me.Col_Import_Description.Index).Value = row.Description
                'cek product Desc
                If row.Description <> rowdgv.Cells(Me.Col_Import_Description.Index).Value Then
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Product description does not match the actual product"
                    Else
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Product description does not match the actual product"
                    End If
                End If

                'cari primary unit
                For Each rowdv As Data.DataRowView In ProductDataTable.DefaultView
                    row = rowdv.Row
                    If row.Unit_Code.ToUpper = rowdgv.Cells(Me.Col_Import_Unit.Index).Value.ToUpper Then
                        rowdgv.Cells(Me.Col_Import_PrimaryUnits.Index).Value = row.Primary_Unit
                        ketemuUnit = True
                        Exit For
                    End If
                Next

                ''cari primary Secondary unit (+ dari unit multi)
                'For Each rowdv As Data.DataRowView In ProductDataTable.DefaultView
                '    row = rowdv.Row
                '    If row.Unit_Code = rowdgv.Cells(Me.Col_Import_SecondaryUnit.Index).Value Then
                '        rowdgv.Cells(Me.Col_Import_PrimarySecondaryUnits.Index).Value = row.Primary_Unit
                '        ketemuSecUnit = True
                '        Exit For
                '    End If

                '    If row.Unit_Code_Multi = rowdgv.Cells(Me.Col_Import_SecondaryUnit.Index).Value Then
                '        rowdgv.Cells(Me.Col_Import_PrimarySecondaryUnits.Index).Value = row.Primary_Unit_Multi
                '        ketemuSecUnit = True
                '        Exit For
                '    End If
                'Next
            Else
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Product code Invalid"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Product code invalid"
                End If
                Exit Sub
            End If

            'jika product ada, unit code tidak ada, tambahkan error
            If ketemuUnit = False Then
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Unit code invalid"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Unit code invalid"
                End If
            End If

            ''jika product ada, secondaty unit code tidak ada, tambahkan error
            'If ketemuSecUnit = False Then
            '    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
            '        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Secondary Unit Code Invalid"
            '    Else
            '        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Secondary Unit Code Invalid"
            '    End If
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub cekLocationData(ByVal vPrimaryProduct As Double, ByVal vLocationCode As String, ByVal rowdgv As DataGridViewRow, _
    ByVal TblLocation As dsLaporanStockInAdjust.TblLocationDataTable)
        Dim row As dsLaporanStockInAdjust.TblLocationRow
        Dim LocationControlled As Boolean = Me.ObjProduct.ApakahLocationControlled(vPrimaryProduct)
        Try
            TblLocation.DefaultView.RowFilter = "[Storage Area Code] = '" + vLocationCode.Trim.ToUpper + "'"
            If TblLocation.DefaultView.Count > 0 And LocationControlled = True Then
                row = TblLocation.DefaultView(0).Row
                rowdgv.Cells(Me.Col_Import_PrimaryLocation.Index).Value = row.Primary
                rowdgv.Cells(Me.Col_Import_WareHouse.Index).Value = row.Warehouse_Code
            ElseIf TblLocation.DefaultView.Count > 0 And LocationControlled = False Then
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Product does not use a location"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Product does not use a location"
                End If
            Else
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Location code invalid"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Location code invalid"
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Function IsExistCode(ByVal vCode As String, ByVal vRowIndex As Integer) As Boolean
        Try
            For Each RowDgv As DataGridViewRow In Me.dgvImport.Rows
                If RowDgv.Index <> vRowIndex Then
                    If RowDgv.Cells(Me.Col_Import_CodeNo.Index).Value = vCode Then
                        Return True
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cekValiditasData(ByVal vPrimaryProduct As Double, ByVal rowdgv As DataGridViewRow, _
    ByVal ProductDataTable As dsLaporanStockInAdjust.TblMasterProductDataTable)
        Dim UnitCost As Decimal = 0
        Dim BedaUnitcost As Boolean = False
        Dim defprec As Integer = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct)
        Dim row As dsLaporanStockInAdjust.TblMasterProductRow

        Try
            'Product Tidak boleh double
            If Me.IsExistCode(rowdgv.Cells(Me.Col_Import_CodeNo.Index).Value, rowdgv.Index) Then
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Duplicate Product Code"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Duplicate Product Code"
                End If
            End If

            'Qty < 0
            If Me.ObjCD.CekNumerik(rowdgv.Cells(Me.Col_Import_Qty.Index).Value) Then
                If CDec(rowdgv.Cells(Me.Col_Import_Qty.Index).Value) < 0 Then
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Qty must be a positive number or zero"
                    Else
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Qty must be a positive number or zero"
                    End If
                End If
            Else
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Qty must be a positive number or zero"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Qty must be a positive number or zero"
                End If
            End If

            'Amount < 0
            If Me.ObjCD.CekNumerik(rowdgv.Cells(Me.Col_Import_Amount.Index).Value) Then
                If CDec(rowdgv.Cells(Me.Col_Import_Amount.Index).Value) < 0 Then
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Amount must be a positive number or zero"
                    Else
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Amount must be a positive number or zero"
                    End If
                End If
            Else
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Amount must be a positive number or zero"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Amount must be a positive number or zero"
                End If
            End If

            'cek unit cost = Amount / Qty
            If Me.ObjCD.CekNumerik(rowdgv.Cells(Me.Col_Import_Qty.Index).Value) = True And Me.ObjCD.CekNumerik(rowdgv.Cells(Me.Col_Import_Amount.Index).Value) = True AndAlso _
            CDec(rowdgv.Cells(Me.Col_Import_Qty.Index).Value) > 0 Then
                UnitCost = Me.ObjCD.DigitGrouping(CDec(rowdgv.Cells(Me.Col_Import_Amount.Index).Value) / CDec(rowdgv.Cells(Me.Col_Import_Qty.Index).Value), 2)
            End If

            If Me.ObjCD.CekNumerik(rowdgv.Cells(Me.Col_Import_UnitCost.Index).Value) = True AndAlso _
            Me.ObjCD.DigitGrouping(CDec(rowdgv.Cells(Me.Col_Import_UnitCost.Index).Value), 2) <> UnitCost Then
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Invalid unit cost. The value must be equal to Amount ÷ Qty"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Invalid unit cost. The value must be equal to Amount ÷ Qty"
                End If
                BedaUnitcost = True
            End If

            'cek unit cost <= 0
            If Me.ObjCD.CekNumerik(rowdgv.Cells(Me.Col_Import_UnitCost.Index).Value) = True AndAlso _
            Me.ObjCD.DigitGrouping(CDec(rowdgv.Cells(Me.Col_Import_UnitCost.Index).Value), 2) <= 0 And BedaUnitcost = False Then
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Unit cost must be a positive number"
                Else
                    rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Unit cost must be a positive number"
                End If
            End If

            ProductDataTable.DefaultView.RowFilter = "[Primary]=" & vPrimaryProduct
            If ProductDataTable.DefaultView.Count > 0 Then
                row = ProductDataTable.DefaultView(0).Row
                'ketika product batch controlled, cek kode batch kosong
                If row.Batch_No_Controlled = 1 And String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Batch.Index).Value) = True Then
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Batch number cannot be empty"
                    Else
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Batch number cannot be empty"
                    End If
                ElseIf row.Batch_No_Controlled = 0 And String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Batch.Index).Value) = False Then
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Product does not use a batch number"
                    Else
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Product does not use a batch number"
                    End If
                End If

                'ketika product expiry controlled, cek tanggal
                If row.Expiry_Date_Controlled = True Then
                    'cek kosong
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_ExpiredDate.Index).Value) Then
                        If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                            rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Expiry date cannot be empty"
                        Else
                            rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Expiry date cannot be empty"
                        End If
                    ElseIf IsDate(rowdgv.Cells(Me.Col_Import_ExpiredDate.Index).Value) = False Then
                        If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                            rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "This field must contain a valid date value"
                        Else
                            rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", This field must contain a valid date value"
                        End If
                    End If
                ElseIf row.Expiry_Date_Controlled = False And String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_ExpiredDate.Index).Value) = False Then
                    If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = "Product does not use a expiry date"
                    Else
                        rowdgv.Cells(Me.Col_Import_Invalid.Index).Value = rowdgv.Cells(Me.Col_Import_Invalid.Index).Value & ", Product does not use a expiry date"
                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Function IsValid_dd_MMM_yyyy(ByVal input As String) As Boolean
    '    Try
    '        Dim parts() As String = input.Trim().Split(" "c)

    '        ' Harus 3 bagian: dd MMM yyyy
    '        If parts.Length <> 3 Then Return False

    '        Dim dd As Integer
    '        Dim MMM As String = parts(1).ToUpper()
    '        Dim yyyy As Integer

    '        ' Cek dd (1-31)
    '        If Not Integer.TryParse(parts(0), dd) Then Return False
    '        If dd < 1 Or dd > 31 Then Return False

    '        ' Cek bulan 3 huruf
    '        Dim validMonths As String() = _
    '            {"JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"}

    '        If Array.IndexOf(validMonths, MMM) = -1 Then Return False

    '        ' Cek yyyy (tahun 4 digit)
    '        If Not Integer.TryParse(parts(2), yyyy) Then Return False
    '        If yyyy < 1000 Or yyyy > 9999 Then Return False

    '        ' Cek apakah tanggal valid benar-benar bisa dibaca
    '        Dim finalDate As Date
    '        If Not Date.TryParse(input, finalDate) Then Return False

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return False
    '    End Try
    'End Function

    Public Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim InvalidForeColor As Color = Color.White
        Dim InvalidBackColor As Color = Color.Red

        Try
            Me.dgvImport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'yang diformat dgv-nya:
            For Each row In Me.dgvImport.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.Col_Import_ExpiredDate.Index).Value) = False Then
                    row.Cells(Me.Col_Import_ExpiredDate.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.Col_Import_ExpiredDate.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.Col_Import_ExpiredDate.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.Col_Import_ExpiredDate.Index).Value = Format(CDate(row.Cells(10).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(Me.Col_Import_SecondaryQty.Index).Value) = True Then
                    If row.Cells(Me.Col_Import_SecondaryQty.Index).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(Me.Col_Import_SecondaryQty.Index).Value = "-1" Then
                    row.Cells(Me.Col_Import_SecondaryQty.Index).Value = "-"
                    row.Cells(Me.Col_Import_SecondaryUnit.Index).Value = "-"
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(Me.Col_Import_PrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.Col_Import_PrimaryLocation.Index).Value = "-"
                    row.Cells(Me.Col_Import_Location.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.Col_Import_Batch.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If

                '1.e. jika ada invalid, latar belakang dimerahkan
                If String.IsNullOrEmpty(row.Cells(Me.Col_Import_Invalid.Index).Value) = False Then
                    row.DefaultCellStyle.BackColor = InvalidBackColor
                    row.DefaultCellStyle.ForeColor = InvalidForeColor
                End If
            Next
            '2. jika semua row pada dgvimport tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                Me.dgvImport.Columns(Me.Col_Import_SecondaryQty.Index).Visible = False
                Me.dgvImport.Columns(Me.Col_Import_SecondaryUnit.Index).Visible = False
            Else
                Me.dgvImport.Columns(Me.Col_Import_SecondaryQty.Index).Visible = True
                Me.dgvImport.Columns(Me.Col_Import_SecondaryUnit.Index).Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvImport.Columns(Me.Col_Import_WareHouse.Index).Visible = False
                Me.dgvImport.Columns(Me.Col_Import_Location.Index).Visible = False
            Else
                Me.dgvImport.Columns(Me.Col_Import_WareHouse.Index).Visible = True
                Me.dgvImport.Columns(Me.Col_Import_Location.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvImport.Columns(Me.Col_Import_Batch.Index).Visible = False
            Else
                Me.dgvImport.Columns(Me.Col_Import_Batch.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvImport.Columns(Me.Col_Import_ExpiredDate.Index).Visible = False
            Else
                Me.dgvImport.Columns(Me.Col_Import_ExpiredDate.Index).Visible = True
            End If
            Me.dgvImport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub removeNonInvalid()
        Dim i As Integer = 0
        Dim rowdgv As DataGridViewRow
        Try
            For i = Me.dgvImport.Rows.Count - 1 To 0 Step -1
                rowdgv = Me.dgvImport.Rows(i)
                If String.IsNullOrEmpty(rowdgv.Cells(Me.Col_Import_Invalid.Index).Value) Then
                    Me.dgvImport.Rows.Remove(rowdgv)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvImport_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvImport.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvImport_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvImport.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvImport)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    'END Coding tab IMPORT

End Class