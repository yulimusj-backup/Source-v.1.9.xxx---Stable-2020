Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_Data1_Exclusion
    Dim Buffer As String = ""
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim DummyIndex As Integer = -1
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Accept data and proceed?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Close and cancel calculation?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Data1_Exclusion_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'tampilkan semua data product yang di exclude untuk perhitungan material planning
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_productTableAdapter
        Dim TBL As New DseTransaksiMaterialPlanningGroup.tblmaster_productDataTable
        Dim row As DseTransaksiMaterialPlanningGroup.tblmaster_productRow
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_AllExclude(TBL)

            Me.DgvExclusion.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvExclusion.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DgvExclusion.AllowUserToAddRows = False
            Me.DgvExclusion.Rows.Clear()

            For Each row In TBL.Rows
                Me.DgvExclusion.Rows.Add()
                rowDGV = Me.DgvExclusion.Rows(Me.DgvExclusion.RowCount - 1)
                rowDGV.Cells(Me.ColCodeNo.Index).Value = row.Product_Code
                rowDGV.Cells(Me.ColDesc.Index).Value = row.Description
                rowDGV.Cells(Me.ColPrimary.Index).Value = row.Primary
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
            Next

            Me.DgvExclusion.AllowUserToAddRows = True
            Me.DgvExclusion.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
        End Try
    End Sub

    Private Sub DgvExclusion_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvExclusion.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCodeNo.Index Then
                Me.Buffer = Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvExclusion_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvExclusion.CellEndEdit
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_productTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If IsNothing(Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColCodeNo.Index Then
                    'kalo kode product dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker2.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                If e.ColumnIndex = Me.ColCodeNo.Index Then
                    'code
                    'pastikan ditrim dan huruf besar, max 250 karakter
                    Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'cek apakah ada ada dalam master product
                    'kemungkinan
                    If Me.ObjProduct.ApakahSudahAda(Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        'update data lama
                        If Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value <> -1 Then
                            DTA.UpdateQuery_SetExclude(0, Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                        End If
                        '1. kode product ada  --> load data nama + primary
                        Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.DgvExclusion.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                        'update data baru
                        DTA.UpdateQuery_SetExclude(1, Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                    Else
                        '2. kode product yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.frmCariMasterProduct
                        ij.txtCari.Text = Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.PrimaryProduct <> -1 Then
                            '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                            If Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value <> -1 Then
                                DTA.UpdateQuery_SetExclude(0, Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                            End If
                            Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value = ij.PrimaryProduct
                            Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                            Me.DgvExclusion.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                            DTA.UpdateQuery_SetExclude(1, Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                        ElseIf ij.PrimaryProduct = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.DummyIndex = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                                ij = Nothing
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                    'pastikan tidak double dengan yang lainnya
                    DummyIndex = Me.GetIndex_Itemproduct(Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex)
                    If DummyIndex <> -1 Then
                        MsgBox("product code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                        If Buffer = "" Then
                            Me.DummyIndex = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            'kembalikan lagi datanya
                            DTA.UpdateQuery_SetExclude(1, Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value) '-> karena double, maka ini tetap exclude
                            Me.DgvExclusion.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.Buffer)
                            Me.DgvExclusion.Item(Me.ColDesc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value)
                            DTA.UpdateQuery_SetExclude(1, Me.DgvExclusion.Item(Me.ColPrimary.Index, e.RowIndex).Value) 'yang buffer juga tetap exclude
                        End If
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Function GetIndex_Itemproduct(ByVal vProductCode As String, ByVal vRefIndex As Integer) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DgvExclusion.Rows
                If rowDGV.Index <> vRefIndex Then
                    If rowDGV.Cells(Me.ColCodeNo.Index).Value = vProductCode Then
                        Return rowDGV.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub DgvExclusion_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvExclusion.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.DgvExclusion.Rows(Me.DgvExclusion.RowCount - 1)
            rowDGV.Cells(Me.ColCodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColDesc.Index).Value = ""
            rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColPrimary.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvExclusion_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvExclusion.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvExclusion)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvExclusion_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvExclusion.UserDeletingRow
        Dim DTA As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_productTableAdapter

        Try
            If MsgBox("Delete item from exclusion?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA.UpdateQuery_SetExclude(0, Me.DgvExclusion.Item(Me.ColPrimary.Index, e.Row.Index).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
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
                Me.DgvExclusion.Rows.RemoveAt(Me.DummyIndex)
                Me.DummyIndex = -1
                Me.ObjCD.SetNumber(Me.DgvExclusion)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class