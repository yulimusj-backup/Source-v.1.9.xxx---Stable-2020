Option Explicit On
Option Strict Off
Public Class FrmTrans_StockBundling_InputLocationPriority
    Public TBLLocationPriority As New Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Dim rowDgv As DataGridViewRow
    Dim rowLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row
    Dim objLocation As New Core_Module.clsMasterStorageArea
    Dim objCD As New Core_Module.ClsCekData
    Dim buffer As Object
    Dim rowIndexToDelete As Integer

    Private Sub FrmTrans_StockBundling_InputLocationPriority_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.TBLLocationPriority.Rows.Count > 0 Then
                Me.rowLocationPriority = Me.TBLLocationPriority.Rows(0)
                Me.rowDgv = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Add)
                With Me.rowDgv
                    .Cells(Me.colNo.Index).Value = 1
                    .Cells(Me.ColLocationCode.Index).Value = Me.rowLocationPriority.Storage_Area_Code
                    .Cells(Me.ColLocationDesc.Index).Value = Me.rowLocationPriority.Storage_Area_Desc
                    .Cells(Me.ColPrimaryLocation.Index).Value = Me.rowLocationPriority.primary_location
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColLocationCode.Index Then
                e.Cancel = True
                Exit Sub
            End If
            Me.buffer = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim frmFindLocation As New Core_Module.FrmCariMasterStorageArea
        Try
            If e.ColumnIndex = Me.ColLocationCode.Index Then
                If Not IsNothing(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    If Not String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) Then
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                        If Me.objLocation.ApakahSudahAda(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Dim vPrimaryLocation As Double = Me.objLocation.GetPrimary(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value)
                            If Me.GetIndex(vPrimaryLocation, e.RowIndex) > -1 Then
                                MsgBox(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                If String.IsNullOrEmpty(Me.buffer) Then
                                    Me.rowIndexToDelete = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Exit Sub
                            End If
                            Me.DataGridView1.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value = vPrimaryLocation
                            Me.DataGridView1.Item(Me.ColLocationDesc.Index, e.RowIndex).Value = Me.objLocation.GetDescription(vPrimaryLocation)
                        Else
                            frmFindLocation.txtCari.Text = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
                            frmFindLocation.ShowDialog()
                            If frmFindLocation.PrimaryStorageArea > -1 Then
                                Dim code As String = Me.objLocation.GetStorageAreaCode(frmFindLocation.PrimaryStorageArea)
                                If Me.GetIndex(frmFindLocation.PrimaryStorageArea, e.RowIndex) > -1 Then
                                    MsgBox(code & " already exist.", MsgBoxStyle.Information, "Information")
                                    If String.IsNullOrEmpty(Me.buffer) Then
                                        Me.rowIndexToDelete = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                                Me.DataGridView1.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value = frmFindLocation.PrimaryStorageArea
                                Me.DataGridView1.Item(Me.ColLocationCode.Index, e.RowIndex).Value = code
                                Me.DataGridView1.Item(Me.ColLocationDesc.Index, e.RowIndex).Value = Me.objLocation.GetDescription(frmFindLocation.PrimaryStorageArea)
                            Else
                                If String.IsNullOrEmpty(Me.buffer) Then
                                    Me.rowIndexToDelete = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                                Exit Sub
                            End If
                        End If
                    Else
                        If String.IsNullOrEmpty(Me.buffer) Then
                            Me.rowIndexToDelete = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        Exit Sub
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) Then
                        Me.rowIndexToDelete = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    Else
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex(ByVal vPrimaryLocation As Double, ByVal vCurrentRow As Integer) As Integer
        Try
            For Each row As DataGridViewRow In Me.DataGridView1.Rows
                If Me.DataGridView1.Rows.IndexOf(row) <> vCurrentRow Then
                    If Not IsNothing(row.Cells(Me.ColPrimaryLocation.Index).Value) Then
                        If Not String.IsNullOrEmpty(row.Cells(Me.ColPrimaryLocation.Index).Value) Then
                            If row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                                Return Me.DataGridView1.Rows.IndexOf(row)
                            End If
                        End If
                    End If
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -2
        End Try
    End Function

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        Dim row As DataGridViewRow
        Try
            row = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
            row.Cells(Me.colNo.Index).Value = Me.DataGridView1.Rows.Count
            row.Cells(Me.ColPrimaryLocation.Index).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        Try
            Me.objCD.SetNumber(Me.DataGridView1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Threading.Thread.Sleep(10)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.rowIndexToDelete > -1 And Me.rowIndexToDelete <> Me.DataGridView1.Rows.Count - 1 Then
                Me.DataGridView1.Rows.RemoveAt(Me.rowIndexToDelete)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.TBLLocationPriority.Rows.Clear()
            Me.DataGridView1.AllowUserToAddRows = False
            For Each rowDgv In Me.DataGridView1.Rows
                Me.rowLocationPriority = Me.TBLLocationPriority.Newtblmaster_location_priority1_Row
                rowLocationPriority.primary = Me.DataGridView1.Rows.IndexOf(rowDgv) + 1
                rowLocationPriority.user_name = UserCode
                rowLocationPriority.primary_location = rowDgv.Cells(Me.ColPrimaryLocation.Index).Value
                rowLocationPriority.urutan = Me.DataGridView1.Rows.IndexOf(rowDgv) + 1
                rowLocationPriority.tipe_transaksi = 2
                rowLocationPriority.Storage_Area_Code = rowDgv.Cells(Me.ColLocationCode.Index).Value
                rowLocationPriority.Storage_Area_Desc = rowDgv.Cells(Me.ColLocationDesc.Index).Value
                rowLocationPriority.Warehouse_Code = String.Empty
                rowLocationPriority.Warehouse_Desc = String.Empty
                rowLocationPriority.include_other_location = 0
                Me.TBLLocationPriority.Rows.Add(Me.rowLocationPriority)
            Next
            Me.DialogResult = Windows.Forms.DialogResult.Yes
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DataGridView1.AllowUserToAddRows = True
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Sorted
        Try
            Me.objCD.SetNumber(Me.DataGridView1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            Me.DataGridView1.AllowUserToAddRows = False
            If Me.DataGridView1.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DataGridView1.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.DataGridView1.Rows.Remove(row)
                    Me.DataGridView1.Rows.InsertRange(indexbaru, row)
                    Me.DataGridView1.Focus()
                    Me.DataGridView1.ClearSelection()
                    row.Selected = True
                    Me.objCD.SetNumber(Me.DataGridView1)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DataGridView1.AllowUserToAddRows = True
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            Me.DataGridView1.AllowUserToAddRows = False
            If Me.DataGridView1.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DataGridView1.SelectedRows(0)
                If row.Index <> Me.DataGridView1.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.DataGridView1.Rows.Remove(row)
                    Me.DataGridView1.Rows.InsertRange(indexbaru, row)
                    Me.DataGridView1.Focus()
                    Me.DataGridView1.ClearSelection()
                    row.Selected = True
                    Me.objCD.SetNumber(Me.DataGridView1)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DataGridView1.AllowUserToAddRows = True
        End Try
    End Sub

    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.DataGridView1.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.DataGridView1.FirstDisplayedScrollingRowIndex = Me.DataGridView1.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.DataGridView1.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.DataGridView1.Rows(row.Index).Displayed = False
                    Me.DataGridView1.FirstDisplayedScrollingRowIndex = Me.DataGridView1.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.DataGridView1.DisplayedRowCount(True) <> Me.DataGridView1.DisplayedRowCount(False) Then
                    Me.DataGridView1.FirstDisplayedScrollingRowIndex = Me.DataGridView1.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class