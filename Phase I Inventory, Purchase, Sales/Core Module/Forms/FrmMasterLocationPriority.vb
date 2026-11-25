Option Explicit On
Option Strict Off

Public Class FrmMasterLocationPriority
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vTipeTransaksi As ClsMasterLocationPriority.EnumTipeTransaksi = ClsMasterLocationPriority.EnumTipeTransaksi.StockOutAdjust

    Dim Buffer As String = ""
    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim RemoveIndex As Integer = -1
    Dim vLangsungClose As Boolean = False
    Dim vBtnCloseNo As Boolean = False

    Private Sub FrmMasterLocationPriority_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If vBtnCloseNo = True Then
                e.Cancel = True
                vBtnCloseNo = False
            Else
                If vLangsungClose = False Then
                    If MsgBox("Save data before close?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.SimpanData()
                    Else
                        Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanData()
        Dim DTA As New DseMasterLocationPriorityTableAdapters.tblmaster_location_priority_TableAdapter
        Dim rowDGV As DataGridViewRow
        Dim Urutan As Integer = 1
        Dim vIsInclude As Integer = 0
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvItem.AllowUserToAddRows = False

            'simpan ke form
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.DeleteQuery(Me.UserCode, Me.vTipeTransaksi)

            If chkIncludeOtherLocation.Checked = True Then
                vIsInclude = 1
            End If
            For Each rowDGV In Me.dgvItem.Rows
                DTA.InsertQuery(Me.UserCode, rowDGV.Cells(Me.ColPrimaryLocation.Index).Value, _
                Urutan, Me.vTipeTransaksi, vIsInclude)

                Urutan += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Close form and save location priority data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.SimpanData()

                'close form
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.vLangsungClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Close and discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.vLangsungClose = True
                Me.Close()
            Else
                vBtnCloseNo = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterLocationPriority_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DTA As New DseMasterLocationPriorityTableAdapters.tblmaster_location_priority1_TableAdapter
        Dim TBL As New DseMasterLocationPriority.tblmaster_location_priority1_DataTable
        Dim row As DseMasterLocationPriority.tblmaster_location_priority1_Row
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'back color all
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'load data sesuai db
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AllowUserToAddRows = False

            DTA.Fill(TBL, Me.UserCode, Me.vTipeTransaksi)
            For Each row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColLocationCode.Index).Value = row.Storage_Area_Code
                    .Cells(Me.ColLocationDesc.Index).Value = row.Storage_Area_Desc
                    .Cells(Me.ColPrimaryLocation.Index).Value = row.primary_location
                    .Cells(Me.ColWarehouseCode.Index).Value = row.Warehouse_Code
                    .Cells(Me.ColWarehouseDesc.Index).Value = row.Warehouse_Desc
                    If row.include_other_location = 1 Then
                        Me.chkIncludeOtherLocation.Checked = True
                    End If
                End With
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AllowUserToAddRows = True
            Me.btnOK.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If String.IsNullOrEmpty(Me.dgvItem(Me.ColLocationCode.Index, e.RowIndex).Value) = False Then
                e.Cancel = True
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColLocationCode.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimaryLoc As Double) As Double
        Dim rowDGV As DataGridViewRow
        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLoc Then
                    Return rowDGV.Index
                End If
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim PrimaryLoc As Double

        Try
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'bila tidak ada perubahan
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'location code
                    If e.ColumnIndex = Me.ColLocationCode.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjLocation.ApakahSudahAda(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            PrimaryLoc = Me.ObjLocation.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.GetIndexItem(PrimaryLoc)
                            If index < 0 Then
                                Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value = Me.ObjLocation.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvItem.Item(Me.ColLocationCode.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(PrimaryLoc)
                                Me.dgvItem.Item(Me.ColLocationDesc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(PrimaryLoc)
                                Me.dgvItem.Item(Me.ColWarehouseCode.Index, e.RowIndex).Value = Me.ObjLocation.GetWarehouseCode(PrimaryLoc)
                                Me.dgvItem.Item(Me.ColWarehouseDesc.Index, e.RowIndex).Value = Me.ObjLocation.GetWarehouseDesc(PrimaryLoc)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication in location code, please change your location code", MsgBoxStyle.Information, "Information")
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                        Me.RemoveIndex = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            ''diketik tidak ada lalu tampilkan form cari
                            Dim ij As New Core_Module.FrmCariMasterStorageArea
                            ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.ShowDialog()
                            index = Me.GetIndexItem(ij.PrimaryStorageArea)
                            If ij.PrimaryStorageArea = -1 Then
                                ''lewat
                                If Me.Buffer = "" Then
                                    Me.RemoveIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                    Exit Sub
                                Else
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                End If
                            Else
                                If index <= 0 Then
                                    Me.dgvItem.Item(Me.ColWarehouseCode.Index, e.RowIndex).Value = ij.PrimaryStorageArea.ToString
                                    Me.dgvItem.Item(Me.ColLocationCode.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(ij.PrimaryStorageArea)
                                    Me.dgvItem.Item(Me.ColLocationDesc.Index, e.RowIndex).Value = Me.ObjLocation.GetDescription(ij.PrimaryStorageArea)
                                    Me.dgvItem.Item(Me.ColWarehouseCode.Index, e.RowIndex).Value = Me.ObjLocation.GetWarehouseCode(ij.PrimaryStorageArea)
                                    Me.dgvItem.Item(Me.ColWarehouseDesc.Index, e.RowIndex).Value = Me.ObjLocation.GetWarehouseDesc(ij.PrimaryStorageArea)
                                    Me.dgvItem.Item(Me.ColPrimaryLocation.Index, e.RowIndex).Value = ij.PrimaryStorageArea
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication in location code, please change your location code", MsgBoxStyle.Information, "Information")
                                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                            Me.RemoveIndex = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        End If
                                        Exit Sub
                                    End If
                                End If
                                ij = Nothing
                            End If
                        End If
                    End If
                Else
                    'klo nothing
                    If e.ColumnIndex = Me.ColLocationCode.Index Then
                        If Me.Buffer = "" Then
                            'remove row
                            Me.RemoveIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                            Exit Sub
                        Else
                            'balikin ke buffer
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowRowProcess(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False

            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)

                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRowProcess(row, True)
                    End If
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvItem.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            Me.dgvItem.AllowUserToAddRows = False

            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False

                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    ''Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRowProcess(row, False)
                End If

                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.dgvItem.AllowUserToAddRows = True
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
            If Me.RemoveIndex <> -1 Then
                Me.dgvItem.Rows.RemoveAt(Me.RemoveIndex)
                Me.RemoveIndex = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvItem.AllowUserToAddRows = True Then
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColLocationCode.Index).Value = ""
                rowDGV.Cells(Me.ColLocationDesc.Index).Value = ""
                rowDGV.Cells(Me.ColWarehouseCode.Index).Value = ""
                rowDGV.Cells(Me.ColWarehouseDesc.Index).Value = ""
                rowDGV.Cells(Me.ColPrimaryLocation.Index).Value = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            Else
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class