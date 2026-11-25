Option Explicit On
Option Strict Off

Partial Public Class FrmSettingAP_AR
    Dim obj As New core_module.ClsMasterPriorityLevel
    Dim indexRowDgv As Integer = -1

    Private Sub FrmMasterPriorityLevel_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.obj.refreshAllData(Me.DataGridView1, Me.ColDP_Primary.Index, Me.ColDP_Priority.Index, _
            Me.ColDP_Code.Index, Me.ColDP_Description.Index, Me.ColDP_Default.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColDP_Code.Index Or e.ColumnIndex = Me.ColDP_Description.Index Or e.ColumnIndex = Me.ColDP_Default.Index Then ''Hanya periksa kolom tertentu
                If (e.ColumnIndex = Me.ColDP_Description.Index Or e.ColumnIndex = Me.ColDP_Default.Index) And (Me.DataGridView1.Item(Me.ColDP_Code.Index, e.RowIndex).Value = "") Then ''Bisa edit yang lain, asal code sudah diisi
                    e.Cancel = True
                Else
                    Me.Buffer = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColDP_Code.Index And String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
            End If
            If Me.DataGridView1.Item(Me.ColDP_Primary.Index, e.RowIndex).Value = "-1" Then ''INSERT::
                ''Code::
                If e.ColumnIndex = Me.ColDP_Code.Index Then
                    ''Berisi::
                    If String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        'MsgBox("Code can't be an empty string", MsgBoxStyle.Information, "Information")
                        If e.RowIndex <> Me.DataGridView1.Rows.Count - 1 Then
                            Me.indexRowDgv = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If
                    ''Unik::
                    If Me.obj.apakahSudahAda(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = True Then
                        MsgBox("Duplicate code, please try with another Code", MsgBoxStyle.Information, "Information")
                        Me.indexRowDgv = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    End If
                    ''Bersih::
                    If Me.isCleanString(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        MsgBox("Code can't contain these charaters; ', #, %, *", MsgBoxStyle.Information, "Information")
                        Me.indexRowDgv = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    End If
                    ''Save It::
                    If Me.obj.simpanBaru(Me.DataGridView1.Item(Me.ColDP_Priority.Index, e.RowIndex).Value.ToString, _
                            Me.DataGridView1.Item(Me.ColDP_Code.Index, e.RowIndex).Value.ToString, _
                            Me.DataGridView1.Item(Me.ColDP_Description.Index, e.RowIndex).Value.ToString, _
                            Me.UserCode) = True Then
                        Me.DataGridView1.Item(Me.ColDP_Primary.Index, e.RowIndex).Value = Me.obj.getPrimary(Me.DataGridView1.Item(Me.ColDP_Code.Index, e.RowIndex).Value.ToString).ToString
                        Exit Sub
                    Else
                        MsgBox("Data save failed", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            Else ''UPDATE::
                ''Code::
                If e.ColumnIndex = Me.ColDP_Code.Index Then
                    ''Berisi::
                    If String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        If MsgBox("Do you want to delete this data?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Delete") = MsgBoxResult.Yes Then
                            If e.RowIndex <> Me.DataGridView1.Rows.Count - 1 Then
                                Me.indexRowDgv = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                        Else
                            Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                        'MsgBox("Code can't be an empty string", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    ''Unik::
                    If Me.obj.apakahSudahAda(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = True And Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> Me.Buffer Then
                        MsgBox("Code duplication, please try with another code", MsgBoxStyle.Information, "Information")
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    ''Bersih::
                    If Me.isCleanString(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        MsgBox("Code data can't contain these charaters; ', #, %, *", MsgBoxStyle.Information, "Information")
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
                If e.ColumnIndex = Me.ColDP_Description.Index Then
                    ''Bersih::
                    If String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.isCleanString(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Description can't contain these charaters; ', #, %, *", MsgBoxStyle.Information, "Information")
                            Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    Else
                        MsgBox("Description can't be an empty string", MsgBoxStyle.Information, "Information")
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
                If e.ColumnIndex = Me.ColDP_Default.Index Then
                    If Me.GetJmlDefault = 0 Then
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                        Me.indexRowDgv = e.RowIndex
                        Dim rowDgv As New DataGridViewRow
                        For Each rowDgv In Me.DataGridView1.Rows
                            If Me.DataGridView1.Rows.IndexOf(rowDgv) <> Me.indexRowDgv Then
                                rowDgv.Cells(Me.ColDP_Default.Index).Value = Boolean.FalseString
                            End If
                        Next
                        ''set default all::
                        Me.obj.setDefaultAll(Me.UserCode, CDbl(Me.DataGridView1.Item(Me.ColDP_Primary.Index, e.RowIndex).Value.ToString))
                    ElseIf Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString Then
                        ''set default once::
                        Me.obj.setDefaultByPrimary(Me.UserCode, Core_Module.ClsMasterPriorityLevel.enumDefault.no, CDbl(Me.DataGridView1.Item(Me.ColDP_Primary.Index, e.RowIndex).Value.ToString))
                    End If
                End If

                ''Update it
                If Me.obj.updateDefault(Me.DataGridView1.Item(Me.ColDP_Code.Index, e.RowIndex).Value, _
                        Me.DataGridView1.Item(Me.ColDP_Description.Index, e.RowIndex).Value, _
                        Me.UserCode, Me.DataGridView1.Item(Me.ColDP_Primary.Index, e.RowIndex).Value) = True Then
                    Exit Sub
                Else
                    MsgBox("Data change failed", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Function GetJmlDefault() As Double
        Dim rowDGV As DataGridViewRow
        Dim Retval As Integer = 0

        Try
            For Each rowDGV In Me.DataGridView1.Rows
                If rowDGV.Cells(Me.ColDP_Default.Index).Value = Boolean.TrueString Then
                    Retval += 1
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        Try
            Me.DataGridView1.Item(Me.ColDP_Primary.Index, Me.DataGridView1.Rows.Count - 1).Value = "-1"
            Me.DataGridView1.Item(Me.ColDP_Priority.Index, Me.DataGridView1.Rows.Count - 1).Value = (Me.DataGridView1.Rows.Count).ToString
            Me.DataGridView1.Item(Me.ColDP_Code.Index, Me.DataGridView1.Rows.Count - 1).Value = ""
            Me.DataGridView1.Item(Me.ColDP_Description.Index, Me.DataGridView1.Rows.Count - 1).Value = "-"
            Me.DataGridView1.Item(Me.ColDP_Default.Index, Me.DataGridView1.Rows.Count - 1).Value = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DataGridView1.Rows.RemoveAt(Me.indexRowDgv)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Sub

    Private Sub DataGridView1_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DataGridView1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Sub

    Private Function isCleanString(ByVal vStr As String) As Boolean
        Try
            If String.IsNullOrEmpty(vStr) = False Then
                If vStr.Contains("'") = True Or vStr.Contains("#") = True Or vStr.Contains("%") = True Or vStr.Contains("*") = True Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Function

    Private Sub DataGridView1_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DataGridView1.UserDeletedRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.ObjCD.SetNumber(Me.DataGridView1)
            Me.obj.repairPriority(Me.DataGridView1, Me.ColDP_Primary.Index, Me.ColDP_Priority.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DataGridView1_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DataGridView1.UserDeletingRow
        Try
            Me.Cursor = Cursors.WaitCursor
            'jika ini adalah delivery priority terakhir maka tidak bisa dihapuskan
            If Me.DataGridView1.RowCount = 2 Then
                e.Cancel = True
            End If
            'jika default maka tidak bisa diedit
            If Me.DataGridView1.Item(Me.ColDP_Default.Index, e.Row.Index).Value = Boolean.TrueString Then
                e.Cancel = True
                Exit Sub
            End If

            'jika sudah digunakan pada item so, maka tidak bisa dihapuskan
            If Me.obj.ApakahBolehHapus(Me.DataGridView1.Item(Me.ColDP_Primary.Index, e.Row.Index).Value, True) = False Then
                e.Cancel = True
            Else
                Me.obj.hapus(CDbl(Me.DataGridView1.Item(Me.ColDP_Primary.Index, Me.DataGridView1.Rows.IndexOf(e.Row)).Value))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ShowRowProcess(ByVal vRow As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If vRow.Index < Me.DataGridView1.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While vRow.Displayed = False
                    Me.DataGridView1.FirstDisplayedScrollingRowIndex = Me.DataGridView1.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf vRow.Index > Me.DataGridView1.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.DataGridView1.Rows(vRow.Index).Displayed = False
                    Me.DataGridView1.FirstDisplayedScrollingRowIndex = Me.DataGridView1.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row paling bawah, pastikan terlihat
                If Me.DataGridView1.DisplayedRowCount(True) <> Me.DataGridView1.DisplayedRowCount(False) Then
                    Me.DataGridView1.FirstDisplayedScrollingRowIndex = Me.DataGridView1.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.DataGridView1.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = True = False
                Me.btnDown.Enabled = True = False
                Dim Row As DataGridViewRow = Me.DataGridView1.SelectedRows(0)
                If Row.Index <> 0 Then
                    Dim IndexBaru As Integer = Row.Index - 1
                    Me.DataGridView1.Rows.Remove(Row)
                    Me.DataGridView1.Rows.InsertRange(IndexBaru, Row)
                    Me.DataGridView1.Focus()
                    Me.DataGridView1.ClearSelection()
                    Row.Selected = True
                    Me.objCd.SetNumber(Me.DataGridView1)
                    Me.obj.repairPriority(Me.DataGridView1, Me.ColDP_Primary.Index, Me.ColDP_Priority.Index)
                    If Row.Displayed = False Then
                        Me.ShowRowProcess(Row, True)
                    End If
                End If
                Me.btnUP.Enabled = True = True
                Me.btnDown.Enabled = True = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.DataGridView1.SelectedRows.Count > 0 Then
                If Me.DataGridView1.Rows.IndexOf(Me.DataGridView1.SelectedRows.Item(0)) >= Me.DataGridView1.Rows.Count - 2 Then
                    Exit Sub
                End If
                Me.btnUP.Enabled = True = False
                Me.btnDown.Enabled = True = False
                Dim Row As DataGridViewRow = Me.DataGridView1.SelectedRows(0)
                If Row.Index <> Me.DataGridView1.Rows.Count - 1 Then
                    Dim IndexBaru As Integer = Row.Index + 1
                    Me.DataGridView1.Rows.Remove(Row)
                    Me.DataGridView1.Rows.InsertRange(IndexBaru, Row)
                    Me.DataGridView1.Focus()
                    Me.DataGridView1.ClearSelection()
                    Row.Selected = True
                    Me.objCd.SetNumber(Me.DataGridView1)
                    Me.obj.repairPriority(Me.DataGridView1, Me.ColDP_Primary.Index, Me.ColDP_Priority.Index)
                    Me.ShowRowProcess(Row, False)
                End If
                Me.btnUP.Enabled = True = True
                Me.btnDown.Enabled = True = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
