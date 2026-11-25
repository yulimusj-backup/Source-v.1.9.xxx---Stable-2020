Option Explicit On
Option Strict Off

Public Class FrmMasterPriorityLevel

    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim obj As New ClsMasterPriorityLevel
    Dim objCd As New Core_Module.ClsCekData
    Dim buffer As String = ""
    Dim indexRowDgv As Integer = -1

    Private Sub FrmMasterPriorityLevel_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.obj.refreshAllData(Me, Me.DataGridView1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCode.Index Or e.ColumnIndex = Me.ColDescription.Index Or e.ColumnIndex = Me.ColDefault.Index Then ''Hanya periksa kolom tertentu
                If (e.ColumnIndex = Me.ColDescription.Index Or e.ColumnIndex = Me.ColDefault.Index) And (Me.DataGridView1.Item(Me.ColCode.Index, e.RowIndex).Value = "") Then ''Bisa edit yang lain, asal code sudah diisi
                    e.Cancel = True
                Else
                    Me.buffer = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
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
            If e.ColumnIndex = Me.ColCode.Index And String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
            End If
            If Me.DataGridView1.Item(Me.ColPrimary.Index, e.RowIndex).Value = "-1" Then ''INSERT::
                ''Code::
                If e.ColumnIndex = Me.ColCode.Index Then
                    ''Berisi::
                    If String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        'MsgBox("Code can't be an empty string", MsgBoxStyle.Information, "Exlamation")
                        If e.RowIndex <> Me.DataGridView1.Rows.Count - 1 Then
                            Me.indexRowDgv = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If
                        Exit Sub
                    End If
                    ''Unik::
                    If Me.obj.apakahSudahAda(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = True Then
                        MsgBox("Duplicate code, please try with another Code", MsgBoxStyle.Information, "Exlamation")
                        Me.indexRowDgv = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    End If
                    ''Bersih::
                    If Me.isCleanString(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        MsgBox("Code can't contain these charaters; ', #, %, *", MsgBoxStyle.Information, "Exlamation")
                        Me.indexRowDgv = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                        Exit Sub
                    End If
                    ''Save It::
                    If Me.obj.simpanBaru(Me.DataGridView1.Item(Me.ColPriority.Index, e.RowIndex).Value.ToString, _
                            Me.DataGridView1.Item(Me.ColCode.Index, e.RowIndex).Value.ToString, _
                            Me.DataGridView1.Item(Me.ColDescription.Index, e.RowIndex).Value.ToString, _
                            Me.UserCode) = True Then
                        Me.DataGridView1.Item(Me.ColPrimary.Index, e.RowIndex).Value = Me.obj.getPrimary(Me.DataGridView1.Item(Me.ColCode.Index, e.RowIndex).Value.ToString).ToString
                        Exit Sub
                    Else
                        MsgBox("Data save failed", MsgBoxStyle.Information, "Exlamation")
                        Exit Sub
                    End If
                End If
            Else ''UPDATE::
                ''Code::
                If e.ColumnIndex = Me.ColCode.Index Then
                    ''Berisi::
                    If String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        If MsgBox("Do you want to delete this priority?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Delete") = MsgBoxResult.Yes Then
                            If e.RowIndex <> Me.DataGridView1.Rows.Count - 1 Then
                                Me.indexRowDgv = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            End If
                        Else
                            Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                        'MsgBox("Code can't be an empty string", MsgBoxStyle.Information, "Exlamation")
                        Exit Sub
                    End If
                    ''Unik::
                    If Me.obj.apakahSudahAda(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = True And Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> Me.buffer Then
                        MsgBox("Duplicate code, please try with another Code", MsgBoxStyle.Information, "Exlamation")
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    ''Bersih::
                    If Me.isCleanString(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                        MsgBox("Code can't contain these charaters; ', #, %, *", MsgBoxStyle.Information, "Exlamation")
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                If e.ColumnIndex = Me.ColDescription.Index Then
                    ''Bersih::
                    If String.IsNullOrEmpty(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.isCleanString(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("Description can't contain these charaters; ', #, %, *", MsgBoxStyle.Information, "Exlamation")
                            Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    Else
                        MsgBox("Description can't be an empty string", MsgBoxStyle.Information, "Exlamation")
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                If e.ColumnIndex = Me.ColDefault.Index Then
                    If Me.GetJmlDefault = 0 Then
                        Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If

                    If Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                        Me.indexRowDgv = e.RowIndex
                        Dim rowDgv As New DataGridViewRow
                        For Each rowDgv In Me.DataGridView1.Rows
                            If Me.DataGridView1.Rows.IndexOf(rowDgv) <> Me.indexRowDgv Then
                                rowDgv.Cells(Me.ColDefault.Index).Value = Boolean.FalseString
                            End If
                        Next
                        ''set default all::
                        Me.obj.setDefaultAll(Me.UserCode, CDbl(Me.DataGridView1.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString))
                    ElseIf Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString Then
                        ''set default once::
                        Me.obj.setDefaultByPrimary(Me.UserCode, ClsMasterPriorityLevel.enumDefault.no, CDbl(Me.DataGridView1.Item(Me.ColPrimary.Index, e.RowIndex).Value.ToString))
                    End If
                End If

                ''Update it
                If Me.obj.updateDefault(Me.DataGridView1.Item(Me.ColCode.Index, e.RowIndex).Value, _
                        Me.DataGridView1.Item(Me.ColDescription.Index, e.RowIndex).Value, _
                        Me.UserCode, Me.DataGridView1.Item(Me.ColPrimary.Index, e.RowIndex).Value) = True Then
                    Exit Sub
                Else
                    MsgBox("Data change failed", MsgBoxStyle.Information, "Exlamation")
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
                If rowDGV.Cells(Me.ColDefault.Index).Value = Boolean.TrueString Then
                    Retval += 1
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DataGridView1_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValidated
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Sub

    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

    End Sub

    Private Sub DataGridView1_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        Try
            Me.DataGridView1.Item(Me.ColPrimary.Index, Me.DataGridView1.Rows.Count - 1).Value = "-1"
            Me.DataGridView1.Item(Me.ColPriority.Index, Me.DataGridView1.Rows.Count - 1).Value = (Me.DataGridView1.Rows.Count).ToString
            Me.DataGridView1.Item(Me.ColCode.Index, Me.DataGridView1.Rows.Count - 1).Value = ""
            Me.DataGridView1.Item(Me.ColDescription.Index, Me.DataGridView1.Rows.Count - 1).Value = "-"
            Me.DataGridView1.Item(Me.ColDefault.Index, Me.DataGridView1.Rows.Count - 1).Value = False
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
            Me.objCd.SetNumber(Me.DataGridView1)
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally

        End Try
    End Sub

    Private Sub DataGridView1_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DataGridView1.UserDeletedRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.objCd.SetNumber(Me.DataGridView1)
            Me.obj.repairPriority(Me, Me.DataGridView1)
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
            If Me.DataGridView1.Item(Me.ColDefault.Index, e.Row.Index).Value = Boolean.TrueString Then
                e.Cancel = True
                Exit Sub
            End If

            'jika sudah digunakan pada item so, maka tidak bisa dihapuskan
            If Me.obj.ApakahBolehHapus(Me.DataGridView1.Item(Me.ColPrimary.Index, e.Row.Index).Value, True) = False Then
                e.Cancel = True
            Else
                Me.obj.hapus(CDbl(Me.DataGridView1.Item(Me.ColPrimary.Index, Me.DataGridView1.Rows.IndexOf(e.Row)).Value))
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
                    Me.obj.repairPriority(Me, Me.DataGridView1)
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
                    Me.obj.repairPriority(Me, Me.DataGridView1)
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