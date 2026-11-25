Option Explicit On
Option Strict Off

Public Class FrmTransaksiReminder
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen
    Public TanggalReminder As Date

    Dim Baru As Boolean
    Dim ObjReminder As New ClsTransaksiReminder
    Dim TBL As New DseTransaksiReminder.tbltrans_reminder1DataTable

    Private Sub FrmTransaksiReminder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ''Me.ObjReminder = New ClsTransaksiReminder
            Me.Baru = True
            Me.KosongkanField()
            ObjReminder.IsiDGV(Me.dgvItem, Me.TBL, Me.UserCode, Me.dtpDate.Value.Date)
            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtSubject.Text = Me.dgvItem.SelectedRows(0).Cells(1).Value
                Me.TxtMemo.Text = Me.dgvItem.SelectedRows(0).Cells(3).Value
                Me.dgvItem.ClearSelection()
            End If
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.dtpDate.Value = Now
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim ij As New FrmTransaksiReminder_New
        ij.UserCode = Me.UserCode
        ij.BackColorAll = Me.BackColorAll
        Try
            ij.Baru = True
            ij.dtpDate.Value = Me.dtpDate.Value.Date
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.TxtSubject.Clear()
                Me.TxtMemo.Clear()
                Me.ObjReminder.IsiDGV(Me.dgvItem, Me.TBL, Me.UserCode, Me.dtpDate.Value.Date)
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.TxtSubject.Text = Me.dgvItem.SelectedRows(0).Cells(1).Value
                    Me.TxtMemo.Text = Me.dgvItem.SelectedRows(0).Cells(3).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            Me.TxtSubject.Clear()
            Me.TxtMemo.Clear()
            ObjReminder.IsiDGV(Me.dgvItem, Me.TBL, Me.UserCode, Me.dtpDate.Value.Date)
            If Me.dgvItem.Rows.Count > 0 Then
                Me.TxtSubject.Text = Me.dgvItem.SelectedRows(0).Cells(1).Value
                Me.TxtMemo.Text = Me.dgvItem.SelectedRows(0).Cells(3).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim ij As New FrmTransaksiReminder_New
        ij.UserCode = Me.UserCode
        ij.BackColorAll = Me.BackColorAll
        Try
            If ObjReminder.ApakahBolehEdit(Me.dgvItem.SelectedRows(0).Cells(0).Value, UserCode) = True Then
                ij.Baru = False
                ij.PrimaryReminder = Me.dgvItem.SelectedRows(0).Cells(0).Value
                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    ''load data terakhir
                    Me.TxtSubject.Clear()
                    Me.TxtMemo.Clear()
                    Me.ObjReminder.IsiDGV(Me.dgvItem, Me.TBL, UserCode, Me.dtpDate.Value.Date)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim IndexLama As Double = -1

        Try
            If Me.ObjReminder.ApakahBolehHapus(Me.dgvItem.SelectedRows(0).Cells(0).Value, UserCode) = False Then
                Me.btnDelete.Focus()
            Else
                hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    IndexLama = Me.dgvItem.SelectedRows(0).Index
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.ObjReminder.Hapus(Me.dgvItem.SelectedRows(0).Cells(0).Value, UserCode) = True Then
                        ''Me.KosongkanField()
                        Me.btnNew.Focus()
                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        If IndexLama <= Me.dgvItem.RowCount - 1 Then
                            Me.dgvItem.Rows(IndexLama).Selected = True
                        End If
                    End If
                    Me.ObjReminder.IsiDGV(Me.dgvItem, Me.TBL, UserCode, Me.dtpDate.Value.Date)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                If Me.dgvItem.Rows.Count > 0 Then
                    Me.TxtSubject.Text = Me.dgvItem.SelectedRows(0).Cells(1).Value
                    Me.TxtMemo.Text = Me.dgvItem.SelectedRows(0).Cells(3).Value
                Else
                    Me.TxtSubject.Text = "-"
                    Me.TxtMemo.Text = "-"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class