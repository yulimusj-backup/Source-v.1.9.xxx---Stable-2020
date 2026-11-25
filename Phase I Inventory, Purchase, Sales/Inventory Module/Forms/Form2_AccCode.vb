Option Explicit On
Option Strict Off

Public Class Form2_AccCode
    Dim Buffer As String = ""
    Dim ObjAkun As New Core_Module.ClsMasterAkun1
    Dim rowIndex As Integer = -1

    Private Sub Form2_AccCode_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim DTA As New DataSet2TableAdapters.tbltrans_custom_primary_akunTableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Save data to database?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.dgvItem.AllowUserToAddRows = False
                DTA.Delete()
                For Each rowDGV In Me.dgvItem.Rows
                    DTA.Insert(rowDGV.Cells(Me.ColPrimaryAkun.Index).Value)
                Next
                Me.dgvItem.AllowUserToAddRows = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_AccCode_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DTA As New DataSet2TableAdapters.DataTable7TableAdapter
        Dim TBL As New DataSet2.DataTable7DataTable
        Dim row As DataSet2.DataTable7Row
        Dim rowDGV As DataGridViewRow

        Try
            'load dari db
            DTA.Fill(TBL)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.AllowUserToAddRows = False
            For Each row In TBL.Rows
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                rowDGV.Cells(Me.ColAccount.Index).Value = row.Account_Code
                rowDGV.Cells(Me.ColDescAccount.Index).Value = row.Description
                rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = row.primary_akun
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvItem.AllowUserToAddRows = True

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColAccount.Index Then
                e.Cancel = True
                Exit Sub
            Else
                'col account
                If e.ColumnIndex = Me.ColAccount.Index Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True _
            Or Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                'kalo kode akun dihapus, maka hapus satu bari semuanya
                If e.RowIndex <> Me.dgvItem.RowCount - 1 Then
                    Me.rowIndex = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                    '  Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                End If
                Exit Sub
            Else
                ''data account number
                If e.ColumnIndex = Me.ColAccount.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                    'kemungkinan
                    If Me.ObjAkun.ApakahSudahAda_Aktif_NonHeader(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                        Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = Me.ObjAkun.GetPrimary(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                        Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value)
                    Else
                        '2. kode akun yang diketik tidak ada --> tampilkan form cari
                        Dim ij As New Core_Module.FrmMasterAkun_Cari1
                        ij.txtCari.Text = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.ShowDialog()
                        If ij.mPrimaryAkun <> -1 Then
                            '2.a. kode akun yang dicari ada --> load data deskripsi + primary akun + kode akun
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjAkun.GetAccountCode(ij.mPrimaryAkun)
                            Me.dgvItem.Item(Me.ColDescAccount.Index, e.RowIndex).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(ij.mPrimaryAkun)
                            Me.dgvItem.Item(Me.ColPrimaryAkun.Index, e.RowIndex).Value = ij.mPrimaryAkun
                        ElseIf ij.mPrimaryAkun = -1 Then
                            '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                                Me.rowIndex = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                                ' Me.dgvItem.Rows.RemoveAt(e.RowIndex)
                                Exit Sub
                            End If
                        End If
                        ij = Nothing
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Try
            Me.dgvItem.Rows.RemoveAt(Me.rowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Dim DefPrecBase As Integer = 0

        Try
            'set nilai default baris baru
            rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
            rowDGV.Cells(Me.ColPrimaryAkun.Index).Value = -1
            rowDGV.Cells(Me.ColAccount.Index).Value = ""
            rowDGV.Cells(Me.ColDescAccount.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow

    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow

    End Sub
End Class