Option Explicit On
Option Strict Off

Public Class FrmTransaksiPOOtorisasi_Removed
    Dim DTAPO As New DseTransaksiPOTableAdapters.tbltrans_po3TableAdapter
    Dim TBLPO As New DseTransaksiPO.tbltrans_po3DataTable
    Dim rowPO As DseTransaksiPO.tbltrans_po3Row
    Dim DTAPO1 As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub IsiDGV()
        Dim rowDGV As DataGridViewRow

        Try
            Me.DataGridView1.Rows.Clear()
            DTAPO.Connection.ConnectionString = Core_Module.GetConstring
            DTAPO.Fill(TBLPO)
            If TBLPO.Rows.Count > 0 Then
                For Each rowPO In TBLPO.Rows
                    Me.DataGridView1.Rows.Add()
                    rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.DataGridView1.Rows.Count
                    rowDGV.Cells(Me.ColPONumber.Index).Value = rowPO.No_Transaksi
                    rowDGV.Cells(Me.ColPrimaryPO.Index).Value = rowPO.Primary
                    rowDGV.Cells(Me.ColRemoved.Index).Value = Boolean.TrueString
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiPOOtorisasi_Removed_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.IsiDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaYgBerubah() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DataGridView1.Rows
                If rowDGV.Cells(Me.ColRemoved.Index).Value = Boolean.FalseString Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnChanged_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChanged.Click
        Try
            DTAPO1.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahAdaYgBerubah = True Then
                If MsgBox("Save data to database?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                    For Each rowDGV As DataGridViewRow In Me.DataGridView1.Rows
                        If rowDGV.Cells(Me.ColRemoved.Index).Value = Boolean.FalseString Then
                            DTAPO1.UpdateQuery_SetRemoved(0, rowDGV.Cells(Me.ColPrimaryPO.Index).Value)
                        End If
                    Next

                    Me.IsiDGV()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColRemoved.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbChange_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbChange.LinkClicked
        Dim rowDGV As DataGridViewRow

        Try
            If Me.llbChange.Text = "Check All" Then
                For Each rowDGV In Me.DataGridView1.Rows
                    rowDGV.Cells(Me.ColRemoved.Index).Value = Boolean.TrueString
                Next
            ElseIf Me.llbChange.Text = "Uncheck All" Then
                For Each rowDGV In Me.DataGridView1.Rows
                    rowDGV.Cells(Me.ColRemoved.Index).Value = Boolean.FalseString
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Try
            If Me.ApakahAdaYgBerbeda(Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If Me.DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                    Me.llbChange.Text = "Uncheck All"
                Else
                    Me.llbChange.Text = "Check All"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaYgBerbeda(ByVal vReference As Boolean) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DataGridView1.Rows
                If rowDGV.Cells(Me.ColRemoved.Index).Value <> vReference Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class