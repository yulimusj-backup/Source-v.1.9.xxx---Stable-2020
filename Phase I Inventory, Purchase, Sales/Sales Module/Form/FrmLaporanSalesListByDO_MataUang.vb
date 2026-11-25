Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesListByDO_MataUang
    Dim Buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang boleh diedit hanya kolom exhange rate to IDR
            If e.ColumnIndex <> Me.ColExchangeRate.Index Then
                e.Cancel = True
            Else
                Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = CDbl(Me.Buffer) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If e.ColumnIndex <> Me.ColExchangeRate.Index Then
                        'pastikan bilangan positif
                        If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number for exchange rate data.", MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesListByDO_MataUang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class