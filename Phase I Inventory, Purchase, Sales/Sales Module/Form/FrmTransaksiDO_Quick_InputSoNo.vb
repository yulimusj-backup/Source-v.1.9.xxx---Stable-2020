Option Explicit On
Option Strict Off
Public Class FrmTransaksiDO_Quick_InputSoNo
    Public BackColorAll As Color

    Private Sub FrmTransaksiQuick_StockInPurchase_InputSupplierDoNo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItem.AllowUserToDeleteRows = False
            Me.DgvItem.AllowUserToAddRows = False
            Me.ColPrimaryCustomer.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Me.ApakahDataBenar = True Then
                DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                MsgBox("Customer so no in item cannot empty string.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Try
            For Each row As DataGridViewRow In Me.DgvItem.Rows
                If String.IsNullOrEmpty(row.Cells(Me.ColCustomerSoNo.Index).Value) = True And String.IsNullOrEmpty(row.Cells(Me.ColCustomerCode.Index).Value) = False Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColCustomerSoNo.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColCustomerCode.Index Then
                Me.DgvItem.Item(Me.ColCustomerSoNo.Index, e.RowIndex).Value = Me.DgvItem.Item(Me.ColCustomerSoNo.Index, e.RowIndex).Value.ToString.ToUpper.Trim
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class