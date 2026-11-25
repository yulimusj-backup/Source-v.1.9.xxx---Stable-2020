Option Explicit On
Option Strict Off

Public Class frmMasterProduct_Verifikasi_Cek
    Public TBL_Result As New DseMasterProductAccountGroup.tblmaster_productDataTable
    Public row_Result As DseMasterProductAccountGroup.tblmaster_productRow

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmMasterProduct_Verifikasi_Cek_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow

        Try
            TBL_Result.DefaultView.RowFilter = ""
            TBL_Result.DefaultView.Sort = "[Product Code]"

            For Each DV As DataRowView In Me.TBL_Result.DefaultView
                Me.row_Result = DV.Row
                Me.dgvProduct.Rows.Add()
                rowDGV = Me.dgvProduct.Rows(Me.dgvProduct.Rows.Count - 1)
                rowDGV.Cells(Me.ColCodeNo.Index).Value = Me.row_Result.Product_Code
                rowDGV.Cells(Me.ColDesc.Index).Value = Me.row_Result.Description
                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvProduct.Rows.Count
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class