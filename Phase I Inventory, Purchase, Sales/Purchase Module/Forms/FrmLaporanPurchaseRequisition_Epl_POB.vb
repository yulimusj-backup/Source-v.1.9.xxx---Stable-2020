Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseRequisition_Epl_POB
    Public PrimaryProduct As Double = -1
    Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_item_view6TableAdapter
    Dim TBL As New DseTransaksiPO.tbltrans_po_item_view6DataTable
    Dim ObjProd As New Core_Module.clsMasterProduct

    Private Sub FrmLaporanPurchaseRequisition_Epl_POB_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.PrimaryProduct <> -1 Then
                DTA.Fill(Me.TBL, Me.PrimaryProduct)
                Me.DataGridView1.DataSource = Me.TBL.DefaultView
                If Me.DataGridView1.RowCount > 0 Then
                    Me.DataGridView1.Columns(1).DefaultCellStyle.Format = "dd MMM yyyy"
                    Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    Me.DataGridView1.Columns(2).DefaultCellStyle.Format = "N0"
                    Me.DataGridView1.Columns(2).HeaderText = "Purchase Order Balance"
                End If
                Me.Text = "Product Order Balance Detail - " + Me.ObjProd.GetProductCode(Me.PrimaryProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class