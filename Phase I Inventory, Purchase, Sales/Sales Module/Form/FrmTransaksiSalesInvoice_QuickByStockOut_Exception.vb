Option Explicit On
Option Strict Off

Public Class FrmTransaksiSalesInvoice_QuickByStockOut_Exception
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Tbl As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_exceptionDataTable

    Private Sub FrmTransaksiSalesInvoice_QuickByStockOut_Exception_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim row As DseTransaksiSalesInvoice_QuickByStockOut.tbl_exceptionRow
        Dim rd As DataGridViewRow

        Try
            Me.ColNoOutSales.Visible = False
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each row In Tbl
                With row
                    rd = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
                    rd.Cells(colNo.Index).Value = rd.Index + 1
                    rd.Cells(colPrimaryStockOutSales.Index).Value = .primary_stock_out_sales
                    rd.Cells(ColNoOutSales.Index).Value = .stock_out_sales_no
                    rd.Cells(colReason.Index).Value = .reason
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class