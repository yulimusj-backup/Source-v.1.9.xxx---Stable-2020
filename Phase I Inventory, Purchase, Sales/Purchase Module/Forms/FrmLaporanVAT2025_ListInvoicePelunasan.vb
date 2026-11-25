
Option Explicit On
Option Strict Off

Imports System.Data
Imports Microsoft.VisualBasic.FileIO

Public Class FrmLaporanVAT2025_ListInvoicePelunasan
    Dim ObjCD As New Core_Module.ClsCekData
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    '2025 Core Tax Tahap 2.1
    'Public TBLOut_SalesPelunasan As New DseLaporanVAT2025.InvoiceDataTable
    Public DateFrom As Date
    Public DateTo As Date
    Dim rowOut_pelunasan As DseLaporanVAT2025.InvoiceRow
    Dim TBLInvoice_InvoiceDP As New DseLaporanVAT2025_1.tblInvoice_InvoiceDPDataTable
    Dim row As DseLaporanVAT2025_1.tblInvoice_InvoiceDPRow

    Private Sub FrmLaporanVAT2025_ListInvoicePelunasan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Me.Text = "List of Invoice With DP"
            Me.Label1.Text = "There are settlement invoice transactions in this period that cannot be generated for Coretax import. Below is the list of settlement invoices:"
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.Columns(Me.col_TanggalTransaksi.Index).DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvItem.Columns(Me.col_TanggalDP.Index).DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvItem.Rows.Clear()
            Me.loadDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.ObjCD.SetNumber(Me.dgvItem, True)
        End Try
    End Sub

    'Load detail dari tanggal, hanya mengambil Invoice dengan item Invoice DP
    Private Sub loadDetail()
        Dim dta As New DseLaporanVAT2025_1TableAdapters.tblInvoice_InvoiceDPTableAdapter
        Dim rowdgv As DataGridViewRow
        Try
            dta.Connection.ConnectionString = core_module.getconstring
            dta.FillBy_DateFrom_To(Me.TBLInvoice_InvoiceDP, Me.DateFrom, Me.DateTo)
            If Me.TBLInvoice_InvoiceDP.Rows.Count > 0 Then
                For Each Me.row In Me.TBLInvoice_InvoiceDP
                    Me.dgvItem.Rows.Add()
                    rowdgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    With rowdgv
                        .Cells(Me.col_CustomerCode.Index).Value = row.Customer_Code
                        .Cells(Me.col_CustomerName.Index).Value = row.Customer_Name
                        .Cells(Me.col_NoTransaksi.Index).Value = row.No_Invoice
                        .Cells(Me.col_TanggalTransaksi.Index).Value = row.Tanggal_Invoice
                        .Cells(Me.col_NoDP.Index).Value = row.No_DP
                        .Cells(Me.col_TaxInvoiceDP.Index).Value = row.No_Seri_Faktur_DP & row.No_Faktur_Pajak_DP
                        .Cells(Me.col_TanggalDP.Index).Value = row.Tanggal_DP
                        .Cells(Me.col_SubTotal.Index).Value = Me.ObjCD.DigitGrouping(row.Sub_Total_Invoice, 0)
                        .Cells(Me.col_VAT.Index).Value = Me.ObjCD.DigitGrouping(row.Tax_Amount_Invoice, 0)
                        .Cells(Me.col_Total.index).value = Me.ObjCD.DigitGrouping(row.Total_Invoice, 0)
                        .Cells(Me.col_PrimaryTransaksi.Index).Value = row.Primary_Invoice
                        .Cells(Me.col_PrimaryDP.Index).Value = row.Primary_Invoice_DP
                        .Cells(Me.Col_PrimaryCustomer.Index).Value = row.Primary_Customer
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class