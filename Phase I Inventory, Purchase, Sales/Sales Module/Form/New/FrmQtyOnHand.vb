Option Explicit On
Option Strict Off
Public Class FrmQtyOnHand
    Public Tbl As New DseImport_eFaktur.tbl_temporaryDataTable
    Public PrimaryInvListMinus As New List(Of Double)
    Public PrimaryCodeListMinus As New List(Of Double)
    Public BackColorAll As Color
    Dim row As DseImport_eFaktur.tbl_temporaryRow
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmQtyOnHand_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvHeader.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ShowToDgvItem()
            Me.ShowToDgvHeader(Me.dgvItem.Item(Me.ColPrimary.Index, 0).Value)
            Me.Text = "Information qty stock minus for this code no"
            Me.ColPrimary.Visible = False
            Me.ColH_Primary.Visible = False
            Me.dgvItem.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowToDgvItem()
        Dim rd As DataGridViewRow
        Try

            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To Tbl.DefaultView.Count - 1
                row = Tbl.DefaultView(i).Row
                If PrimaryInvListMinus.Contains(row.primary) = True And PrimaryCodeListMinus.Contains(row.primary_code_smess) = True And _
                 row.is_qty_minus = True Then
                    With row
                        rd = Me.dgvItem.Rows(Me.dgvItem.Rows.Add)
                        'rd.Cells(Me.col_No.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColPrimary.Index).Value = .primary
                        rd.Cells(Me.ColCode.Index).Value = .prod_code
                        rd.Cells(Me.ColCodeNoSMESS.Index).Value = .prod_code_smess
                        rd.Cells(Me.ColDesc.Index).Value = .prod_desc
                        rd.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.prod_qty, .prod_dec)
                        rd.Cells(Me.ColUnit.Index).Value = .prod_unit
                        rd.Cells(Me.ColQtyOn.Index).Value = Me.ObjCD.DigitGrouping(.qty_on_hand, .prod_dec)
                    End With
                End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.ObjCD.SetNumber(Me.dgvHeader)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowToDgvHeader(ByVal vPrimary As Double)
        Dim rd As DataGridViewRow

        Try
            Dim ListOfPrimary As New List(Of Double)
            Me.dgvHeader.Rows.Clear()
            Tbl.DefaultView.RowFilter = "[primary]=" & vPrimary
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For i As Integer = 0 To Tbl.DefaultView.Count - 1
                row = Tbl.DefaultView(i).Row
                If ListOfPrimary.Contains(vPrimary) = False Then
                    ListOfPrimary.Add(vPrimary)
                    With row
                        rd = Me.dgvHeader.Rows(Me.dgvHeader.Rows.Add)
                        'rd.Cells(Me.colNo.Index).Value = rd.Index + 1
                        rd.Cells(Me.ColH_InvoiceNo.Index).Value = .invoice_no
                        rd.Cells(Me.ColH_InvoiceDate.Index).Value = Format(.invoice_date, "dd MMM yyyy")
                        rd.Cells(Me.ColH_Customer.Index).Value = .customer
                        rd.Cells(Me.ColH_TaxInvoiceNo.Index).Value = .tax_invoice_no_1 & .tax_invoice_no_2
                        rd.Cells(Me.ColH_Primary.Index).Value = .primary
                    End With
                End If
            Next
            Me.dgvHeader.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvHeader_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHeader.CellClick
        
    End Sub

    Private Sub dgvHeader_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvHeader.SelectionChanged
        
    End Sub

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        'Try
        '    If Me.dgvItem.SelectedCells.Count > 0 Then
        '        Me.ShowToDgvHeader(Me.dgvItem(Me.ColPrimary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value)
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.ShowToDgvHeader(Me.dgvItem(Me.ColPrimary.Index, Me.dgvItem.SelectedCells(0).RowIndex).Value)
            End If
            'If Me.dgvItem.SelectedCells.Count > 0 Then

            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class