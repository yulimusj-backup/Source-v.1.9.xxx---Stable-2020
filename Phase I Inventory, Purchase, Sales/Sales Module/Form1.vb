Public Class Form1
    Dim DTA As New DseClosing_Explain_OutTableAdapters.DataTable5TableAdapter
    Dim TBL As New DseClosing_Explain_Out.DataTable5DataTable
    Dim row As DseClosing_Explain_Out.DataTable5Row

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim counter As Double = 0

        Try
            DTA.Fill(Me.TBL)
            While counter < Me.TBL.Rows.Count
                row = Me.TBL.Rows(counter)
                If row.IsQty_InvoiceNull = True Then
                    row.Qty_Invoice = 0
                    row.Selisih = row.Qty_Out
                End If
                If row.Selisih = 0 Then
                    Me.TBL.Rows.Remove(row)
                Else
                    counter += 1
                End If
            End While

            Me.DataGridView1.DataSource = Me.TBL.DefaultView
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "information")
        End Try
    End Sub
End Class