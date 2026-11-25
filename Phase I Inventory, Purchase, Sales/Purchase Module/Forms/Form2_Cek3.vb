Public Class Form2_Cek3
    Public PeriodClosing As Date = Today
    Dim LangsungClose As Boolean = False
    Dim TBLHasil As New DseCek2.DataTable1DataTable

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Dim PrimaryClosing As Double = -1
        Dim DTA As New DseCek2TableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim TBL As New DseCek2.tbltrans_purchase_closing_itemDataTable
        Dim row As DseCek2.tbltrans_purchase_closing_itemRow
        Dim DTA1 As New DseCek2TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBL1 As New DseCek2.tbltrans_purchase_closing_item_detailDataTable
        Dim row1 As DseCek2.tbltrans_purchase_closing_item_detailRow
        Dim Retval As Nullable(Of Double)
        Dim Cocok As Boolean = False
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
        Dim Qty, Amount As Decimal
        Dim rowHasil As DseCek2.DataTable1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            Me.TBLHasil.Rows.Clear()

            'cari primary closing
            Retval = DTA.ScalarQuery_GetPrimaryClosing(Me.PeriodClosing)
            If Retval.HasValue = False Then
                Me.LangsungClose = True
                Me.Close()
            End If

            PrimaryClosing = Retval.Value

            'isikan data item - item detail
            DTA.Fill(TBL, PrimaryClosing)
            DTA1.Fill(TBL1, PrimaryClosing)

            'cocokan qty begining balance, qty in, qty out, qty ending, amount begining balance, amount in, amount out, amount ending
            Me.prg1.Value = 0
            Me.prg1.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                Cocok = True
                Me.prg1.Value += 1
                If Me.prg1.Value > 25 Then
                    If Math.IEEERemainder(Me.prg1.Value, 100) = 0 Then
                      Me.LblProgress.Text = "Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString + " - " + row.Product_Code
                    ElseIf Me.prg1.Value = Me.prg1.Maximum Then
                      Me.LblProgress.Text = "Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString + " - " + row.Product_Code
                    End If
                Else
                   Me.LblProgress.Text = "Data " + Me.prg1.Value.ToString + " of " + Me.prg1.Maximum.ToString + " - " + row.Product_Code
                End If


                TipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
                TBL1.DefaultView.RowFilter = "[primary code no] = " + row.Primary_Product.ToString + " AND [tipe item detail] = " + CInt(TipeItemDetail).ToString
                Qty = Me.GetQtyAmount(TBL1.qtyColumn.ColumnName, TBL1.DefaultView)
                If row.Begin_Qty <> Qty Then
                    Cocok = False
                End If
                If Cocok = True Then
                    Amount = Me.GetQtyAmount(TBL1.amountColumn.ColumnName, TBL1.DefaultView)
                    If row.Begin_Amount <> Amount Then
                        Cocok = False
                    End If
                End If
                If Cocok = True Then
                    TipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                    TBL1.DefaultView.RowFilter = "[primary code no] = " + row.Primary_Product.ToString + " AND [tipe item detail] = " + CInt(TipeItemDetail).ToString
                    Qty = Me.GetQtyAmount(TBL1.qtyColumn.ColumnName, TBL1.DefaultView)
                    If row.In_Qty <> Qty Then
                        Cocok = False
                    End If
                End If
                If Cocok = True Then
                    Amount = Me.GetQtyAmount(TBL1.amountColumn.ColumnName, TBL1.DefaultView)
                    If row.In_Amount <> Amount Then
                        Cocok = False
                    End If
                End If
                If Cocok = True Then
                    TipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                    TBL1.DefaultView.RowFilter = "[primary code no] = " + row.Primary_Product.ToString + " AND [tipe item detail] = " + CInt(TipeItemDetail).ToString
                    Qty = Me.GetQtyAmount(TBL1.qtyColumn.ColumnName, TBL1.DefaultView)
                    If row.Out_Qty <> Qty Then
                        Cocok = False
                    End If
                End If
                If Cocok = True Then
                    Amount = Me.GetQtyAmount(TBL1.amountColumn.ColumnName, TBL1.DefaultView)
                    If row.Out_Amount <> Amount Then
                        Cocok = False
                    End If
                End If
                If Cocok = True Then
                    TipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
                    TBL1.DefaultView.RowFilter = "[primary code no] = " + row.Primary_Product.ToString + " AND [tipe item detail] = " + CInt(TipeItemDetail).ToString
                    Qty = Me.GetQtyAmount(TBL1.qtyColumn.ColumnName, TBL1.DefaultView)
                    If row.Ending_Qty <> Qty Then
                        Cocok = False
                    End If
                End If
                If Cocok = True Then
                    Amount = Me.GetQtyAmount(TBL1.amountColumn.ColumnName, TBL1.DefaultView)
                    If row.Ending_Amount <> Amount Then
                        Cocok = False
                    End If
                End If

                If Cocok = False Then
                    rowHasil = Me.TBLHasil.NewDataTable1Row
                    rowHasil.Primary_Product = row.Primary_Product
                    rowHasil.Code_No = row.Product_Code
                    rowHasil.Description = row.Description
                    TBLHasil.Rows.Add(rowHasil)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetQtyAmount(ByVal vColumnName As String, ByVal vDV As Data.DataView) As Decimal
        Dim row1 As DseCek2.tbltrans_purchase_closing_item_detailRow
        Dim RetVal As Double
        Try
            For Each DV As DataRowView In vDV
                row1 = DV.Row
                RetVal += row1.Item(vColumnName)
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Dim rowDGV As DataGridViewRow

        Try
            If Me.TBLHasil.Rows.Count = 0 Then
                MsgBox("Data check, no anomalies found.", MsgBoxStyle.Information, "Information")
                Me.LangsungClose = True
                Me.Close()
            Else
                For Each rowHasil As DseCek2.DataTable1Row In Me.TBLHasil.Rows
                    Me.DataGridView1.Rows.Add()
                    rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.DataGridView1.Rows.Count
                    rowDGV.Cells(Me.ColCodeNo.Index).Value = rowHasil.Code_No
                    rowDGV.Cells(Me.ColDesc.Index).Value = rowHasil.Description
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_Cek3_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungClose = False Then
                If MsgBox("Close form?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_Cek3_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DataGridView1.Rows.Clear()

            Me.bgw1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class