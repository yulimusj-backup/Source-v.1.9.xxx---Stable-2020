Option Explicit On
Option Strict Off
Public Class FrmTransaksiStockCard2_CheckProductQtyMonthlyClosing

    Dim objCD As New Core_Module.ClsCekData
    Dim tblDifferent As New DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.Different_DataDataTable
    Dim rowDifferent As DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.Different_DataRow

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Me.objCD.GetFirstDateofTheMonth(Me.DateTimePicker2.Value.Date) < Me.objCD.GetFirstDateofTheMonth(Me.DateTimePicker1.Value.Date) Then
                MsgBox("Invalid monthly closing period.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If MsgBox("Start?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            Me.Button1.Enabled = False
            Me.Button2.Enabled = False
            Me.ControlBox = False
            Me.Cursor = Cursors.WaitCursor
            FrmTransaksiStockCard2_CheckProductQtyMonthlyClosing.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckData()
        Dim dtaMonthlyClosing As New DseTransaksiStockCard2_CheckProductQtyMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim tblMonthlyClosing As New DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.tbltrans_purchase_closingDataTable
        Dim rowMonthlyClosing As DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.tbltrans_purchase_closingRow

        Dim dtaItemClosing As New DseTransaksiStockCard2_CheckProductQtyMonthlyClosingTableAdapters.Item_ClosingTableAdapter
        Dim tblItemClosing As New DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.Item_ClosingDataTable
        Dim rowItemClosing As DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.Item_ClosingRow

        Dim dtaItemQty As New DseTransaksiStockCard2_CheckProductQtyMonthlyClosingTableAdapters.Item_ProductQtyTableAdapter
        Dim tblItemQty As New DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.Item_ProductQtyDataTable
        Dim rowItemQty As DseTransaksiStockCard2_CheckProductQtyMonthlyClosing.Item_ProductQtyRow

        Try
            dtaMonthlyClosing.Connection.ConnectionString = Core_Module.GetConstring
            Me.tblDifferent.Rows.Clear()
            If Me.CheckBox1.Checked Then
                dtaMonthlyClosing.FillBy_Period(tblMonthlyClosing, Me.objCD.GetFirstDateofTheMonth(Me.DateTimePicker1.Value.Date), Me.objCD.GetFirstDateofTheMonth(Me.DateTimePicker2.Value.Date))
            Else
                dtaMonthlyClosing.Fill(tblMonthlyClosing)
            End If
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = tblMonthlyClosing.Count
            For Each rowMonthlyClosing In tblMonthlyClosing
                Me.ProgressBar1.Value += 1
                Me.Label3.Text = "Processing Document No: " & rowMonthlyClosing.No_Transaksi & ", data " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "."
                dtaItemClosing.Connection.ConnectionString = Core_Module.GetConstring
                dtaItemQty.Connection.ConnectionString = Core_Module.GetConstring
                dtaItemClosing.Fill(tblItemClosing, rowMonthlyClosing.Primary)
                dtaItemQty.Fill(tblItemQty, rowMonthlyClosing.Primary)
                tblItemClosing.DefaultView.RowFilter = ""
                tblItemClosing.DefaultView.Sort = "[Primary Product]"
                tblItemQty.DefaultView.RowFilter = ""
                tblItemQty.DefaultView.Sort = "[Primary Product]"
                Dim idx0 As Integer = 0
                Dim idx1 As Integer = 0
                Me.ProgressBar2.Value = 0
                Me.ProgressBar2.Maximum = tblItemClosing.DefaultView.Count
                While idx0 < tblItemClosing.DefaultView.Count
                    Me.ProgressBar2.Value += 1
                    Me.Label4.Text = "Checking Item Closing, data " & Me.ProgressBar2.Value & " of " & Me.ProgressBar2.Maximum & "."
                    rowItemClosing = tblItemClosing.DefaultView(idx0).Row
                    Dim Different As Boolean = False
                    Dim NotFound As Boolean = False
                    While idx1 < tblItemQty.DefaultView.Count
                        rowItemQty = tblItemQty.DefaultView(idx1).Row
                        If rowItemQty.Primary_Product = rowItemClosing.Primary_Product Then
                            If rowItemQty.Qty <> rowItemClosing.Ending_Qty Then
                                Different = True
                            End If
                            Exit While
                        ElseIf rowItemQty.Primary_Product > rowItemClosing.Primary_Product Then
                            NotFound = True
                            Exit While
                        End If
                        idx1 += 1
                    End While
                    If Different Or NotFound Then
                        rowDifferent = tblDifferent.NewDifferent_DataRow
                        rowDifferent.Document_No = rowMonthlyClosing.No_Transaksi
                        rowDifferent.Period = rowMonthlyClosing.Closing_Period
                        rowDifferent.Product_Code = rowItemClosing.Product_Code
                        rowDifferent.Item_Monthly_Closing = rowItemClosing.Ending_Qty
                        If NotFound Then
                            rowDifferent.Item_Product_Qty = 0
                        Else
                            rowDifferent.Item_Product_Qty = rowItemQty.Qty
                        End If
                        tblDifferent.Rows.Add(rowDifferent)
                    End If
                    idx0 += 1
                End While

                Me.ProgressBar2.Value = 0
                Me.ProgressBar2.Maximum = tblItemQty.DefaultView.Count
                While idx0 < tblItemQty.DefaultView.Count
                    Me.ProgressBar2.Value += 1
                    Me.Label4.Text = "Checking Item Qty, data " & Me.ProgressBar2.Value & " of " & Me.ProgressBar2.Maximum & "."
                    rowItemQty = tblItemQty.DefaultView(idx0).Row
                    Dim Different As Boolean = False
                    Dim NotFound As Boolean = False
                    While idx1 < tblItemClosing.DefaultView.Count
                        rowItemClosing = tblItemClosing.DefaultView(idx1).Row
                        If rowItemClosing.Primary_Product > rowItemQty.Primary_Product Then
                            NotFound = True
                            Exit While
                        End If
                        idx1 += 1
                    End While
                    If NotFound Then
                        rowDifferent = tblDifferent.NewDifferent_DataRow
                        rowDifferent.Document_No = rowMonthlyClosing.No_Transaksi
                        rowDifferent.Period = rowMonthlyClosing.Closing_Period
                        rowDifferent.Product_Code = rowItemQty.Product_Code
                        rowDifferent.Item_Product_Qty = rowItemQty.Qty
                        rowDifferent.Item_Monthly_Closing = 0
                        tblDifferent.Rows.Add(rowDifferent)
                    End If
                    idx0 += 1
                End While 
            Next
            Console.WriteLine("")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowOnGrid()
        Dim RowDgv As DataGridViewRow
        Try
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.Rows.Clear()
            Me.ProgressBar1.Value = 0
            Me.ProgressBar2.Value = 0
            Me.Label3.Text = "Showing different data on DataGridView..."
            Me.Label4.Text = ""
            For Each rowDifferent In tblDifferent
                RowDgv = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Add)
                RowDgv.Cells(Me.ColDocumentNo.Index).Value = rowDifferent.Document_No
                RowDgv.Cells(Me.ColItemClosing.Index).Value = Me.objCD.DigitGrouping(rowDifferent.Item_Monthly_Closing, 6)
                RowDgv.Cells(Me.ColItemProductQty.Index).Value = Me.objCD.DigitGrouping(rowDifferent.Item_Product_Qty, 6)
                RowDgv.Cells(Me.ColNo.Index).Value = RowDgv.Index + 1
                RowDgv.Cells(Me.ColPeriod.Index).Value = Format(rowDifferent.Period, "dd MMM yyyy")
                RowDgv.Cells(Me.ColProductCode.Index).Value = rowDifferent.Product_Code
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.objCD.UpdateThreading()
            Me.CheckData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.ShowOnGrid()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.ProgressBar1.Value = 0
            Me.ProgressBar2.Value = 0
            Me.Label3.Text = "Ready..."
            Me.Label4.Text = "Ready..."
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Me.ControlBox = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class