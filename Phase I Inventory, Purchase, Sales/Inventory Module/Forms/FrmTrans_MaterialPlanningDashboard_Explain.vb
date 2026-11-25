Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_Explain
    Public ExplainAvg As String = ""
    Public ExplainSafetyStock As String = ""
    Public ExplainLeadTime As String = ""

    Public ExplainOptimumStock As String = ""

    Public pSafateyStock As Boolean = False
    Public pLeadTime As Boolean = False

    Public pOptimumStock As Boolean = False

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vPrimaryProduct As Double = -1
    Dim DefPrec As Integer = 0

    Dim StartDate, EndDate As Date
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct

    Private Function FormatQtyAvgPerMonth(ByVal vStringAsal As String) As String
        Dim TotalUsage, JmlBulan, Avg As Decimal
        Dim Index1 As Integer = 0
        Dim Index2 As Integer = 0

        Try
            Index1 = vStringAsal.IndexOf("/")
            Index2 = vStringAsal.IndexOf("=")
            TotalUsage = vStringAsal.Substring(0, Index1 - 1)
            JmlBulan = vStringAsal.Substring(Index1 + 1, Index2 - Index1 - 1)
            Avg = vStringAsal.Substring(Index2 + 1)

            Return Me.ObjCD.DigitGrouping(TotalUsage, DefPrec) + " / " + JmlBulan.ToString + " = " + Me.ObjCD.DigitGrouping(Avg, DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return vStringAsal
        End Try
    End Function

    Private Sub FrmTrans_MaterialPlanningDashboard_Explain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim arrayBarisAvg As String()
        Dim arrayPeriodAvg As String()
        Dim rowdgv As DataGridViewRow

        Try
            DefPrec = Me.ObjProduct.GetAngkaDecimal(Me.vPrimaryProduct)
            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            arrayBarisAvg = ExplainAvg.Split(vbNewLine)
            Me.txtExplain.Text = ExplainAvg
            If arrayBarisAvg.Length >= 9 Then
                If arrayBarisAvg(3).Trim = "0" Then
                    Me.cboIncludeSales.Checked = False
                ElseIf arrayBarisAvg(3).Trim = "1" Then
                    Me.cboIncludeSales.Checked = True
                Else
                    MsgBox("The Explanation of Average Monthly Demand is corrupted, it will not be displayed correctly", MsgBoxStyle.Exclamation, "Exclamation")
                    Exit Sub
                End If
                If arrayBarisAvg(4).Trim = "0" Then
                    Me.cboIncludeBundling.Checked = False
                ElseIf arrayBarisAvg(4).Trim = "1" Then
                    Me.cboIncludeBundling.Checked = True
                Else
                    MsgBox("The Explanation of Average Monthly Demand is corrupted, it will not be displayed correctly", MsgBoxStyle.Exclamation, "Exclamation")
                    Exit Sub
                End If
                If arrayBarisAvg(5).Trim = "0" Then
                    Me.cboIncludeAdjustment.Checked = False
                ElseIf arrayBarisAvg(5).Trim = "1" Then
                    Me.cboIncludeAdjustment.Checked = True
                Else
                    MsgBox("The Explanation of Average Monthly Demand is corrupted, it will not be displayed correctly", MsgBoxStyle.Exclamation, "Exclamation")
                    Exit Sub
                End If
                For i As Integer = 6 To arrayBarisAvg.Length - 3
                    arrayPeriodAvg = arrayBarisAvg(i).Split(";")
                    Me.dgvData.Rows.Add()
                    rowdgv = Me.dgvData.Rows(Me.dgvData.Rows.Count - 1)
                    rowdgv.Cells(Me.ColExplain_Periode.Index).Value = arrayPeriodAvg(0)
                    rowdgv.Cells(Me.ColExplain_Sales.Index).Value = Me.ObjCD.DigitGrouping(arrayPeriodAvg(1), DefPrec)
                    rowdgv.Cells(Me.ColExplain_Bundling.Index).Value = Me.ObjCD.DigitGrouping(arrayPeriodAvg(2), DefPrec)
                    rowdgv.Cells(Me.ColExplain_Adjustment.Index).Value = Me.ObjCD.DigitGrouping(arrayPeriodAvg(3), DefPrec)
                    rowdgv.Cells(Me.ColHistorical.Index).Value = Me.ObjCD.DigitGrouping(arrayPeriodAvg(4), DefPrec)
                    rowdgv.Cells(Me.ColExplain_Total.Index).Value = Me.ObjCD.DigitGrouping(arrayPeriodAvg(5), DefPrec)
                Next
                Me.txtExplain.Text = "Total Usage = " & Me.ObjCD.DigitGrouping(arrayBarisAvg(arrayBarisAvg.Length - 2), DefPrec) & vbNewLine & _
                "Avg. Per Month = " & Me.FormatQtyAvgPerMonth(arrayBarisAvg(arrayBarisAvg.Length - 1))

                If Me.pSafateyStock = True Then
                    arrayBarisAvg = ExplainSafetyStock.Split(vbNewLine)
                    If arrayBarisAvg.Length < 2 Then
                        MsgBox("The Explanation of Safety Stock is corrupted, it will not be displayed correctly", MsgBoxStyle.Exclamation, "Exclamation")
                        Exit Sub
                    Else
                        Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & vbNewLine & "Explain Safety Stock" & vbNewLine & "Safety Time: " & arrayBarisAvg(0) & " days."
                        arrayPeriodAvg = arrayBarisAvg(1).Split(";")
                        Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & "Safety Stock: " & Me.ObjCD.DigitGrouping(arrayPeriodAvg(0), DefPrec) & " * " & arrayPeriodAvg(1) & " / " & arrayPeriodAvg(2) & " = " & Me.ObjCD.DigitGrouping(arrayPeriodAvg(3), DefPrec)
                    End If
                Else
                    Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & vbNewLine & "Safety Stock is User Defined"
                End If

                If Me.pLeadTime = True Then
                    arrayBarisAvg = ExplainLeadTime.Split(vbNewLine)
                    If arrayBarisAvg.Length < 2 Then
                        MsgBox("The Explanation of Lead Time is corrupted, it will not be displayed correctly", MsgBoxStyle.Exclamation, "Exclamation")
                        Exit Sub
                    Else
                        Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & vbNewLine & "Explain Lead Stock" & vbNewLine & "Lead Time: " & arrayBarisAvg(0) & " days."
                        arrayPeriodAvg = arrayBarisAvg(1).Split(";")
                        Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & "Lead Stock: " & Me.ObjCD.DigitGrouping(arrayPeriodAvg(0), DefPrec) & " * " & arrayPeriodAvg(1) & " / " & arrayPeriodAvg(2) & " = " & Me.ObjCD.DigitGrouping(arrayPeriodAvg(3), DefPrec)
                    End If
                Else
                    Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & vbNewLine & "Lead Stock is User Defined"
                End If

                If Me.pOptimumStock = True Then
                    arrayBarisAvg = ExplainOptimumStock.Split(vbNewLine)
                    If arrayBarisAvg.Length < 2 Then
                        MsgBox("The Explanation of Optimum Stock is corrupted, it will not be displayed correctly", MsgBoxStyle.Exclamation, "Exclamation")
                        Exit Sub
                    Else
                        Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & vbNewLine & "Explain Optimum Stock" & vbNewLine & "Optimum Stock Time: " & arrayBarisAvg(0) & " days."
                        arrayPeriodAvg = arrayBarisAvg(1).Split(";")
                        Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & "Optimum Stock: " & Me.ObjCD.DigitGrouping(arrayPeriodAvg(0), DefPrec) & " * " & arrayPeriodAvg(1) & " / " & arrayPeriodAvg(2) & " = " & Me.ObjCD.DigitGrouping(arrayPeriodAvg(3), DefPrec)
                    End If
                Else
                    Me.txtExplain.Text = Me.txtExplain.Text & vbNewLine & vbNewLine & "Optimum Stock is User Defined"
                End If

            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub dgvData_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvData.CellBeginEdit
        Try
            e.Cancel = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub dgvData_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellClick
        Dim ij As New FrmTrans_MaterialPlanningDashboard_Explain2
        Dim rowDGV As DataGridViewRow
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Unit As String = ""
        Dim ObjUnit As New Core_Module.ClsMasterUnit

        Try
            Unit = ObjUnit.GetUnitCode(ObjProd.GetPrimaryUnit(Me.vPrimaryProduct))

            If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
                If Me.dgvData.Rows.Count > 0 Then
                    rowDGV = Me.dgvData.Rows(e.RowIndex)
                    Me.StartDate = CDate("1/" + rowDGV.Cells(Me.ColExplain_Periode.Index).Value)
                    Me.EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.StartDate)
                End If
                If e.ColumnIndex = Me.ColExplain_Adjustment.Index Then
                    If CDec(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                        Dim DTA3 As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.DataTable3TableAdapter
                        Dim TBL3 As New DseTrans_MaterialPlanningDashboard_Data.DataTable3DataTable
                        Dim row3 As DseTrans_MaterialPlanningDashboard_Data.DataTable3Row
                        Dim TipeTrans1 As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust

                        DTA3.Fill(TBL3, TipeTrans1, EndDate, StartDate, Me.vPrimaryProduct)

                        ij.dgvData.Rows.Clear()
                        ij.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColExpiry.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColQty.DefaultCellStyle.Format = "N" + ObjProd.GetAngkaDecimal(vPrimaryProduct).ToString
                        ij.ColQty.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColNoTrans.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColNo.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColExpiry.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColDate.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColCustomer.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColBatchNo.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColCustomer.Visible = False

                        For Each row3 In TBL3.Rows
                            ij.dgvData.Rows.Add()
                            rowDGV = ij.dgvData.Rows(ij.dgvData.Rows.Count - 1)
                            With rowDGV
                                .Cells(ij.ColBatchNo.Index).Value = row3.Batch_No
                                .Cells(ij.ColDate.Index).Value = row3.Tanggal_Transaksi
                                .Cells(ij.ColExpiry.Index).Value = row3.Expiry_Date
                                .Cells(ij.ColNo.Index).Value = ij.dgvData.Rows.Count
                                .Cells(ij.ColNoTrans.Index).Value = row3.No_Transaksi
                                .Cells(ij.ColQty.Index).Value = row3.jml
                                .Cells(ij.ColUnit.Index).Value = Unit
                            End With
                        Next
                        If ObjProd.ApakahBatchNoControlled(Me.vPrimaryProduct) <> True Then
                            ij.ColBatchNo.Visible = False
                            ij.ColExpiry.Visible = False
                        Else
                            If ObjProd.ApakahExpiryDateControlled(Me.vPrimaryProduct) <> True Then
                                ij.ColExpiry.Visible = False
                            End If
                        End If

                        ij.TabControl1.TabPages.Remove(ij.TabPage_SalesReturn)
                        ij.TabPage_Sales.Text = "Adjustment"
                        ij.ShowDialog()
                        DTA3.Dispose()
                        TBL3.Dispose()
                    End If
                ElseIf e.ColumnIndex = Me.ColExplain_Bundling.Index Then
                    If CDec(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                        Dim DTA2 As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.DataTable2TableAdapter
                        Dim TBL2 As New DseTrans_MaterialPlanningDashboard_Data.DataTable2DataTable
                        Dim row2 As DseTrans_MaterialPlanningDashboard_Data.DataTable2Row

                        DTA2.Fill(TBL2, EndDate, StartDate, Me.vPrimaryProduct)

                        ij.dgvData.Rows.Clear()
                        ij.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColExpiry.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColQty.DefaultCellStyle.Format = "N" + ObjProd.GetAngkaDecimal(vPrimaryProduct).ToString
                        ij.ColQty.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColNoTrans.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColNo.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColExpiry.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColDate.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColCustomer.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColBatchNo.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColCustomer.Visible = False

                        For Each row2 In TBL2.Rows
                            ij.dgvData.Rows.Add()
                            rowDGV = ij.dgvData.Rows(ij.dgvData.Rows.Count - 1)
                            With rowDGV
                                .Cells(ij.ColBatchNo.Index).Value = row2.Batch_No
                                .Cells(ij.ColDate.Index).Value = row2.Tanggal_Transaksi
                                .Cells(ij.ColExpiry.Index).Value = row2.Expiry_Date
                                .Cells(ij.ColNo.Index).Value = ij.dgvData.Rows.Count
                                .Cells(ij.ColNoTrans.Index).Value = row2.No_Transaksi
                                .Cells(ij.ColQty.Index).Value = row2.jml
                                .Cells(ij.ColUnit.Index).Value = Unit
                            End With
                        Next
                        If ObjProd.ApakahBatchNoControlled(Me.vPrimaryProduct) <> True Then
                            ij.ColBatchNo.Visible = False
                            ij.ColExpiry.Visible = False
                        Else
                            If ObjProd.ApakahExpiryDateControlled(Me.vPrimaryProduct) <> True Then
                                ij.ColExpiry.Visible = False
                            End If
                        End If
                        ij.TabControl1.TabPages.Remove(ij.TabPage_SalesReturn)
                        ij.TabPage_Sales.Text = "Bundling"
                        ij.ShowDialog()
                        DTA2.Dispose()
                        TBL2.Dispose()
                    End If
                ElseIf e.ColumnIndex = Me.ColExplain_Sales.Index Then
                    If CDec(Me.dgvData.Item(e.ColumnIndex, e.RowIndex).Value) > 0 Then
                        Dim DTA1 As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.DataTable1TableAdapter
                        Dim TBL1 As New DseTrans_MaterialPlanningDashboard_Data.DataTable1DataTable
                        Dim row1 As DseTrans_MaterialPlanningDashboard_Data.DataTable1Row
                        Dim TipeTrans1 As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales

                        DTA1.Fill(TBL1, TipeTrans1, EndDate, StartDate, Me.vPrimaryProduct)

                        ij.dgvData.Rows.Clear()
                        ij.ColDate.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColExpiry.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColQty.DefaultCellStyle.Format = "N" + ObjProd.GetAngkaDecimal(vPrimaryProduct).ToString
                        ij.ColQty.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColNoTrans.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColNo.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColExpiry.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColDate.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColCustomer.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColBatchNo.SortMode = DataGridViewColumnSortMode.NotSortable

                        For Each row1 In TBL1.Rows
                            ij.dgvData.Rows.Add()
                            rowDGV = ij.dgvData.Rows(ij.dgvData.Rows.Count - 1)
                            With rowDGV
                                .Cells(ij.ColBatchNo.Index).Value = row1.Batch_No
                                .Cells(ij.ColCustomer.Index).Value = row1.Customer_Name
                                .Cells(ij.ColDate.Index).Value = row1.Tanggal_Transaksi
                                .Cells(ij.ColExpiry.Index).Value = row1.Expiry_Date
                                .Cells(ij.ColNo.Index).Value = ij.dgvData.Rows.Count
                                .Cells(ij.ColNoTrans.Index).Value = row1.No_Transaksi
                                .Cells(ij.ColQty.Index).Value = row1.jml
                                .Cells(ij.ColUnit.Index).Value = Unit
                            End With
                        Next
                        If ObjProd.ApakahBatchNoControlled(Me.vPrimaryProduct) <> True Then
                            ij.ColBatchNo.Visible = False
                            ij.ColExpiry.Visible = False
                        Else
                            If ObjProd.ApakahExpiryDateControlled(Me.vPrimaryProduct) <> True Then
                                ij.ColExpiry.Visible = False
                            End If
                        End If

                        'isikan data sales return
                        TBL1.Rows.Clear()
                        TipeTrans1 = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                        DTA1.Fill(TBL1, TipeTrans1, EndDate, StartDate, Me.vPrimaryProduct)

                        ij.dgvDataSR.Rows.Clear()
                        ij.ColSR_Date.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColSR_Expiry.DefaultCellStyle.Format = "dd MMM yyyy"
                        ij.ColSR_Qty.DefaultCellStyle.Format = "N" + ObjProd.GetAngkaDecimal(vPrimaryProduct).ToString
                        ij.ColSR_Qty.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColSR_DocumentNo.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColSR_No.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColSR_Expiry.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColSR_Date.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColSR_Customer.SortMode = DataGridViewColumnSortMode.NotSortable
                        ij.ColSR_BatchNo.SortMode = DataGridViewColumnSortMode.NotSortable

                        For Each row1 In TBL1.Rows
                            ij.dgvDataSR.Rows.Add()
                            rowDGV = ij.dgvDataSR.Rows(ij.dgvDataSR.Rows.Count - 1)
                            With rowDGV
                                .Cells(ij.ColSR_BatchNo.Index).Value = row1.Batch_No
                                .Cells(ij.ColSR_Customer.Index).Value = row1.Customer_Name
                                .Cells(ij.ColSR_Date.Index).Value = row1.Tanggal_Transaksi
                                .Cells(ij.ColSR_Expiry.Index).Value = row1.Expiry_Date
                                .Cells(ij.ColSR_No.Index).Value = ij.dgvDataSR.Rows.Count
                                .Cells(ij.ColSR_DocumentNo.Index).Value = row1.No_Transaksi
                                .Cells(ij.ColSR_Qty.Index).Value = row1.jml
                                .Cells(ij.ColSR_Unit.Index).Value = Unit
                            End With
                        Next
                        If ObjProd.ApakahBatchNoControlled(Me.vPrimaryProduct) <> True Then
                            ij.ColSR_BatchNo.Visible = False
                            ij.ColSR_Expiry.Visible = False
                        Else
                            If ObjProd.ApakahExpiryDateControlled(Me.vPrimaryProduct) <> True Then
                                ij.ColSR_Expiry.Visible = False
                            End If
                        End If

                        ij.ShowDialog()
                        DTA1.Dispose()
                        TBL1.Dispose()
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
            ObjProd = Nothing
            ObjUnit = Nothing
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class