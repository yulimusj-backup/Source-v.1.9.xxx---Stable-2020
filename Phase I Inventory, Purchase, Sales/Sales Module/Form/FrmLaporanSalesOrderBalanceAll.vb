Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesOrderBalanceAll

    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 3
    Dim FormatYear As Integer = 1
    Dim FormatQty As Integer = 0

    Private Sub cetak()
        Dim report As New crpLaporanSalesOrderBalance2
        Dim dta As New dseLaporanSalesOrderBalance2TableAdapters.laporan_sales_orderbalance_view9bTableAdapter
        Dim TBL As dseLaporanSalesOrderBalance2.laporan_sales_orderbalance_view9bDataTable
        Dim row As dseLaporanSalesOrderBalance2.laporan_sales_orderbalance_view9bRow
        Dim ds As New dseLaporanSalesOrderBalance2
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowlaporan As dseLaporanSalesOrderBalance2.TblCetakRow

        Try
            TBL = dta.GetData(Me.dtpTo.Value, Me.dtpFrom.Value)
            For Each row In TBL.Rows
                rowlaporan = ds.TblCetak.NewRow()
                rowlaporan.CategoryCode = row.Category_Code
                rowlaporan.CategoryDesc = row.Category_Desc
                rowlaporan.ProductCode = row.Product_Code
                rowlaporan.Description = row.Description
                rowlaporan.SONo = row.SO_No
                rowlaporan.SODate = row.SO_Date
                rowlaporan.CustPONo = row.Cust_PO_No
                rowlaporan.QtySO = row.Qty_SO
                rowlaporan.OrderBalance = row.OrderBalance
                rowlaporan.UnitCode = row.Unit_Code
                rowlaporan.FreeQty = row.Free_Stock
                rowlaporan.CustomerCode = row.Customer_Code
                rowlaporan.CustomerName = row.Customer_Name
                ds.TblCetak.Rows.Add(rowlaporan)
            Next

            Dim rowcetak1 As dseLaporanSalesOrderBalance2.TblCetakRow
            Dim i As Integer = 1
            For Each rowcetak1 In ds.TblCetak.Rows
                If i = 1 Then
                    rowcetak1.Group = 1
                Else
                    Dim ketemu As Boolean = False
                    For x As Integer = 0 To ds.TblCetak.Rows.Count - 1
                        If ds.TblCetak(x).IsGroupNull = True Then
                            ketemu = False
                            Exit For
                        End If
                        If rowcetak1.CategoryCode = ds.TblCetak(x).CategoryCode And rowcetak1.ProductCode = ds.TblCetak(x).ProductCode Then
                            ketemu = True
                            rowcetak1.Group = ds.TblCetak(x).Group
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        rowcetak1.Group = i
                    End If
                End If
                i += 1
            Next

            ''hitung Total Qty Sales Order Balance
            Dim listGroup As New List(Of String)
            For Each rowcetak1 In ds.TblCetak.Rows
                If listGroup.Count = 0 Then
                    listGroup.Add(rowcetak1.Group)
                ElseIf listGroup.Contains(rowcetak1.Group) = False Then
                    listGroup.Add(rowcetak1.Group)
                End If
            Next
            Dim GroupString As String
            Dim TotalSalesOrderBalance As Double = 0
            Dim ListTotalSalesOrderBalance As New List(Of Double)
            For Each GroupString In listGroup
                For Each rowcetak1 In ds.TblCetak.Rows
                    If rowcetak1.Group = GroupString Then
                        TotalSalesOrderBalance += rowcetak1.OrderBalance
                    End If
                Next
                ListTotalSalesOrderBalance.Add(TotalSalesOrderBalance)
                TotalSalesOrderBalance = 0
            Next

            Dim index As Integer
            For Each rowcetak1 In ds.TblCetak.Rows
                index = listGroup.IndexOf(rowcetak1.Group)
                rowcetak1.TotalSalesOrderBalance = ListTotalSalesOrderBalance(index)
            Next
            ListTotalSalesOrderBalance = Nothing
            listGroup = Nothing

            ''Hitung Total Free Stock
            Dim listGroup2 As New List(Of String)
            For Each rowcetak1 In ds.TblCetak.Rows
                If listGroup2.Count = 0 Then
                    listGroup2.Add(rowcetak1.Group)
                ElseIf listGroup2.Contains(rowcetak1.Group) = False Then
                    listGroup2.Add(rowcetak1.Group)
                End If
            Next
            Dim GroupString2 As String
            Dim TotalFreeQty As Double = 0
            Dim ListTotalFreeQty As New List(Of Double)
            For Each GroupString2 In listGroup2
                For Each rowcetak1 In ds.TblCetak.Rows
                    If rowcetak1.Group = GroupString2 Then
                        TotalFreeQty += rowcetak1.FreeQty
                    End If
                Next
                ListTotalFreeQty.Add(TotalFreeQty)
                TotalFreeQty = 0
            Next

            Dim index2 As Integer
            For Each rowcetak1 In ds.TblCetak.Rows
                index2 = listGroup2.IndexOf(rowcetak1.Group)
                rowcetak1.TotalFreeQty = ListTotalFreeQty(index2)
            Next
            ListTotalFreeQty = Nothing
            listGroup2 = Nothing


            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYear)
            report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, Me.FormatQty)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            Me.crv1.ReportSource = report
            Me.crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ds.Dispose()
            dta.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            Me.cetak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class