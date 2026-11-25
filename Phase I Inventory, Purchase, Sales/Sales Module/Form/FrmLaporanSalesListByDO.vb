Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesListByDO
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0

    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatHarga As Integer = 2

    Dim DS As New DseLaporanSalesListByDO
    Dim ReportByDO As New CrpLaporanSalesListByDO2
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency

    Dim RowView As DataRowView
    Dim TBL2 As DseLaporanSalesListByDO.DataTable2DataTable = DS.DataTable2
    Dim TBL3 As DseLaporanSalesListByDO.DataTable3DataTable = DS.DataTable3
    Dim TBL4 As DseLaporanSalesListByDO.tbltrans_doDataTable = DS.tbltrans_do
    Dim TBL5 As DseLaporanSalesListByDO.tbltransaksi_stockinout_salesDataTable = DS.tbltransaksi_stockinout_sales
    Dim DTA_StockOutSales As New DseLaporanSalesListByDOTableAdapters.tbltransaksi_stockinout_salesTableAdapter

    Dim TBLGroupByDO As DseLaporanSalesListByDO.Group_By_DODataTable = DS.Group_By_DO
    Dim StartPeriode, EndPeriod, StartPeriodInvoice, EndPeriodInvoice As Date
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim DefPrecIDR As Integer = 0

    Private Sub GenerateByDO()
        Dim rowDO As DseLaporanSalesListByDO.tbltrans_doRow
        Dim row2 As DseLaporanSalesListByDO.Group_By_DORow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim cmdDO As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByDOView As DataView = Me.TBLGroupByDO.DefaultView
        Dim Lanjut As Boolean = False
        Dim DTADO_Detail As New DseLaporanSalesListByDOTableAdapters.DataTable4TableAdapter
        Dim TBLDO_Detail As New DseLaporanSalesListByDO.DataTable4DataTable
        Dim TBL_ExchangeRate As New DseLaporanSalesListByDO.DataTable5DataTable
        Dim RateString As String = ""
        Dim Filter_DO As String = ""
        'Dim row_StockOut As DseLaporanSalesListByDO.tbltransaksi_stockinout_salesRow
        Dim DtaInv As New DseLaporanSalesListByDOTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim TblInv As New DseLaporanSalesListByDO.tbltrans_salesinvoice_itemDataTable
        Dim PrimaryMaxDO As Double = -1
        Dim PrimaryMinDO As Double = -1

        Try

            'modifikasi sales amount yang ditampilkan pada DO adalah sales berdasarkan SO
            'dengan melibatkan perhitungan konversi mata uang
            DTADO_Detail.Connection.ConnectionString = Core_Module.GetConstring
            DTADO_Detail.Fill(TBLDO_Detail, Me.dtpSampai.Value.Date, Me.dtpDari.Value.Date)
            Me.SetConvertionRate(TBLDO_Detail, TBL_ExchangeRate)
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            RateString = Me.GetRateString(TBL_ExchangeRate)

            ''OPTIMASI LAPORAN SALES LIST BY DO YANG LAMBAT
            ''fungsi ini ditutup karena lambat, untuk tanggal stock out langsung include saat fill TBL4
            'isikan data tanggal stock out, karena nilai cost of sales ditentukan dari tanggal stock out
            'DTA_StockOutSales.Connection.ConnectionString = Core_Module.GetConstring
            'Me.DTA_StockOutSales.Fill(Me.TBL5)
            'For Each rowDO In TBL4.Rows
            '    If rowDO.Primary_Stock_Out_Sales <> -1 Then
            '        Filter_DO = "Primary =" + rowDO.Primary_Stock_Out_Sales.ToString
            '        Me.TBL5.DefaultView.RowFilter = Filter_DO
            '        If Me.TBL5.DefaultView.Count > 0 Then
            '            row_StockOut = Me.TBL5.DefaultView(0).Row
            '            rowDO.Tanggal_Stock_Out = row_StockOut.Tanggal_Transaksi
            '        End If
            '    Else
            '        rowDO.Tanggal_Stock_Out = #1/1/1950#
            '    End If
            'Next

            For Each rowDO In TBL4.Rows
                If PrimaryMaxDO = -1 And PrimaryMinDO = -1 Then
                    PrimaryMaxDO = rowDO.Primary_DO
                    PrimaryMinDO = rowDO.Primary_DO
                End If
                If PrimaryMaxDO < rowDO.Primary_DO Then
                    PrimaryMaxDO = rowDO.Primary_DO
                End If
                If PrimaryMinDO > rowDO.Primary_DO Then
                    PrimaryMinDO = rowDO.Primary_DO
                End If

                ''OPTIMASI LAPORAN SALES LIST BY DO YANG LAMBAT
                ''karena tidak semua DO ada stock out sales nya, maka jika primary stock out sales = -1
                ''berarti tanggal stock out sales (null) di ganti menjadi 1/1/1950 sama seperti fungsi sebelumnya
                If rowDO.Primary_Stock_Out_Sales = -1 Then
                    rowDO.Tanggal_Stock_Out = #1/1/1950#
                End If

                Lanjut = False
                'filter berdasarkan cbo
                If Me.rboAll.Checked = True Then
                    Lanjut = True
                ElseIf Me.rboInvoiceDO.Checked = True And rowDO.Primary_Invoice <> -1 Then
                    'hanya yang memiliki invoice
                    Lanjut = True
                    If Lanjut = True Then
                        If Me.chkNonPendingDO.Checked = True And rowDO.Primary_Stock_Out_Sales = -1 Then
                            Lanjut = False
                        Else
                            Lanjut = True
                        End If
                    End If
                ElseIf Me.rboNotInvoiceDO.Checked = True And rowDO.Primary_Invoice = -1 Then
                    'hanya yang tidak memiliki invoice
                    Lanjut = True
                    If Lanjut = True Then
                        If Me.chkNonPendingDO.Checked = True And rowDO.Primary_Stock_Out_Sales = -1 Then
                            Lanjut = False
                        Else
                            Lanjut = True
                        End If
                    End If
                ElseIf Me.rboDONotStockedOut.Checked = True And rowDO.Primary_Invoice = -1 And rowDO.Primary_Stock_Out_Sales = -1 And _
                Me.chkNonPendingDO.Checked = False Then 'Update SMESS 2026
                    Lanjut = True 'Update SMESS 2026
                Else
                    Lanjut = False
                End If



                If Lanjut = True Then
                    'reset nilai penjualan per do
                    SalesAmount = 0
                    'bentuk filter
                    SalesAmount = Me.GetSalesAmount(rowDO.DO_No, _
                    TBL_ExchangeRate, TBLDO_Detail)
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Do No] ='" + rowDO.DO_No + "'"
                    TBLGroupByDOView.RowFilter = cmd
                    If TBLGroupByDOView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupByDO.Rows.Add(rowDO.DO_No, rowDO.Tanggal_Transaksi, SalesAmount, 0, _
                        rowDO.No_Invoice, 0, rowDO.Customer_Code + " - " + rowDO.Customer_Name, rowDO.Tanggal_Stock_Out, rowDO.Primary_DO)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        'masukkan data total sales
                        TotalSales += SalesAmount
                    Else
                        row2 = TBLGroupByDOView.Item(0).Row
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        row2.Sales_Amount += SalesAmount
                        'masukkan data total sales
                        TotalSales += SalesAmount
                    End If
                End If
            Next

            'hitung cost of sales
            Dim row3 As DseLaporanSalesListByDO.DataTable3Row
            If Me.TBL3.Rows.Count > 0 Then
                row3 = Me.TBL3.DefaultView(0).Row
                DtaInv.Connection.ConnectionString = Core_Module.GetConstring
                DtaInv.Fill(TblInv, Me.ObjCD.GetLasDateOfTheMonth(row3.Closing_Period), PrimaryMaxDO, PrimaryMinDO)
            End If
            For Each row2 In Me.TBLGroupByDO.Rows
                With row2
                    'hitung cost of sales
                    If .Stock_Out_Date <> #1/1/1950# Then
                        .Cost_Of_Sales = Me.GetCostOfSales(.Primary_DO, TblInv)
                        '.Cost_Of_Sales = Me.GetCostOfSales_Obs(.Stock_Out_Date, .DO_No, TBLDO_Detail)
                    Else
                        'belum ada stock out maka belum ada cost of sales
                        .Cost_Of_Sales = 0
                    End If
                    'hitung GM %
                    If .Sales_Amount <> 0 Then
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) * 100 / .Sales_Amount, 2, MidpointRounding.AwayFromZero)
                    End If
                End With
            Next
            DtaInv.Dispose()
            TblInv.Dispose()

            'generate report
            Dim FilterBy As String = ""
            If Me.rboAll.Checked = True Then
                FilterBy = "All"
            ElseIf Me.rboInvoiceDO.Checked = True Then
                FilterBy = "Invoice DO"
            ElseIf Me.rboNotInvoiceDO.Checked = True Then
                FilterBy = "Not Invoice DO"
            ElseIf Me.rboDONotStockedOut.Checked Then 'Update SMESS 2026
                FilterBy = "DO Not Stocked Out"
            End If

            ReportByDO.SetDataSource(DS)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_FilterBy.ParameterFieldName, FilterBy)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_FormatHarga.ParameterFieldName, Me.DefPrecIDR)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_TglDari.ParameterFieldName, Me.dtpDari.Value)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_TglSampai.ParameterFieldName, Me.dtpSampai.Value)
            ReportByDO.SetParameterValue(ReportByDO.Parameter_Rate.ParameterFieldName, RateString)

            Me.crv1.ReportSource = ReportByDO
            Me.crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTADO_Detail.Dispose()
            TBLDO_Detail.Dispose()
            TBL_ExchangeRate.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Function GetRateString(ByVal vTBLRate As DseLaporanSalesListByDO.DataTable5DataTable) As String
        Dim row As DseLaporanSalesListByDO.DataTable5Row
        Dim Dummy As String = ""
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            For Each row In vTBLRate.Rows
                If row.Primary_Currency <> 1 Then
                    If Dummy = "" Then
                        Dummy = ObjCD.DigitGrouping(row.Rate) + " IDR/" + row.Currency_Code
                    Else
                        Dummy = Dummy + "; " + ObjCD.DigitGrouping(row.Rate) + " IDR/" + row.Currency_Code
                    End If
                End If
            Next

            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function GetSalesAmount(ByVal vDONumber As String, ByVal vTBL_ExchangeRate As DseLaporanSalesListByDO.DataTable5DataTable, _
    ByVal vTBL_DO_Detail As DseLaporanSalesListByDO.DataTable4DataTable) As Double
        Dim Filter As String = "[DO Number] = '" + vDONumber + "'"
        Dim rowView As DataRowView
        Dim rowDetail As DseLaporanSalesListByDO.DataTable4Row
        Dim SalesAmount, ExchangeRate As Double
        Dim QtyHitung As Decimal = 0

        Try
            vTBL_DO_Detail.DefaultView.RowFilter = Filter
            For Each rowView In vTBL_DO_Detail.DefaultView
                rowDetail = rowView.Row
                ExchangeRate = Me.GetExchangeRate(vTBL_ExchangeRate, rowDetail.Primary_Currency)

                If rowDetail.Primary_Units <> rowDetail.Primary_Units_Multi And rowDetail.price_by_multi_unit = 0 Then
                    QtyHitung = rowDetail.Total_Qty
                Else
                    QtyHitung = rowDetail.Total_Qty_Multi 'awalnya selalu menggunakan qty multi
                End If

                If rowDetail.Price_Exclude = True Then
                    If rowDetail.Secondary_Qty = -1 Then
                        SalesAmount += Math.Round(rowDetail.Nett_Price * ExchangeRate * QtyHitung, Me.DefPrecIDR, MidpointRounding.AwayFromZero)
                    Else
                        SalesAmount += Math.Round(rowDetail.Nett_Price * ExchangeRate * QtyHitung * rowDetail.Secondary_Qty, Me.DefPrecIDR, MidpointRounding.AwayFromZero)
                    End If
                Else
                    If rowDetail.Secondary_Qty = -1 Then
                        SalesAmount += Math.Round(Math.Round(rowDetail.Nett_Price * 100 / (100 + rowDetail.Tax), rowDetail.Default_Precision, MidpointRounding.AwayFromZero) _
                        * ExchangeRate * QtyHitung, Me.DefPrecIDR, MidpointRounding.AwayFromZero)
                    Else
                        SalesAmount += Math.Round(Math.Round(rowDetail.Nett_Price * 100 / (100 + rowDetail.Tax), rowDetail.Default_Precision, MidpointRounding.AwayFromZero) _
                        * ExchangeRate * QtyHitung * rowDetail.Secondary_Qty, Me.DefPrecIDR, MidpointRounding.AwayFromZero)
                    End If
                End If
            Next

            Return SalesAmount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetExchangeRate(ByVal vTBLExchangeRate As DseLaporanSalesListByDO.DataTable5DataTable, _
    ByVal vPrimaryCurr As Double) As Double
        Dim row As DseLaporanSalesListByDO.DataTable5Row

        Try
            For Each row In vTBLExchangeRate.Rows
                If row.Primary_Currency = vPrimaryCurr Then
                    Return row.Rate
                End If
            Next

            Return 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Private Sub SetConvertionRate(ByVal vTBL As DseLaporanSalesListByDO.DataTable4DataTable, _
    ByVal vTBLMataUang As DseLaporanSalesListByDO.DataTable5DataTable)
        Dim i As New FrmLaporanSalesListByDO_MataUang
        Dim rowDO As DseLaporanSalesListByDO.DataTable4Row
        Dim ListCurr As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim rowMataUang As DseLaporanSalesListByDO.DataTable5Row

        Try
            'kosongkan bari exchange rate
            vTBLMataUang.Rows.Clear()
            'set konversi ke rupiah menjadi 0
            ListCurr.Add(1)
            rowMataUang = vTBLMataUang.NewDataTable5Row
            rowMataUang.Primary_Currency = 1
            rowMataUang.Rate = 1
            rowMataUang.Currency_Code = "IDR"
            vTBLMataUang.Rows.Add(rowMataUang)
            'masukkan mata uang untuk menghitung rate
            For Each rowDO In vTBL.Rows
                If ListCurr.Contains(rowDO.Primary_Currency) = False Then
                    ListCurr.Add(rowDO.Primary_Currency)
                    i.dgvItem.Rows.Add()
                    rowDGV = i.dgvItem.Rows(i.dgvItem.Rows.Count - 1)
                    rowDGV.Cells(i.ColCurrency.Index).Value = rowDO.Currency_Code
                    rowDGV.Cells(i.ColExchangeRate.Index).Value = 1
                    rowDGV.Cells(i.ColNo.Index).Value = i.dgvItem.Rows.Count
                    rowDGV.Cells(i.ColPrimaryCurrency.Index).Value = rowDO.Primary_Currency
                End If
            Next
            If i.dgvItem.Rows.Count Then
                i.ShowDialog()
            End If
            'masukkan data yang sudah dituliskan user ke dalam tabel untuk perhitungan
            For Each rowDGV In i.dgvItem.Rows
                rowMataUang = vTBLMataUang.NewDataTable5Row
                rowMataUang.Primary_Currency = rowDGV.Cells(i.ColPrimaryCurrency.Index).Value
                rowMataUang.Rate = rowDGV.Cells(i.ColExchangeRate.Index).Value
                rowMataUang.Currency_Code = rowDGV.Cells(i.ColCurrency.Index).Value
                vTBLMataUang.Rows.Add(rowMataUang)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            ListCurr = Nothing
        End Try
    End Sub

    Private Function GetCostOfSales(ByVal vPrimaryDO As Double, _
    ByVal vTblInv As DseLaporanSalesListByDO.tbltrans_salesinvoice_itemDataTable) As Double
        Dim TotalCost As Double = 0
        Dim rowInv As DseLaporanSalesListByDO.tbltrans_salesinvoice_itemRow

        Try
            vTblInv.DefaultView.RowFilter = "[Primary DO] = " & vPrimaryDO
            If vTblInv.DefaultView.Count > 0 Then
                rowInv = vTblInv.DefaultView(0).Row
                TotalCost = rowInv.Cost
            End If

            Return TotalCost
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return TotalCost
        End Try
    End Function

    Private Function GetCostOfSales_Obs(ByVal vDODate As Date, _
    ByVal DoNo As String, ByVal vTBLDODetail As DseLaporanSalesListByDO.DataTable4DataTable) As Double
        Dim cmd As String = ""
        Dim cmd2 As String = ""
        Dim row1 As DseLaporanSalesListByDO.DataTable4Row
        Dim row2 As DseLaporanSalesListByDO.DataTable3Row
        Dim TBLView As DataView = vTBLDODetail.DefaultView
        Dim TBLView2 As DataView = Me.TBL3.DefaultView
        Dim TotalCost As Double = 0

        Try
            'filter data penjualan
            cmd = "[Do Number] ='" + DoNo + "'"

            TBLView.RowFilter = cmd
            'hitung total cost
            For Each RowView In TBLView
                row1 = RowView.Row
                cmd2 = "[Primary Product] =" + CStr(row1.Primary_Product) + " and " + _
                "[Closing Period] =#" + Me.ObjCD.GetFirstDateofTheMonth(vDODate).ToShortDateString + "#"
                TBLView2.RowFilter = cmd2
                If TBLView2.Count <> 0 Then
                    row2 = TBLView2.Item(0).Row
                    If row2.Ending_Amount = 0 Then
                        If row2.Out_Qty <> 0 Then
                            TotalCost += Math.Round(row2.Out_Amount / row2.Out_Qty, _
                            0, MidpointRounding.AwayFromZero) * row1.Total_Qty
                        End If
                    Else
                        If row2.Ending_Qty <> 0 Then
                            TotalCost += Math.Round(row2.Ending_Amount / row2.Ending_Qty, _
                            0, MidpointRounding.AwayFromZero) * row1.Total_Qty
                        End If
                    End If
                End If
            Next

            Return TotalCost
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnShowReport_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim DTA2 As New DseLaporanSalesListByDOTableAdapters.DataTable2TableAdapter
        Dim DTA3 As New DseLaporanSalesListByDOTableAdapters.DataTable3TableAdapter
        Dim DTA4 As New DseLaporanSalesListByDOTableAdapters.tbltrans_doTableAdapter

        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim row2 As DseLaporanSalesListByDO.DataTable2Row
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            Me.DefPrecIDR = Me.ObjCurr.GetDecimal_DefaultCurrency()

            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpDari.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpDari.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpSampai.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpSampai.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA3.Connection.ConnectionString = Core_Module.GetConstring
            DTA4.Connection.ConnectionString = Core_Module.GetConstring

            'set awal dan akhir periode perhitungan
            Me.StartPeriode = Me.dtpDari.Value.Date
            Me.EndPeriod = Me.dtpSampai.Value.Date

            TBL2.Rows.Clear()
            Me.TBL3.Rows.Clear()
            Me.TBL4.Rows.Clear()
            Me.DS.Group_By_DO.Rows.Clear()

            'isikan tabel yang akan digunakan untuk perhitungan
            If Me.cboFilterBy.SelectedIndex = 0 Then
                DTA2.FillBy_TanggalStockInOut(TBL2, Me.StartPeriode, Me.EndPeriod)
            Else
                DTA2.Fill(TBL2, EndPeriod, Me.StartPeriode)
            End If
            'Dim vx As New FormTest
            'vx.DataGridView1.DataSource = Nothing
            'vx.DataGridView1.DataSource = TBL2
            'vx.ShowDialog()



            'TBL2View = Me.TBL2.DefaultView
            'TBL2View.Sort = "[Do Date] ASC"
            'If TBL2View.Count > 0 Then
            '    row2 = TBL2View.Item(0).Row
            '    Me.StartPeriodInvoice = Me.ObjCD.GetFirstDateofTheMonth(row2.Do_Date)
            '    row2 = TBL2View.Item(TBL2View.Count - 1).Row
            '    Me.EndPeriodInvoice = Me.ObjCD.GetLasDateOfTheMonth(row2.Do_Date)
            '    DTA3.FillBy_TanggalSampaiDari(TBL3, Me.EndPeriodInvoice, Me.StartPeriodInvoice)
            '    DTA4.Fill(TBL4, Me.EndPeriod, Me.StartPeriode)
            '    'lakukan perhitungan
            '    Me.GenerateByDO()
            'Else
            '    Me.crv1.ReportSource = Nothing
            'End If



            'If Me.cboFilterBy.SelectedIndex = 0 Then
            '    Me.StartPeriodInvoice = Me.ObjCD.GetFirstDateofTheMonth(Me.GetMinInvoiceDateByStock(Me.StartPeriode, Me.EndPeriod))
            '    Me.EndPeriodInvoice = Me.ObjCD.GetLasDateOfTheMonth(Me.GetMaxInvoiceDateByStock(Me.StartPeriode, Me.EndPeriod))
            'Else
            '    Me.StartPeriodInvoice = Me.ObjCD.GetFirstDateofTheMonth(Me.GetMinInvoiceDateByDO(Me.StartPeriode, Me.EndPeriod))
            '    Me.EndPeriodInvoice = Me.ObjCD.GetLasDateOfTheMonth(Me.GetMaxInvoiceDateByDO(Me.StartPeriode, Me.EndPeriod))
            'End If

            Me.StartPeriodInvoice = Me.ObjCD.GetFirstDateofTheMonth(Me.GetMinInvoiceDateByDO(Me.StartPeriode, Me.EndPeriod))
            Me.EndPeriodInvoice = Me.ObjCD.GetLasDateOfTheMonth(Me.GetMaxInvoiceDateByDO(Me.StartPeriode, Me.EndPeriod))

            DTA3.FillBy_TanggalSampaiDari(TBL3, Me.EndPeriodInvoice, Me.StartPeriodInvoice)

            'If Me.cboFilterBy.SelectedIndex = 0 Then
            '    'DTA4.FillBy_TanggalSampaiDariStockOUT(TBL4, Me.EndPeriod, Me.StartPeriode)
            'Else
            DTA4.Fill(TBL4, Me.EndPeriod, Me.StartPeriode)
            'End If
            'lakukan perhitungan
            Me.GenerateByDO()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA2.Dispose()
            DTA3.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Public Function GetMaxInvoiceDateByDO(ByVal vStartDateDO As Date, ByVal vEndDateDO As Date) As Date
        Dim DTA As New DseLaporanSalesListByDOTableAdapters.tbltrans_do1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).HasValue > 0 AndAlso _
               IsDate(DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).Value) Then
                Return DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).Value
            Else
                Return vEndDateDO
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function GetMinInvoiceDateByDO(ByVal vStartDateDO As Date, ByVal vEndDateDO As Date) As Date
        Dim DTA As New DseLaporanSalesListByDOTableAdapters.tbltrans_do1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).HasValue > 0 AndAlso _
               IsDate(DTA.ScalarQuery_GetMinTanggalInvoice(vStartDateDO, vEndDateDO).Value) Then
                Return DTA.ScalarQuery_GetMinTanggalInvoice(vStartDateDO, vEndDateDO).Value
            Else
                Return vStartDateDO
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function GetMaxInvoiceDateByStock(ByVal vStartDateDO As Date, ByVal vEndDateDO As Date) As Date
        Dim DTA As New DseLaporanSalesListByDOTableAdapters.MaxMinTanggalStokTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).HasValue > 0 AndAlso _
               IsDate(DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).Value) Then
                Return DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).Value
            Else
                Return vEndDateDO
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function GetMinInvoiceDateByStock(ByVal vStartDateDO As Date, ByVal vEndDateDO As Date) As Date
        Dim DTA As New DseLaporanSalesListByDOTableAdapters.MaxMinTanggalStokTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetMaxTanggalInvoice(vStartDateDO, vEndDateDO).HasValue > 0 AndAlso _
               IsDate(DTA.ScalarQuery_GetMinTanggalInvoice(vStartDateDO, vEndDateDO).Value) Then
                Return DTA.ScalarQuery_GetMinTanggalInvoice(vStartDateDO, vEndDateDO).Value
            Else
                Return vStartDateDO
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmLaporanSalesListByDO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Update SMESS 2026
    Private Sub rboDONotStockedOut_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboDONotStockedOut.CheckedChanged
        Try
            If rboDONotStockedOut.Checked Then
                Me.chkNonPendingDO.Checked = False
                Me.chkNonPendingDO.Enabled = False
            Else
                Me.chkNonPendingDO.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class