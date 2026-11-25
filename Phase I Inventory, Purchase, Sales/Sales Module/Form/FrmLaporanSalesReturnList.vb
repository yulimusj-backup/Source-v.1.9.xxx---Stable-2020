Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesReturnList
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDec As Integer = 0

    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim TanggalClosing As Date = #1/1/1950#
    Dim objCD As New Core_Module.ClsCekData

    Dim DTAClosing As New DseLaporanSalesReturnListTableAdapters.DataTable2TableAdapter
    Dim TBLClosing As New DseLaporanSalesReturnList.DataTable2DataTable

    Dim DTA_InAdjustment As New DseLaporanSalesReturnListTableAdapters.DataTable6TableAdapter
    Dim TBL_InAdjustment As New DseLaporanSalesReturnList.DataTable6DataTable
    Dim TanggalClosing2 As Date = #1/1/1950#

    'Dim reportSR As New CrpLaporanSalesReturnStatistic
    Dim DS As New DseLaporanSalesStatistic

    Dim ObjCur As New Core_Module.clsMasterCurrency

    Dim DTA_Prod As New DseLaporanSalesReturnListTableAdapters.tblmaster_productTableAdapter
    Dim TBL_Prod As New DseLaporanSalesReturnList.tblmaster_productDataTable
    Dim row_Prod As DseLaporanSalesReturnList.tblmaster_productRow
    Dim ObjClosing As New ClsTransaksiMonthlyClosing
    Dim ObjLicense As New Core_Module.ClsLicense

    Public Enum EnumTipeLaporan As Integer
        InExplain
        SalesReturnList
    End Enum
    Public TipeLaporan As EnumTipeLaporan = EnumTipeLaporan.InExplain

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Dim PrimaryClosing As Double = -1
        Dim ObjClosing As New ClsTransaksiPurchaseClosing

        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            PrimaryClosing = Me.GetPrimaryPurchaseClosing(Me.dtpFrom.Value)
            If PrimaryClosing = -1 Then
                MsgBox("Please do inventory first for: " + Format(Me.dtpFrom.Value.Date, "MMM-yyyy") + ".", MsgBoxStyle.Information, "Information")
                Me.CrystalReportViewer1.ReportSource = Nothing
                Me.CrystalReportViewer1.Show()
                Exit Sub
            End If

            Me.GenerateReturn_InExplain2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Public Function GetPrimaryPurchaseClosing(ByVal vClosingPeriod As Date) As Double
        Dim DTA As New DseTransaksiPurchaseClosingTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Double)

        Try
            Dummy = DTA.ScalarQuery_GetPrimaryClosing(vClosingPeriod)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub GenerateReturn_InExplain()
        If Me.ApakahBolehCetak = False Then
            Exit Sub
        End If
        Dim report As New CrpLaporanSalesReturnList_Rev_Invoice
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DTASR As New DseLaporanSalesReturnListTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter
        Dim TBLSR As New DseLaporanSalesReturnList.tbltrans_purchaseclosing_item9_1DataTable
        Dim rowSR As DseLaporanSalesReturnList.tbltrans_purchaseclosing_item9_1Row
        Dim DTABrand1 As New DseLaporanSalesReturnListTableAdapters.DataTable5TableAdapter
        Dim TBLBRand1 As New DseLaporanSalesReturnList.DataTable5DataTable
        Dim rowBrand1 As DseLaporanSalesReturnList.DataTable5Row
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        Dim DV As Data.DataRowView
        Dim PrimaryBrand As Double = -1
        Dim counter As Double = 0
        Dim Filter As String = ""
        Dim row As DseLaporanSalesReturnList.SR_RevInvoiceRow
        Dim DS As New DseLaporanSalesReturnList
        Dim Brand As String = ""
        Dim Filter2 As String = ""
        Dim TipeProduct As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            Me.DTA_Prod.Fill(Me.TBL_Prod)

            If Me.chkGoods.Checked = True Then
                TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
            Else
                TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
            End If
            If Me.chkGoods.Checked = True And Me.chkService.Checked Then
                DTASR.Fill(TBLSR, TglSampai, TglDari)
            Else
                DTASR.FillBy1(TBLSR, TglSampai, TglDari, TipeProduct)
            End If

            DTABrand1.ClearBeforeFill = False
            For Each DV In Me.clbBrand.CheckedItems
                PrimaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                DTABrand1.Fill(TBLBRand1, PrimaryBrand)
            Next
            'isikan data brand
            While counter < TBLSR.Rows.Count
                rowSR = TBLSR.Rows(counter)
                Filter = "[Primary Product] =" + rowSR.Primary_Product.ToString
                TBLBRand1.DefaultView.RowFilter = Filter
                If TBLBRand1.DefaultView.Count > 0 Then
                    rowBrand1 = TBLBRand1.DefaultView(0).Row
                    rowSR.Primary_Brand = rowBrand1.Primary_Brand
                    rowSR.Primary_Product = rowBrand1.Primary_Product
                    rowSR.Product_Code = rowBrand1.Product_Code
                    rowSR.Description = rowBrand1.Description
                    rowSR.Brand = rowBrand1.Brand
                    rowSR.Unit_Code = rowBrand1.Unit_Code
                    counter += 1
                Else
                    TBLSR.Rows.Remove(rowSR)
                End If
            End While

            For Each rowSR In TBLSR.Rows
                With rowSR
                    If TglSampai >= .DO_Date.Date And TglDari <= .DO_Date.Date Then
                        'hitung ulang unit price in sales return untuk periode yang sama
                        .Unit_Price = Me.GetUnitPriceStockInSalesReturnPeriodeSama(.Primary_Product, _
                        .Primary_Stock_In_Sales_Return, TglDari)
                        .Total_Unit_Price = .Total_Qty * .Unit_Price
                    ElseIf .DO_Date < TglDari And .Unit_Price = 0 Then
                        'karena data pada closing end qty =0 maka unit pricenya seharusnya dihitung dari out qty
                        .Unit_Price = Me.GetCostOfSales(.Primary_Product, _
                        objCD.GetFirstDateofTheMonth(.DO_Date))
                        .Total_Unit_Price = .Total_Qty * .Unit_Price
                    End If


                    row = DS.SR_RevInvoice.NewSR_RevInvoiceRow
                    row.Brand_Code = .Brand
                    row.Cost = .Unit_Price * .Total_Qty
                    row.Description = .Description
                    row.Invoice_Date = Nothing
                    row.Invoice_Number = "-"
                    row.Primary_Brand = .Primary_Brand
                    row.Primary_Product = .Primary_Product
                    row.Product_Code = .Product_Code
                    row.SR_Number = .SR_Number
                    row.Unit_Code = .Unit_Code
                    row.Quantity = .Total_Qty

                    Filter2 = "Primary = " + .Primary_Product.ToString
                    Me.TBL_Prod.DefaultView.RowFilter = Filter2
                    Me.row_Prod = Me.TBL_Prod.DefaultView(0).Row
                    row._Decimal = Me.row_Prod._Decimal

                    DS.SR_RevInvoice.Rows.Add(row)
                End With
            Next

            If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
                Brand = "All"
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    If Brand = "" Then
                        Brand = DV.Item(1)
                    Else
                        Brand += ", " + DV.Item(1)
                    End If
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FromatDecimal.ParameterFieldName, Me.FormatDec)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "Sales Return List")
            report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR.Dispose()
            TBLSR.Dispose()
            DTABrand1.Dispose()
            TBLBRand1.Dispose()
            DS.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub GenerateReturn_InExplain2()
        If Me.ApakahBolehCetak = False Then
            Exit Sub
        End If
        Dim report As New CrpLaporanSalesReturnList_Rev_Invoice
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DTASR As New DseLaporanSalesReturnListTableAdapters.tbltrans_purchase_closing_item_detail1TableAdapter
        Dim TBLSR As New DseLaporanSalesReturnList.tbltrans_purchase_closing_item_detail1DataTable
        Dim rowSR As DseLaporanSalesReturnList.tbltrans_purchase_closing_item_detail1Row
        Dim DTABrand1 As New DseLaporanSalesReturnListTableAdapters.DataTable5TableAdapter
        Dim TBLBRand1 As New DseLaporanSalesReturnList.DataTable5DataTable

        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        Dim DV As Data.DataRowView
        Dim PrimaryBrand As Double = -1
        Dim counter As Double = 0
        Dim Filter As String = ""
        Dim row As DseLaporanSalesReturnList.SR_RevInvoiceRow
        Dim DS As New DseLaporanSalesReturnList
        Dim Brand As String = ""
        Dim Filter2 As String = ""
        Dim DTA_Inv As New DseLaporanSalesReturnListTableAdapters.DataTable4TableAdapter
        Dim TBL_Inv As New DseLaporanSalesReturnList.DataTable4DataTable
        Dim row_Inv As DseLaporanSalesReturnList.DataTable4Row
        Dim DTAQuickSR As New DseLaporanSalesReturnListTableAdapters.DataTable7TableAdapter
        Dim TBLQuickSR As New DseLaporanSalesReturnList.DataTable7DataTable
        Dim rowQuickSR As DseLaporanSalesReturnList.DataTable7Row
        Dim TipeItemDetail As ClsTransaksiMonthlyClosing.EnumTipeItemDetail = ClsTransaksiMonthlyClosing.EnumTipeItemDetail.InItem

        Try
            TBLSR.Rows.Clear()
            TBL_Inv.Rows.Clear()
            TBLQuickSR.Rows.Clear()

            DTAQuickSR.Connection.ConnectionString = Core_Module.GetConstring
            DTAQuickSR.Fill(TBLQuickSR, Me.objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value), Me.objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value))

            'Me.DTA_Prod.Fill(Me.TBL_Prod)
            DTASR.Fill(TBLSR, ClsTransaksiMonthlyClosing.EnumTipeTranskasiAsal.InSR, Me.ObjClosing.GetPrimaryClosing(Me.objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)), CInt(TipeItemDetail))
            DTA_Inv.Fill(TBL_Inv, ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak, TglSampai, TglDari)

            'isikan data brand
            DTABrand1.ClearBeforeFill = False
            For Each DV In Me.clbBrand.CheckedItems
                PrimaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                DTABrand1.Fill(TBLBRand1, PrimaryBrand)
            Next

            'buang data yang tidak dibutuhkan
            While counter < TBLSR.Rows.Count
                rowSR = TBLSR.Rows(counter)
                Filter = "[Primary Product] =" + rowSR.Primary_Product.ToString
                TBLBRand1.DefaultView.RowFilter = Filter
                If TBLBRand1.DefaultView.Count > 0 Then
                    'isikan data berkaitan dengan invoice
                    TBL_Inv.DefaultView.RowFilter = "[Primary Stock In Sales Return] = " + rowSR.Primary_Stock_In_Sales_Return.ToString

                    If TBL_Inv.DefaultView.Count = 0 Then
                        TBLQuickSR.DefaultView.RowFilter = "[Primary Stock In SR] = " + rowSR.Primary_Stock_In_Sales_Return.ToString
                        If TBLQuickSR.DefaultView.Count > 0 Then
                            rowQuickSR = TBLQuickSR.DefaultView(0).Row
                            rowSR.Invoice_Date = rowQuickSR.Invoice_Date_Quick
                            rowSR.Invoice_Number = rowQuickSR.Invoice_No_Quick
                            rowSR.SR_Number = rowQuickSR.SR_Number
                        Else
                            MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    Else
                        row_Inv = TBL_Inv.DefaultView(0).Row
                        rowSR.Invoice_Date = row_Inv.Invoice_Date
                        rowSR.Invoice_Number = row_Inv.No_Transaksi
                        rowSR.SR_Number = row_Inv.SR_Number
                    End If

                    counter += 1
                Else
                    TBLSR.Rows.Remove(rowSR)
                End If
            End While

            For Each rowSR In TBLSR.Rows
                With rowSR
                    row = DS.SR_RevInvoice.NewSR_RevInvoiceRow
                    row.Brand_Code = .Brand
                    row.Cost = .Cost
                    row.Description = .Description
                    row.Invoice_Date = .Invoice_Date
                    row.Invoice_Number = .Invoice_Number
                    row.Primary_Brand = .Primary_Brand
                    row.Primary_Product = .Primary_Product
                    row.Product_Code = .Product_Code
                    row.SR_Number = .SR_Number
                    row.Unit_Code = .Unit_Code
                    row.Quantity = .Quantity
                    row._Decimal = ._Decimal

                    DS.SR_RevInvoice.Rows.Add(row)
                End With
            Next

            If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
                Brand = "All"
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    If Brand = "" Then
                        Brand = DV.Item(1)
                    Else
                        Brand += ", " + DV.Item(1)
                    End If
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FromatDecimal.ParameterFieldName, Me.FormatDec)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "Sales Return Cost Report")
            report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR.Dispose()
            TBLSR.Dispose()
            DTABrand1.Dispose()
            TBLBRand1.Dispose()
            DTA_Inv.Dispose()
            TBL_Inv.Dispose()
            DS.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTAQuickSR.Dispose()
            TBLQuickSR.Dispose()
        End Try
    End Sub

    Private Sub FrmLaporanSalesReturnList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup

        Try
            Me.FormatDec = Me.ObjCur.GetDecimal_DefaultCurrency()
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.cboTipe.SelectedIndex = 0
            Me.dtpFrom.Value = Me.objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Me.Text = "Sales Return Cost Report"
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBjBrand = Nothing
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahBolehCetak() As Boolean
        Try
            If Me.clbBrand.CheckedItems.Count <= 0 Then
                MsgBox("Please select at least 1 Brand.", MsgBoxStyle.Information, "Information")
                Me.clbBrand.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function GetCostOfSales(ByVal vPrimaryProduct As Double, ByVal vClosingPeriod As Date) As Double
        Dim filter As String = ""
        Dim row As DseLaporanSalesReturnList.DataTable2Row

        Try
            'sesuaikan tanggal closing, isikan data closing sesuai dengan tanggal closing
            If vClosingPeriod <> Me.TanggalClosing Then
                Me.DTAClosing.Fill(Me.TBLClosing, vClosingPeriod)
                Me.TanggalClosing = vClosingPeriod
            End If

            filter = "[Primary Product] =" + vPrimaryProduct.ToString
            Me.TBLClosing.DefaultView.RowFilter = filter
            If Me.TBLClosing.DefaultView.Count <> 0 Then
                row = Me.TBLClosing.DefaultView(0).Row
                If row.Ending_Qty = 0 Or row.Ending_Amount = 0 Then
                    If row.Out_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(row.Out_Amount / row.Out_Qty, 0, MidpointRounding.AwayFromZero)
                    End If
                Else
                    If row.Ending_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(row.Ending_Amount / row.Ending_Qty, 0, MidpointRounding.AwayFromZero)
                    End If
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Function GetUnitPriceStockInSalesReturnPeriodeSama(ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryStockInSalesReturn As Double, ByVal vClosingPeriod As Date) As Double
        Dim rowAdjustment As DseLaporanSalesReturnList.DataTable6Row
        Dim Filter As String = ""

        Try
            If vClosingPeriod <> Me.TanggalClosing2 Then
                Me.DTA_InAdjustment.Fill(Me.TBL_InAdjustment, vClosingPeriod)
                Me.TanggalClosing2 = vClosingPeriod
            End If

            Filter = "[Primary Code No] =" + vPrimaryProduct.ToString + _
            " AND [Primary Dokumen Acuan] =" + vPrimaryStockInSalesReturn.ToString

            Me.TBL_InAdjustment.DefaultView.RowFilter = Filter
            If Me.TBL_InAdjustment.DefaultView.Count <> 0 Then
                rowAdjustment = Me.TBL_InAdjustment.DefaultView(0).Row

                Return rowAdjustment.Unit_Price
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub MasukkanDataSalesReturn(ByVal rowSumber As DseLaporanSalesReturnStatistic.DataTable1Row, _
    ByVal rowTarget As DseLaporanSalesStatistic.Group_By_Cust_Code1Row)
        Dim Bulan As Integer = 0

        Try
            Bulan = rowSumber.Tanggal_Transaksi.Month
            With rowTarget
                If Bulan = 1 Then
                    .January = .January + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 2 Then
                    .February = .February + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 3 Then
                    .Maret = .Maret + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 4 Then
                    .April = .April + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 5 Then
                    .Mei = .Mei + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 6 Then
                    .Juni = .Juni + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 7 Then
                    .Juli = .Juli + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 8 Then
                    .Agustus = .Agustus + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 9 Then
                    .September = .September + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 10 Then
                    .Oktober = .Oktober + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 11 Then
                    .Nopember = .Nopember + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                ElseIf Bulan = 12 Then
                    .Desember = .Desember + rowSumber.Total_Return
                    .Total_Sales = .Total_Sales + rowSumber.Total_Return
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class