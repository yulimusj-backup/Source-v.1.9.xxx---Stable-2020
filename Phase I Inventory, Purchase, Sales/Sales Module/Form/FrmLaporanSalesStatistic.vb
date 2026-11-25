Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesStatistic
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDecimal As Integer = 2
    Dim DS As New DseLaporanSalesStatistic
    Dim ReportByCustCode As New CrpLaporanSalesStatisticByCustomer
    Dim ReportByCustGroup As New CrpLaporanSalesStatisticByCustomerGroup
    Dim ObjCD As New Core_Module.ClsCekData

    Dim RowView As DataRowView
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim rowBrand As Core_Module.dseMasterProductGroup.tblmaster_product_group2Row
    Dim TBlCustType As New Core_Module.dseMasterCustomerType.tblmaster_customer_type2DataTable
    Dim rowCustType As Core_Module.dseMasterCustomerType.tblmaster_customer_type2Row
    Dim TblCust As New Core_Module.dseCustomer.tblmaster_customer1DataTable
    Dim rowCust As Core_Module.dseCustomer.tblmaster_customer1Row
    Dim TblSegmentation As New Core_Module.dseMasterCustomerMarketSegmentation.tblmaster_customer_segmentation2DataTable
    Dim rowSegmentation As Core_Module.dseMasterCustomerMarketSegmentation.tblmaster_customer_segmentation2Row
    Dim TblSalesArea As New Core_Module.dseMasterCustomerSalesArea.tblmaster_customer_sales_area2DataTable
    Dim rowSalesArea As Core_Module.dseMasterCustomerSalesArea.tblmaster_customer_sales_area2Row

    Dim ListCustType As New List(Of String)
    Dim ListSalesArea As New List(Of String)
    Dim ListBrand As New List(Of String)
    Dim ListCust As New List(Of String)
    Dim ListSegmentation As New List(Of String)

    Dim TBL As DseLaporanSalesStatistic.DataTable1DataTable = DS.DataTable1
    Dim TBL2 As DseLaporanSalesStatistic.DataTable2DataTable = DS.DataTable2
    Dim TBLGroupByCustCode As DseLaporanSalesStatistic.Group_By_Cust_CodeDataTable = DS.Group_By_Cust_Code
    Dim TBLGroupByCustGroupCode As DseLaporanSalesStatistic.Group_By_Cust_Group_CodeDataTable = DS.Group_By_Cust_Group_Code
    Dim StartPeriode, StartPeriodYTD, EndPeriod As Date
    Dim TBLReturn As New DseLaporanSalesStatistic.TblReturn_InvoiceDataTable
    Dim DTAReturn As New DseLaporanSalesStatisticTableAdapters.TblReturn_InvoiceTableAdapter
    Dim rowReturn As DseLaporanSalesStatistic.TblReturn_InvoiceRow

    Private Function ApakahBolehCetak() As Boolean
        Try
            'If Me.clbBrand.CheckedItems.Count <= 0 Then
            '    MsgBox("Please select at least 1 product brand.", MsgBoxStyle.Information, "Information")
            '    Me.clbBrand.Focus()
            '    Return False
            'End If
            'If Me.clbCustomer.CheckedItems.Count <= 0 Then
            '    MsgBox("Please select at least 1 customer.", MsgBoxStyle.Information, "Information")
            '    Me.clbCustomer.Focus()
            '    Return False
            'End If
            'If Me.clbCustomerType.CheckedItems.Count <= 0 Then
            '    MsgBox("Please select at least 1 customer type.", MsgBoxStyle.Information, "Information")
            '    Me.clbCustomerType.Focus()
            '    Return False
            'End If
            'If Me.clbMarketSegmentation.CheckedItems.Count <= 0 Then
            '    MsgBox("Please select at least 1 market segmentation.", MsgBoxStyle.Information, "Information")
            '    Me.clbMarketSegmentation.Focus()
            '    Return False
            'End If
            'If Me.clbSalesArea.CheckedItems.Count <= 0 Then
            '    MsgBox("Please select at least 1 sales area.", MsgBoxStyle.Information, "Information")
            '    Me.clbSalesArea.Focus()
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub KurangkanItemReturn()
        Dim Filter As String = ""
        Dim row1 As DseLaporanSalesStatistic.DataTable1Row
        Dim i As Integer = 0

        Try
            For Each rowReturn In Me.TBLReturn.Rows
                Filter = "[Primary Invoice] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                Me.TBL.DefaultView.RowFilter = Filter

                i = 0
                While Me.TBL.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row1 = Me.TBL.DefaultView(i).Row
                    If row1.Total_Qty >= rowReturn.Total_Qty Then
                        row1.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    End If

                    If row1.Total_Qty = 0 Then
                        Me.TBL.Rows.Remove(row1)
                    End If
                End While
            Next

            Me.TBL.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim DTA As New DseLaporanSalesStatisticTableAdapters.DataTable1TableAdapter
        Dim DTA2 As New DseLaporanSalesStatisticTableAdapters.DataTable2TableAdapter
        Try
            If Me.ApakahBolehCetak = True Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                'kosongkan data
                Me.DS.DataTable1.Rows.Clear()
                Me.DS.DataTable2.Rows.Clear()
                Me.DS.DataTable3.Rows.Clear()
                Me.DS.Group_By_Cust_Code.Rows.Clear()
                Me.DS.Group_By_Cust_Code1.Rows.Clear()
                Me.DS.Group_By_Cust_Group_Code.Rows.Clear()
                Me.DS.Group_By_Cust_Group_Code1.Rows.Clear()

                'set awal dan akhir periode perhitungan
                Me.StartPeriode = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSalesInvoicePeriod.Value)
                Me.EndPeriod = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSalesInvoicePeriod.Value)
                Me.StartPeriodYTD = CDate("1/1/" + CStr(Me.dtpSalesInvoicePeriod.Value.Year))

                'isikan tabel yang akan digunakan untuk perhitungan
                DTA.FillBy_TanggalSampaiDari(TBL, Me.EndPeriod, Me.StartPeriodYTD)
                DTA2.Fill(TBL2, Me.EndPeriod, Me.StartPeriodYTD)
                'kurangkan dengan item return pada invoice
                DTAReturn.Fill(Me.TBLReturn, Me.EndPeriod, Me.StartPeriodYTD)
                Me.KurangkanItemReturn()

                'set kriteria filter
                Me.ListBrand.Clear()
                Me.ListCust.Clear()
                Me.ListCustType.Clear()
                Me.ListSalesArea.Clear()
                Me.ListSegmentation.Clear()
                'If Me.chkBrand.Checked = False Then
                '    For Each RowView In Me.clbBrand.CheckedItems
                '        rowBrand = RowView.Row
                '        Me.ListBrand.Add(rowBrand.Brand)
                '    Next
                'End If
                'If Me.chkCustomer.Checked = False Then
                '    For Each RowView In Me.clbCustomer.CheckedItems
                '        rowCust = RowView.Row
                '        Me.ListCust.Add(rowCust.Customer_Code)
                '    Next
                'End If
                'If Me.chkCustType.Checked = False Then
                '    For Each RowView In Me.clbCustomerType.CheckedItems
                '        rowCustType = RowView.Row
                '        Me.ListCustType.Add(rowCustType.Type_Code)
                '    Next
                'End If
                'If Me.chkMarketSegmentation.Checked = False Then
                '    For Each RowView In Me.clbMarketSegmentation.CheckedItems
                '        rowSegmentation = RowView.Row
                '        Me.ListSegmentation.Add(rowSegmentation.Segmentation_Code)
                '    Next
                'End If
                'If Me.chkSalesArea.Checked = False Then
                '    For Each RowView In Me.clbSalesArea.CheckedItems
                '        rowSalesArea = RowView.Row
                '        Me.ListSalesArea.Add(rowSalesArea.Sales_Area_Code)
                '    Next
                'End If
                'set nilai maksimum progress bar
                Me.pbr1.Value = 0
                Me.pbr1.Maximum = TBL.Rows.Count
                'lakukan perhitungan
                If Me.rboByCustCode.Checked = True Then
                    Me.GenerateByCustCode()
                ElseIf Me.rboByCustGroupCode.Checked = True Then
                    Me.GenerateByCustGroup()
                End If
                'pidahkan tab order
                Me.TabControl1.SelectedIndex = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Function FilterByDgvItem() As String
        Dim FilterBy_CustomerType As String = ""
        Dim FilterBy_Customer As String = ""
        Dim FilterBy_Market As String = ""
        Dim FilterBy_SalesArea As String = ""
        Dim FilterBy_Brand As String = ""
        Dim Filter As String = ""
        Try
            If Me.DgvCustomerType.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvCustomerType.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColCT_CustType.Index).Value) = False Then
                        If FilterBy_CustomerType = "" Then
                            FilterBy_CustomerType = "[Type Code]='" & row.Cells(Me.ColCT_CustType.Index).Value & "'"
                        Else
                            FilterBy_CustomerType &= " OR [Type Code]='" & row.Cells(Me.ColCT_CustType.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_CustomerType <> "" Then
                    Filter = "(" & FilterBy_CustomerType & ")"
                End If
            End If

            If Me.DgvCustomer.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvCustomer.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColC_Code.Index).Value) = False Then
                        If FilterBy_Customer = "" Then
                            FilterBy_Customer = "[Customer Code]='" & row.Cells(Me.ColC_Code.Index).Value & "'"
                        Else
                            FilterBy_Customer &= " OR [Customer Code]='" & row.Cells(Me.ColC_Code.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_Customer <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_Customer & ")"
                    Else
                        Filter &= "(" & FilterBy_Customer & ")"
                    End If
                End If
            End If

            If Me.DgvMarketSegmentation.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvMarketSegmentation.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColMS_MarketSeg.Index).Value) = False Then
                        If FilterBy_Market = "" Then
                          FilterBy_Market = "[Segmentation Code]='" & row.Cells(Me.ColMS_MarketSeg.Index).Value & "'"
                        Else
                            FilterBy_Market &= " OR [Segmentation Code]='" & row.Cells(Me.ColMS_MarketSeg.Index).Value & "'"
                        End If
                    End If
                Next

                If FilterBy_Market <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_Market & ")"
                    Else
                        Filter = "(" & FilterBy_Market & ")"
                    End If
                End If
            End If

            If Me.DgvSalesArea.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvSalesArea.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColSA_SalesArea.Index).Value) = False Then
                        If FilterBy_SalesArea = "" Then
                            FilterBy_SalesArea = "[Sales Area Code]='" & row.Cells(Me.ColSA_SalesArea.Index).Value & "'"
                        Else
                            FilterBy_SalesArea &= " OR [Sales Area Code]='" & row.Cells(Me.ColSA_SalesArea.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_SalesArea <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_SalesArea & ")"
                    Else
                        Filter = "(" & FilterBy_SalesArea & ")"
                    End If
                End If
            End If

            If Me.DgvBrand.RowCount > 0 Then
                For Each row As DataGridViewRow In Me.DgvBrand.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColB_Brand.Index).Value) = False Then
                        If FilterBy_Brand = "" Then
                            FilterBy_Brand = "[Brand]='" & row.Cells(Me.ColB_Brand.Index).Value & "'"
                        Else
                            FilterBy_Brand &= " OR [Brand]='" & row.Cells(Me.ColB_Brand.Index).Value & "'"
                        End If
                    End If
                Next
                If FilterBy_Brand <> "" Then
                    If Filter <> "" Then
                        Filter &= "AND (" & FilterBy_Brand & ")"
                    Else
                        Filter &= "(" & FilterBy_Brand & ")"
                    End If
                End If
            End If

            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function ApakahMasukKriteria(ByVal row1 As DseLaporanSalesStatistic.DataTable1Row) As Boolean
        Dim ketemu As Boolean = False
        Try
            'cek cust type
            'If Me.chkCustType.Checked = False Then
            '    ketemu = False
            '    For Each RowView In Me.clbCustomerType.CheckedItems
            '        rowCustType = RowView.Row
            '        If rowCustType.Type_Code = row1.Type_Code Then
            '            ketemu = True
            '        End If
            '    Next
            '    If ketemu = False Then
            '        Return False
            '    End If
            'End If
            ''cek brand
            'If Me.chkBrand.Checked = False Then
            '    ketemu = False
            '    For Each RowView In Me.clbBrand.CheckedItems
            '        rowBrand = RowView.Row
            '        If rowBrand.Brand = row1.Brand Then
            '            ketemu = True
            '        End If
            '    Next
            '    If ketemu = False Then
            '        Return False
            '    End If
            'End If
            ''cek cust code
            'If Me.chkCustomer.Checked = False Then
            '    ketemu = False
            '    For Each RowView In Me.clbCustomer.CheckedItems
            '        rowCust = RowView.Row
            '        If rowCust.Customer_Code = row1.Customer_Code Then
            '            ketemu = True
            '        End If
            '    Next
            '    If ketemu = False Then
            '        Return False
            '    End If
            'End If
            ''cek sales area
            'If Me.chkSalesArea.Checked = False Then
            '    ketemu = False
            '    For Each RowView In Me.clbSalesArea.CheckedItems
            '        rowSalesArea = RowView.Row
            '        If rowSalesArea.Sales_Area_Code = row1.Sales_Area_Code Then
            '            ketemu = True
            '        End If
            '    Next
            '    If ketemu = False Then
            '        Return False
            '    End If
            'End If
            ''cek segmentation
            'If Me.chkMarketSegmentation.Checked = False Then
            '    ketemu = False
            '    For Each RowView In Me.clbMarketSegmentation.CheckedItems
            '        rowSegmentation = RowView.Row
            '        If rowSegmentation.Segmentation_Code = row1.Segmentation_Code Then
            '            ketemu = True
            '        End If
            '    Next
            '    If ketemu = False Then
            '        Return False
            '    End If
            'End If
            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GenerateByCustCode()
        Dim row1 As DseLaporanSalesStatistic.DataTable1Row
        Dim row2 As DseLaporanSalesStatistic.Group_By_Cust_CodeRow
        Dim SalesAmount As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByCustCodeView As DataView = Me.TBLGroupByCustCode.DefaultView

        Try
            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To TBL.DefaultView.Count - 1
                row1 = TBL.DefaultView(i).Row
                'If Me.ApakahMasukKriteria(row1) = True Then
                'hitung exchage rate
                'CalculatesBy = row1.Calculates_By
                'If CalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
                'cmd = "[Primary Sales Invoice] = " + CStr(row1.Primary_Invoice) + _
                '" and [Primary Transaksi Rujukan] =" + CStr(row1.Primary_DO)
                'ElseIf CalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
                'cmd = "[Primary Sales Invoice] = " + CStr(row1.Primary_Invoice) + _
                '" and [Primary Transaksi Rujukan] =" + CStr(row1.Primary_Invoice)
                'ElseIf CalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
                'cmd = "[Primary Sales Invoice] = " + CStr(row1.Primary_Invoice) + _
                '" and [Primary Transaksi Rujukan] =" + CStr(row1.Primary_SO)
                'End If
                'TBL2View.RowFilter = cmd
                'If TBL2View.Count > 0 Then
                'rowRate = TBL2View.Item(0).Row
                'Rate = rowRate.Exhange_Rate
                'Else
                'Rate = 1
                'End If
                'hitung sales amount berdasarkan rate tsb
                'SalesAmount = Math.Round(Math.Round(row1.Amount * Rate, row1.Default_Precision, MidpointRounding.AwayFromZero) * row1.Tax_Exchange_Rate, _
                '0, MidpointRounding.AwayFromZero)
                SalesAmount = Math.Round(Math.Round(row1.Sales_Amount), _
                0, MidpointRounding.AwayFromZero)

                'masukkan hasilnya ke total sales ytd + baris
                cmd = "[Cust Code] ='" + row1.Customer_Code + "'"
                TBLGroupByCustCodeView.RowFilter = cmd
                If TBLGroupByCustCodeView.Count = 0 Then
                    'tambahkan baris baru
                    TBLGroupByCustCode.Rows.Add(row1.Customer_Code, row1.Customer_Name, 0, 0, 0, _
                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, row1.Sales_Area_Code, row1.Segmentation_Code)
                    'masukkan data sales bulanan dan total sales
                    row2 = TBLGroupByCustCode.Item(TBLGroupByCustCode.Rows.Count - 1)
                    Me.InsertSales(row2, row1.Invoice_Date, SalesAmount)
                Else
                    row2 = TBLGroupByCustCodeView.Item(0).Row
                    'tambahan sales amount ytd
                    Me.InsertSales(row2, row1.Invoice_Date, SalesAmount)
                End If
                'End If
            Next

            'sorting
            If Me.rboSortByCode.Checked = True Then
                If Me.rboAscending.Checked = True Then
                    TBLGroupByCustCodeView.Sort = "[Cust Code] ASC"
                Else
                    TBLGroupByCustCodeView.Sort = "[Cust Code] DESC"
                End If
            Else
                If Me.rboAscending.Checked = True Then
                    TBLGroupByCustCodeView.Sort = "[Total Sales] ASC"
                Else
                    TBLGroupByCustCodeView.Sort = "[Total Sales] DESC"
                End If
            End If

            TBLGroupByCustCodeView.RowFilter = ""
            For Each RowView In TBLGroupByCustCodeView
                'pindahin ke tabel1
                row2 = RowView.Row
                With row2
                    Me.DS.Group_By_Cust_Code1.Rows.Add(.Cust_Code, .Cust_Name, .January, .February, _
                    .Maret, .April, .Mei, .Juni, .Juli, .Agustus, .September, .Oktober, .Nopember, .Desember, _
                    .Total_Sales, .Area, .Segmentation)
                End With
            Next
            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByCustCode.SetDataSource(DS)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByCustCode.SetParameterValue(ReportByCustCode.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            Me.crv1.ReportSource = ReportByCustCode
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanCustomerCodeCustomerTypeSegmentationSalesArea(ByRef CustomerCode As String, _
       ByRef CustomerType As String, ByRef Segmentation As String, ByRef SalesArea As String)
        Dim row As New DataGridViewRow
        Try
            If Me.DgvCustomerType.Rows.Count = 1 Then
                row = Me.DgvCustomerType.Rows(Me.DgvCustomerType.Rows.Count - 1)
                If String.IsNullOrEmpty(row.Cells(Me.ColCT_CustType.Index).Value) = False Then
                    CustomerType &= row.Cells(Me.ColCT_CustType.Index).Value
                Else
                    CustomerType = "All"
                End If
            Else
                For Each row In Me.DgvCustomerType.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColCT_CustType.Index).Value) = False Then
                        If CustomerType = "" Then
                            CustomerType = row.Cells(Me.ColCT_CustType.Index).Value
                        Else
                            CustomerType &= ", " & row.Cells(Me.ColCT_CustType.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvCustomer.Rows.Count = 1 Then
                row = Me.DgvCustomer.Rows(Me.DgvCustomer.Rows.Count - 1)
                If String.IsNullOrEmpty(row.Cells(Me.ColC_Code.Index).Value) = False Then
                    CustomerCode &= row.Cells(Me.ColC_Code.Index).Value
                Else
                    CustomerCode = "All"
                End If
            Else
                For Each row In Me.DgvCustomer.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColC_Code.Index).Value) = False Then
                        If CustomerCode = "" Then
                            CustomerCode = row.Cells(Me.ColC_Code.Index).Value
                        Else
                            CustomerCode &= ", " & row.Cells(Me.ColC_Code.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvMarketSegmentation.Rows.Count = 1 Then
                row = Me.DgvMarketSegmentation.Rows(Me.DgvMarketSegmentation.Rows.Count - 1)
                If String.IsNullOrEmpty(row.Cells(Me.ColMS_MarketSeg.Index).Value) = False Then
                    Segmentation &= row.Cells(Me.ColMS_MarketSeg.Index).Value
                Else
                    Segmentation = "All"
                End If
            Else
                For Each row In Me.DgvMarketSegmentation.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColMS_MarketSeg.Index).Value) = False Then
                        If Segmentation = "" Then
                            Segmentation = row.Cells(Me.ColMS_MarketSeg.Index).Value
                        Else
                            Segmentation &= ", " & row.Cells(Me.ColMS_MarketSeg.Index).Value
                        End If
                    End If
                Next
            End If

            If Me.DgvSalesArea.Rows.Count = 1 Then
                row = Me.DgvSalesArea.Rows(Me.DgvSalesArea.Rows.Count - 1)
                If String.IsNullOrEmpty(row.Cells(Me.ColSA_SalesArea.Index).Value) = False Then
                    SalesArea &= row.Cells(Me.ColSA_SalesArea.Index).Value
                Else
                    SalesArea = "All"
                End If
            Else
                For Each row In Me.DgvSalesArea.Rows
                    If String.IsNullOrEmpty(row.Cells(Me.ColSA_SalesArea.Index).Value) = False Then
                        If SalesArea = "" Then
                            SalesArea = row.Cells(Me.ColSA_SalesArea.Index).Value
                        Else
                            SalesArea &= ", " & row.Cells(Me.ColSA_SalesArea.Index).Value
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InsertSales(ByVal row As DseLaporanSalesStatistic.Group_By_Cust_CodeRow, _
    ByVal InvoiceDate As Date, ByVal SalesAmount As Double)
        Try
            Select Case InvoiceDate.Month
                Case 1
                    row.January += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 2
                    row.February += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 3
                    row.Maret += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 4
                    row.April += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 5
                    row.Mei += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 6
                    row.Juni += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 7
                    row.Juli += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 8
                    row.Agustus += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 9
                    row.September += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 10
                    row.Oktober += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 11
                    row.Nopember += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 12
                    row.Desember += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InsertSales(ByVal row As DseLaporanSalesStatistic.Group_By_Cust_Group_CodeRow, _
    ByVal InvoiceDate As Date, ByVal SalesAmount As Double)
        Try
            Select Case InvoiceDate.Month
                Case 1
                    row.January += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 2
                    row.February += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 3
                    row.Maret += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 4
                    row.April += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 5
                    row.Mei += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 6
                    row.Juni += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 7
                    row.Juli += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 8
                    row.Agustus += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 9
                    row.September += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 10
                    row.Oktober += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 11
                    row.Nopember += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
                Case 12
                    row.Desember += SalesAmount
                    row.Total_Sales += SalesAmount
                    Exit Sub
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanSalesStatistic_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ObjBrand As New Core_Module.ClsMasterProductGroup
        Dim ObjCustType As New Core_Module.clsMasterCustomerType
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim ObjSegmentation As New Core_Module.clsMasterCustomerMarketSegmentation
        Dim ObjSA As New Core_Module.clsMasterCustomerSalesArea

        Try
            'pastikan crystalreport terlihat
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True

            'tentukan tanggal periode
            Me.dtpSalesInvoicePeriod.Value = ObjCD.GetFirstDateofTheMonth(Today)
            'isikan semua checkbox
            '1. product brand
            'ObjBrand.IsiListBoxBrand(Me.clbBrand, TBLBrand)
            'Me.chkBrand.Checked = True
            ''2. cust type
            'ObjCustType.IsiListBoxCustType(Me.clbCustomerType, TBlCustType)
            'Me.chkCustType.Checked = True
            ''3. customer
            'ObjCust.IsiListBoxCustomer(Me.clbCustomer, TblCust)
            'Me.chkCustomer.Checked = True
            ''4. sales area
            'ObjSA.IsiListBoxCustomerSalesArea(Me.clbSalesArea, Me.TblSalesArea)
            'Me.chkSalesArea.Checked = True
            ''5. market segmentation
            'ObjSegmentation.IsiListBoxCustomerSegmentation(Me.clbMarketSegmentation, Me.TblSegmentation)
            'Me.chkMarketSegmentation.Checked = True
            Me.DgvCustomerType.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvCustomer.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvMarketSegmentation.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvSalesArea.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvBrand.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.ColMS_Primary.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCustType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'If Me.chkCustType.Checked = True Then
            '    Me.clbCustomerType.Enabled = False
            '    For i As Integer = 0 To Me.clbCustomerType.Items.Count - 1
            '        Me.clbCustomerType.SetItemCheckState(i, CheckState.Checked)
            '    Next
            'Else
            '    For i As Integer = 0 To Me.clbCustomerType.Items.Count - 1
            '        Me.clbCustomerType.SetItemCheckState(i, CheckState.Unchecked)
            '    Next
            '    Me.clbCustomerType.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'If Me.chkCustomer.Checked = True Then
            '    Me.clbCustomer.Enabled = False
            '    For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
            '        Me.clbCustomer.SetItemCheckState(i, CheckState.Checked)
            '    Next
            'Else
            '    For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
            '        Me.clbCustomer.SetItemCheckState(i, CheckState.Unchecked)
            '    Next
            '    Me.clbCustomer.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkBrand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'If Me.chkBrand.Checked = True Then
            '    Me.clbBrand.Enabled = False
            '    For i As Integer = 0 To Me.clbBrand.Items.Count - 1
            '        Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            '    Next
            'Else
            '    For i As Integer = 0 To Me.clbBrand.Items.Count - 1
            '        Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
            '    Next
            '    Me.clbBrand.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSalesArea_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'If Me.chkSalesArea.Checked = True Then
            '    Me.clbSalesArea.Enabled = False
            '    For i As Integer = 0 To Me.clbSalesArea.Items.Count - 1
            '        Me.clbSalesArea.SetItemCheckState(i, CheckState.Checked)
            '    Next
            'Else
            '    For i As Integer = 0 To Me.clbSalesArea.Items.Count - 1
            '        Me.clbSalesArea.SetItemCheckState(i, CheckState.Unchecked)
            '    Next
            '    Me.clbSalesArea.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkMarketSegmentation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'If Me.chkMarketSegmentation.Checked = True Then
            '    Me.clbMarketSegmentation.Enabled = False
            '    For i As Integer = 0 To Me.clbMarketSegmentation.Items.Count - 1
            '        Me.clbMarketSegmentation.SetItemCheckState(i, CheckState.Checked)
            '    Next
            'Else
            '    For i As Integer = 0 To Me.clbMarketSegmentation.Items.Count - 1
            '        Me.clbMarketSegmentation.SetItemCheckState(i, CheckState.Unchecked)
            '    Next
            '    Me.clbMarketSegmentation.Enabled = True
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub GenerateByCustGroup()
        Dim row1 As DseLaporanSalesStatistic.DataTable1Row
        Dim row2 As DseLaporanSalesStatistic.Group_By_Cust_Group_CodeRow
        Dim SalesAmount As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByCustGroupCodeView As DataView = Me.TBLGroupByCustGroupCode.DefaultView

        Try
            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To TBL.DefaultView.Count - 1
                row1 = TBL.DefaultView(i).Row
                'If Me.ApakahMasukKriteria(row1) = True Then
                'hitung exchage rate
                'CalculatesBy = row1.Calculates_By
                'If CalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
                'cmd = "[Primary Sales Invoice] = " + CStr(row1.Primary_Invoice) + _
                '" and [Primary Transaksi Rujukan] =" + CStr(row1.Primary_DO)
                'ElseIf CalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
                'cmd = "[Primary Sales Invoice] = " + CStr(row1.Primary_Invoice) + _
                '" and [Primary Transaksi Rujukan] =" + CStr(row1.Primary_Invoice)
                'ElseIf CalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
                'cmd = "[Primary Sales Invoice] = " + CStr(row1.Primary_Invoice) + _
                '" and [Primary Transaksi Rujukan] =" + CStr(row1.Primary_SO)
                'End If
                'TBL2View.RowFilter = cmd
                'If TBL2View.Count > 0 Then
                'rowRate = TBL2View.Item(0).Row
                'Rate = rowRate.Exhange_Rate
                'Else
                'Rate = 1
                'End If
                'hitung sales amount berdasarkan rate tsb
                'SalesAmount = Math.Round(Math.Round(row1.Sales_Amount * Rate, _
                'row1.Default_Precision, MidpointRounding.AwayFromZero) * row1.Tax_Exchange_Rate, _
                '0, MidpointRounding.AwayFromZero)
                SalesAmount = Math.Round(row1.Sales_Amount, 0, MidpointRounding.AwayFromZero)

                'masukkan hasilnya ke total sales ytd + baris
                cmd = "[Group Code] ='" + row1.Customer_Group_Code + "'"
                TBLGroupByCustGroupCodeView.RowFilter = cmd
                If TBLGroupByCustGroupCodeView.Count = 0 Then
                    'tambahkan baris baru
                    TBLGroupByCustGroupCode.Rows.Add(row1.Customer_Group_Code, row1.Cust_Group_Name, 0, 0, 0, _
                    0, 0, 0, 0, 0, 0, 0, 0, 0)
                    'masukkan data sales bulanan dan total sales
                    row2 = TBLGroupByCustGroupCode.Item(TBLGroupByCustGroupCode.Rows.Count - 1)
                    Me.InsertSales(row2, row1.Invoice_Date, SalesAmount)
                Else
                    row2 = TBLGroupByCustGroupCodeView.Item(0).Row
                    'tambahan sales amount ytd
                    Me.InsertSales(row2, row1.Invoice_Date, SalesAmount)
                End If
            Next

            'sorting
            If Me.rboSortByCode.Checked = True Then
                If Me.rboAscending.Checked = True Then
                    TBLGroupByCustGroupCodeView.Sort = "[Group Code] ASC"
                Else
                    TBLGroupByCustGroupCodeView.Sort = "[Group Code] DESC"
                End If
            Else
                If Me.rboAscending.Checked = True Then
                    TBLGroupByCustGroupCodeView.Sort = "[Total Sales] ASC"
                Else
                    TBLGroupByCustGroupCodeView.Sort = "[Total Sales] DESC"
                End If
            End If
            TBLGroupByCustGroupCodeView.RowFilter = ""

            For Each RowView In TBLGroupByCustGroupCodeView
                'pindahin ke tabel1
                row2 = RowView.Row
                With row2
                    Me.DS.Group_By_Cust_Group_Code1.Rows.Add(.Group_Code, .Group_Name, .January, .February, _
                    .Maret, .April, .Mei, .Juni, .Juli, .Agustus, .September, .Oktober, .Nopember, .Desember, _
                    .Total_Sales)
                End With
            Next
            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)

            ReportByCustGroup.SetDataSource(DS)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByCustGroup.SetParameterValue(ReportByCustGroup.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)

            Me.crv1.ReportSource = ReportByCustGroup
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim Buffer As String = ""
#Region "Customer Type"
    Dim EraseIndex_CT As Integer = -1
    Dim ObjCT As New Core_Module.clsMasterCustomerType
    Private Sub DgvCustomerType_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomerType.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColCT_CustType.Index Then
                Me.Buffer = Me.DgvCustomerType(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomerType.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomerType.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjCT As New Core_Module.clsMasterCustomerType
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCT_CustType.Index Then
                        .Cells(Me.ColCT_CustType.Index).Value = .Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper
                        If ObjCT.ApakahSudahAda(.Cells(Me.ColCT_CustType.Index).Value) = True Then
                            primary = ObjCT.GetPrimary(.Cells(Me.ColCT_CustType.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_CT(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColCT_Primary.Index).Value = primary
                                .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(primary)
                                .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterCustomerType
                            ij.txtCari.Text = Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerType <> -1 Then
                                index = Me.GetIndexItem_CT(ij.PrimaryCustomerType, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                    .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                    .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_CT = e.RowIndex
                                            Me.BackgroundWorker1.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColCT_Primary.Index).Value = ij.PrimaryCustomerType
                                        .Cells(Me.ColCT_CustType.Index).Value = ObjCT.GetTypeCode(ij.PrimaryCustomerType)
                                        .Cells(Me.ColCT_Desc.Index).Value = ObjCT.GetTypeDesc(ij.PrimaryCustomerType)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerType = -1 Then
                                If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_CT = e.RowIndex
                                        Me.BackgroundWorker1.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomerType)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomerType)
                        Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value = Me.DgvCustomerType.Item(Me.ColCT_CustType.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustomerType.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColCT_CustType.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_CT = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            Else
                                Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomerType)
                            Exit Sub
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColCT_CustType.Index Then
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvCustomerType.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvCustomerType_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomerType.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomerType_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomerType.Sorted
        Try
            ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvCustomerType.Rows.RemoveAt(EraseIndex_CT)
            Me.ObjCD.SetNumber(Me.DgvCustomerType)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_CT(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomerType.Rows.Count - 1
                RowDGV = Me.DgvCustomerType.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCT_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
#End Region

#Region "Customer"
    Dim EraseIndex_C As Integer = -1
    Dim ObjCust As New Core_Module.clsMasterCustomer
    Private Sub DgvCustomer_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvCustomer.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColC_Code.Index Then
                Me.Buffer = Me.DgvCustomer(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCustomer.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvCustomer.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjRef As New Core_Module.clsMasterCustomer
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColC_Code.Index Then
                        .Cells(Me.ColC_Code.Index).Value = .Cells(Me.ColC_Code.Index).Value.ToString.ToUpper
                        If ObjRef.ApakahSudahAda(.Cells(Me.ColC_Code.Index).Value) = True Then
                            primary = ObjRef.GetPrimary(.Cells(Me.ColC_Code.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_C(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColC_Primary.Index).Value = primary
                                .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(primary)
                                .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_C = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomer)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomer
                            ij.txtCari.Text = Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomer <> -1 Then
                                index = Me.GetIndexItem_C(ij.PrimaryCustomer, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColC_Primary.Index).Value = ij.PrimaryCustomer
                                    .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                    .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_C = e.RowIndex
                                            Me.BackgroundWorker2.RunWorkerAsync()
                                        Else
                                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvCustomer)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColC_Primary.Index).Value = ij.PrimaryCustomer
                                        .Cells(Me.ColC_Code.Index).Value = ObjRef.GetCustomerCode(ij.PrimaryCustomer)
                                        .Cells(Me.ColC_Desc.Index).Value = ObjRef.GetCustName(ij.PrimaryCustomer)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomer = -1 Then
                                If Me.DgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_C = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvCustomer)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvCustomer)
                        Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value = Me.DgvCustomer.Item(Me.ColC_Code.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvCustomer.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColC_Code.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_C = e.RowIndex
                                Me.BackgroundWorker2.RunWorkerAsync()
                            Else
                                Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvCustomer)
                            Exit Sub
                        Else
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColC_Code.Index Then
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvCustomer.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvCustomer_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvCustomer.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvCustomer_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvCustomer.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvCustomer, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvCustomer.Rows.RemoveAt(EraseIndex_C)
            Me.ObjCD.SetNumber(Me.DgvCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_C(ByVal vPrimary As Double, _
  ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvCustomer.Rows.Count - 1
                RowDGV = Me.DgvCustomer.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColC_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
#End Region

#Region "Market Segmentasi"
    Dim EraseIndex_MS As Integer = -1
    Dim ObjMS As New Core_Module.clsMasterCustomerMarketSegmentation
    Private Sub DgvMarketSegmentation_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvMarketSegmentation.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                Me.Buffer = Me.DgvMarketSegmentation(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvMarketSegmentation_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvMarketSegmentation.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvMarketSegmentation.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjMS As New Core_Module.clsMasterCustomerMarketSegmentation
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                        .Cells(Me.ColMS_MarketSeg.Index).Value = .Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper
                        If ObjMS.ApakahSudahAda(.Cells(Me.ColMS_MarketSeg.Index).Value) = True Then
                            primary = ObjMS.GetPrimary(.Cells(Me.ColMS_MarketSeg.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_MS(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColMS_Primary.Index).Value = primary
                                .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(primary)
                                .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterCustomerSegmentation
                            ij.txtCari.Text = Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerMarketSegmentation <> -1 Then
                                index = Me.GetIndexItem_MS(ij.PrimaryCustomerMarketSegmentation, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                    .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                    .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_MS = e.RowIndex
                                            Me.BackgroundWorker3.RunWorkerAsync()
                                        Else
                                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColMS_Primary.Index).Value = ij.PrimaryCustomerMarketSegmentation
                                        .Cells(Me.ColMS_MarketSeg.Index).Value = ObjMS.GetSegmentationCode(ij.PrimaryCustomerMarketSegmentation)
                                        .Cells(Me.ColMS_Desc.Index).Value = ObjMS.GetDesc(ij.PrimaryCustomerMarketSegmentation)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerMarketSegmentation = -1 Then
                                If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_MS = e.RowIndex
                                        Me.BackgroundWorker3.RunWorkerAsync()
                                    Else
                                        Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                        Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value = Me.DgvMarketSegmentation.Item(Me.ColMS_MarketSeg.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvMarketSegmentation.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColMS_MarketSeg.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_MS = e.RowIndex
                                Me.BackgroundWorker3.RunWorkerAsync()
                            Else
                                Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
                            Exit Sub
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColMS_MarketSeg.Index Then
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvMarketSegmentation.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_MS(ByVal vPrimary As Double, _
     ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvMarketSegmentation.Rows.Count - 1
                RowDGV = Me.DgvMarketSegmentation.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColMS_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvMarketSegmentation_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvMarketSegmentation.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.DgvMarketSegmentation.Rows.RemoveAt(EraseIndex_MS)
            Me.ObjCD.SetNumber(Me.DgvMarketSegmentation)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region
   
#Region "Sales Area"
    Dim EraseIndex_SA As Integer = -1
    Private Sub DgvSalesArea_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvSalesArea.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                Me.Buffer = Me.DgvSalesArea(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvSalesArea.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvSalesArea.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjSA As New Core_Module.clsMasterCustomerSalesArea
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                        .Cells(Me.ColSA_SalesArea.Index).Value = .Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper
                        If ObjSA.ApakahSudahAda(.Cells(Me.ColSA_SalesArea.Index).Value) = True Then
                            primary = ObjSA.GetPrimary(.Cells(Me.ColSA_SalesArea.Index).Value.ToString.ToUpper)
                            index = Me.GetIndexItem_SA(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColSA_Primary.Index).Value = primary
                                .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(primary)
                                .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterSalesArea
                            ij.txtCari.Text = Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCustomerSalesArea <> -1 Then
                                index = Me.GetIndexItem_SA(ij.PrimaryCustomerSalesArea, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                    .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                    .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_SA = e.RowIndex
                                            Me.BackgroundWorker4.RunWorkerAsync()
                                        Else
                                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColSA_Primary.Index).Value = ij.PrimaryCustomerSalesArea
                                        .Cells(Me.ColSA_SalesArea.Index).Value = ObjSA.GetSalesAreaCode(ij.PrimaryCustomerSalesArea)
                                        .Cells(Me.ColSA_Desc.Index).Value = ObjSA.GetSalesAreaDesc(ij.PrimaryCustomerSalesArea)
                                    End If
                                End If
                            ElseIf ij.PrimaryCustomerSalesArea = -1 Then
                                If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_SA = e.RowIndex
                                        Me.BackgroundWorker4.RunWorkerAsync()
                                    Else
                                        Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvSalesArea)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvSalesArea)
                        Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value = Me.DgvSalesArea.Item(Me.ColSA_SalesArea.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvSalesArea.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColSA_SalesArea.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_SA = e.RowIndex
                                Me.BackgroundWorker4.RunWorkerAsync()
                            Else
                                Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvSalesArea)
                            Exit Sub
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColSA_SalesArea.Index Then
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvSalesArea.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub DgvSalesArea_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvSalesArea.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvSalesArea_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvSalesArea.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvSalesArea, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            Me.DgvSalesArea.Rows.RemoveAt(EraseIndex_SA)
            Me.ObjCD.SetNumber(Me.DgvSalesArea)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem_SA(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvSalesArea.Rows.Count - 1
                RowDGV = Me.DgvSalesArea.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColSA_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
#End Region

#Region "Brand"
    Dim EraseIndex_B As Integer = -1
    Dim ProductGroupLevel As Integer = 2
    Private Sub DgvBrand_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvBrand.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColB_Brand.Index Then
                Me.Buffer = Me.DgvBrand(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvBrand_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvBrand.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvBrand.Rows(e.RowIndex)
        Dim primary As Double = -1
        Dim index As Integer = -1
        Dim Identifier_Process As String = ""
        Dim ObjB As New Core_Module.ClsMasterProductGroup
        Try
            Dim vStrTipe As String = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            With row
                If IsNothing(Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColB_Brand.Index Then
                        .Cells(Me.ColB_Brand.Index).Value = .Cells(Me.ColB_Brand.Index).Value.ToString.ToUpper
                        If ObjB.ApakahSudahAda(.Cells(Me.ColB_Brand.Index).Value, Me.ProductGroupLevel) = True Then
                            primary = ObjB.GetPrimary(.Cells(Me.ColB_Brand.Index).Value.ToString.ToUpper, Me.ProductGroupLevel)
                            index = Me.GetIndexItem_B(primary, e.RowIndex)

                            If index < 0 Then
                                .Cells(Me.ColB_Primary.Index).Value = primary
                                .Cells(Me.ColB_Brand.Index).Value = ObjB.GetGroupCode(primary)
                                .Cells(Me.ColB_Desc.Index).Value = ObjB.GetGroupDesc(primary)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvBrand.Item(Me.ColB_Brand.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_B = e.RowIndex
                                        Me.BackgroundWorker5.RunWorkerAsync()
                                    Else
                                        Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvBrand)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterProductGroup
                            ij.ProductGroupLevel = Me.ProductGroupLevel
                            ij.txtCari.Text = Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryProductGroup <> -1 Then
                                index = Me.GetIndexItem_B(ij.PrimaryProductGroup, e.RowIndex)
                                If index < 0 Then
                                    .Cells(Me.ColB_Primary.Index).Value = ij.PrimaryProductGroup
                                    .Cells(Me.ColB_Brand.Index).Value = ObjB.GetGroupCode(ij.PrimaryProductGroup)
                                    .Cells(Me.ColB_Desc.Index).Value = ObjB.GetGroupDesc(ij.PrimaryProductGroup)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                                            Me.EraseIndex_B = e.RowIndex
                                            Me.BackgroundWorker5.RunWorkerAsync()
                                        Else
                                            Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                        End If
                                        Me.ObjCD.SetNumber(Me.DgvBrand)
                                        Exit Sub
                                    Else
                                        .Cells(Me.ColB_Primary.Index).Value = ij.PrimaryProductGroup
                                        .Cells(Me.ColB_Brand.Index).Value = ObjB.GetGroupCode(ij.PrimaryProductGroup)
                                        .Cells(Me.ColB_Desc.Index).Value = ObjB.GetGroupDesc(ij.PrimaryProductGroup)
                                    End If
                                End If
                            ElseIf ij.PrimaryProductGroup = -1 Then
                                If Me.DgvBrand.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseIndex_B = e.RowIndex
                                        Me.BackgroundWorker5.RunWorkerAsync()
                                    Else
                                        Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvBrand)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvBrand)
                        Me.DgvBrand.Item(Me.ColB_Brand.Index, e.RowIndex).Value = Me.DgvBrand.Item(Me.ColB_Brand.Index, e.RowIndex).Value.ToString.ToUpper
                    End If
                Else
                    If Me.DgvBrand.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex = Me.ColB_Brand.Index Then
                            If String.IsNullOrEmpty(Me.Buffer) = True Then
                                Me.EraseIndex_B = e.RowIndex
                                Me.BackgroundWorker5.RunWorkerAsync()
                            Else
                                Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            End If
                            Me.ObjCD.SetNumber(Me.DgvBrand)
                            Exit Sub
                        Else
                            Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.ColB_Brand.Index Then
                            Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvBrand.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Function GetIndexItem_B(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvBrand.Rows.Count - 1
                RowDGV = Me.DgvBrand.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColB_Primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvBrand_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvBrand.RowsRemoved
        Try
            ObjCD.SetNumber(Me.DgvBrand)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvBrand_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvBrand.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvBrand, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            Me.DgvBrand.Rows.RemoveAt(EraseIndex_B)
            Me.ObjCD.SetNumber(Me.DgvBrand)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region
   
End Class