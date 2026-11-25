Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesAnalysis
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDecimal As Integer = 0

    Dim DS As New DseLaporanSalesAnalysis
    Dim ReportByBrand As New CrpLaporanSalesAnalysisByBrand
    Dim ReportByBrand_Return As New CrpLaporanSalesAnalysisByBrand_Return
    Dim ReportByCustType As New CrpLaporanSalesAnalysisByCustomerType
    Dim ReportByMarketSegmentation As New CrpLaporanSalesAnalysisByMarketSegmentation
    Dim ReportBySalesAreaDetail As New CrpLaporanSalesAnalysisBySalesAreaDetail
    Dim ReportBySalesAreaRekap As New CrpLaporanSalesAnalysisBySalesAreaRekap
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjPurchaseClosing As New ClsTransaksiPurchaseClosing
    Dim ObjCur As New Core_Module.clsMasterCurrency

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

    Dim TBL As DseLaporanSalesAnalysis.DataTable1DataTable = DS.DataTable1
    Dim TBL2 As DseLaporanSalesAnalysis.DataTable2DataTable = DS.DataTable2
    Dim TBL3 As DseLaporanSalesAnalysis.DataTable3DataTable = DS.DataTable3
    Dim TBLDO As DseLaporanSalesAnalysis.tbl_trans_sales_area_deliveryDataTable = DS.tbl_trans_sales_area_delivery
    Dim TBLGroupBySalesAreaRekap As DseLaporanSalesAnalysis._Group_By_Market_Area_RekapDataTable = DS._Group_By_Market_Area_Rekap
    Dim TBLGroupBySalesAreaDetail As DseLaporanSalesAnalysis._Group_By_Market_Area_DetailDataTable = DS._Group_By_Market_Area_Detail
    Dim TBLGroupByCustType As DseLaporanSalesAnalysis.Group_By_Customer_TypeDataTable = DS.Group_By_Customer_Type
    Public TBLGroupByBrand As DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable = DS.Group_By_Product_Brand
    Dim TBLGroupBySegmentation As DseLaporanSalesAnalysis.Group_By_Customer_Market_SegmentationDataTable = DS.Group_By_Customer_Market_Segmentation
    Dim StartPeriode, StartPeriodYTD, EndPeriod As Date
    Dim TBLReturn As New DseLaporanSalesAnalysis.TblReturnInvoiceDataTable
    Dim DTAReturn As New DseLaporanSalesAnalysisTableAdapters.TblReturnInvoiceTableAdapter
    Dim rowReturn As DseLaporanSalesAnalysis.TblReturnInvoiceRow
    Dim ObjClosing As New ClsTransaksiMonthlyClosing

    Private Enum EnumKriteriaPerhitunganCostOfSales As Integer
        ByBrand
        ByCustType
        ByMarketSegmentation
        BySalesAreaRekap
        BySalesAreaDetail
        BySalesAreaDetailCategory
        ByCategory
    End Enum

    Private Sub KurangkanItemReturn()
        Dim Filter As String = ""
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
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

    Private Sub FrmLaporanSalesAnalysis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ObjBrand As New Core_Module.ClsMasterProductGroup
        Dim ObjCustType As New Core_Module.clsMasterCustomerType
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim ObjSegmentation As New Core_Module.clsMasterCustomerMarketSegmentation
        Dim ObjSA As New Core_Module.clsMasterCustomerSalesArea

        Try
            Me.FormatDecimal = Me.ObjCur.GetDecimal_DefaultCurrency
            'pastikan crystalreport terlihat
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.ShowGroupTreeButton = False
            'tentukan tanggal periode
            Me.dtpSalesInvoicePeriod.Value = ObjCD.GetFirstDateofTheMonth(Today)
            'isikan semua checkbox
            ''1. product brand
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
            Me.cboDeliveryType.Visible = False
            Me.cboDeliveryType.TabStop = False
            Me.ChkMarkGM.Checked = True
            Me.ChkMarkGM.Checked = False
            Me.DgvCustomerType.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvCustomer.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvMarketSegmentation.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvSalesArea.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.DgvBrand.AlternatingRowsDefaultCellStyle.BackColor = BackColorAll
            Me.CboFilterGM.SelectedIndex = 0
            Me.ColMS_Primary.Visible = False
            Me.TabPage1.Text = "Data Filter"
            Me.ChkMarkGM.Left = Me.chkService.Left + Me.chkService.Width + 5
            Me.chkIncludeSalesReturn.Left = Me.ChkMarkGM.Left
            Me.CboFilterGM.Left = Me.ChkMarkGM.Left + Me.ChkMarkGM.Width + 5
            Me.txtGMPercentage.Left = Me.CboFilterGM.Left + Me.CboFilterGM.Width + 5
            Me.Label2.Left = Me.txtGMPercentage.Left + Me.txtGMPercentage.Width + 5
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Function ApakahMasukKriteria(ByVal row1 As DseLaporanSalesAnalysis.DataTable1Row) As Boolean
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

    Private Function ApakahMasukKriteria(ByVal row1 As DseLaporanSalesAnalysis.DataTable5Row) As Boolean
        Dim ketemu As Boolean = False
        Dim ProductType As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            ''cek cust type
            'If Me.chkCustType.Checked = False Then
            '    ketemu = False
            '    For Each RowView In Me.clbCustomerType.CheckedItems
            '        rowCustType = RowView.Row
            '        If rowCustType.Type_Code = row1.Cust_Type_Code Then
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
            '        If rowBrand.Brand = row1.Brand_Code Then
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
            'cek tipe product
            ProductType = row1.Type
            If Me.chkGoods.Checked = False And ProductType = Core_Module.clsMasterProduct.EnumTipeProduct.Goods Then
                Return False
            End If
            If Me.chkService.Checked = False And ProductType = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                Return False
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim DTA As New DseLaporanSalesAnalysisTableAdapters.DataTable1TableAdapter
        Dim DTA2 As New DseLaporanSalesAnalysisTableAdapters.DataTable2TableAdapter
        Dim DTA3 As New DseLaporanSalesAnalysisTableAdapters.DataTable3TableAdapter
        Dim DTADO As New DseLaporanSalesAnalysisTableAdapters.tbl_trans_sales_area_deliveryTableAdapter
        Dim TipeProduct As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Dim row As DseLaporanSalesAnalysis.DataTable1Row
        Dim rowDO As DseLaporanSalesAnalysis.tbl_trans_sales_area_deliveryRow

        Dim PrimaryDO_Max As Integer = 0
        Dim PrimaryDO_Min As Integer = 0

        Dim Count As Integer = 0
        Dim Dv As DataRowView
        Dim Dv_2 As DataRowView
        Try
            If Me.ApakahBolehCetak = True Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                Me.KosongkanTabel()
                'set awal dan akhir periode perhitungan
                Me.StartPeriode = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSalesInvoicePeriod.Value)
                Me.EndPeriod = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpSalesInvoicePeriod.Value)
                If Me.rboByBrandWithReturn.Checked = False Then
                    Me.StartPeriodYTD = CDate("1/1/" + CStr(Me.dtpSalesInvoicePeriod.Value.Year))
                Else
                    Me.StartPeriodYTD = Me.StartPeriode
                End If

                'isikan tabel yang akan digunakan untuk perhitungan

                If Me.chkGoods.Checked = True And Me.chkService.Checked = True Then
                    If Me.rboByBrand.Checked = True Then
                        DTA.FillBy_BrandOrder(TBL, Me.EndPeriod, Me.StartPeriodYTD)
                    ElseIf Me.rboByProductCategory.Checked = True Then
                        DTA.FillBy_CategoryCode(TBL, Me.EndPeriod, Me.StartPeriodYTD)
                    Else
                        DTA.Fill_TanggalSampaiDari(TBL, Me.EndPeriod, Me.StartPeriodYTD)
                    End If
                Else
                    If Me.chkGoods.Checked = True Then
                        TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
                    Else
                        TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
                    End If
                    If Me.rboByBrand.Checked = True Then
                        DTA.FillBy_TipeProduct3(TBL, Me.EndPeriod, Me.StartPeriodYTD, CInt(TipeProduct))
                    ElseIf Me.rboByProductCategory.Checked = True Then
                        DTA.FillBy_TipeProduct2(TBL, Me.EndPeriod, Me.StartPeriodYTD, CInt(TipeProduct))
                    Else
                        DTA.FillBy_TipeProduct1(TBL, Me.EndPeriod, Me.StartPeriodYTD, CInt(TipeProduct))
                    End If
                End If

                DTA2.Fill(TBL2, Me.EndPeriod, Me.StartPeriodYTD)
                'Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
                'For Each row1 In Me.TBL.Rows
                '    If row1.Sales_Amount < 0 Then
                '        MsgBox("")
                '    End If
                'Next
                DTA3.FillBy_TanggalSampaiDari(Me.TBL3, Me.EndPeriod, Me.StartPeriodYTD)
                Me.DTAReturn.Fill(Me.TBLReturn, Me.EndPeriod, Me.StartPeriodYTD)
                Me.KurangkanItemReturn()

                'set ulang sales area jika berdasarkan delivery sales area
                If Me.cboDeliveryType.SelectedIndex = 1 Then
                    'dapatkan primary DO Min Max
MinMax:
                    If Count = 0 Then
                        TBL.DefaultView.Sort = "[Primary DO]"
                    Else
                        TBL.DefaultView.Sort = "[Primary DO] DESC"
                    End If
                    For Each Dv In TBL.DefaultView
                        row = Dv.Row
                        Count += 1
                        If Count = 1 Then
                            PrimaryDO_Min = row.Primary_DO
                            GoTo MinMax
                        Else
                            PrimaryDO_Max = row.Primary_DO
                            Exit For
                        End If
                    Next

                    'fill tbl_trans_sales_area_delivery berdasarkan primary DO MinMax
                    DTADO.FillBy_PrimaryDOMinMax(TBLDO, PrimaryDO_Min, PrimaryDO_Max)

                    For Each row In TBL.Rows
                        With row
                            TBLDO.DefaultView.RowFilter = Nothing
                            TBLDO.DefaultView.RowFilter = "[Primary DO] = " & .Primary_DO
                            If TBLDO.DefaultView.Count > 0 Then
                                'ubah sales area code-desc dengan sales area code-desc DO
                                For Each Dv_2 In TBLDO.DefaultView
                                    rowDO = Dv_2.Row
                                    With rowDO
                                        row.Sales_Area_Code = .Sales_Area_Code
                                        row.Sales_Area_Description = .Sales_Area_Desc
                                    End With
                                Next
                            End If
                        End With
                    Next

                    DTADO.Dispose()
                End If

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

                'perbarui nilai cost
                'Me.MasukkanNilaiCost()

                'set nilai maksimum progress bar
                Me.pbr1.Value = 0
                Me.pbr1.Maximum = TBL.Rows.Count
                'lakukan perhitungan
                If Me.rboByBrand.Checked = True Then
                    Me.GenerateByBrand()
                ElseIf Me.rboByProductCategory.Checked = True Then
                    Me.GenerateByCategory()
                ElseIf Me.rboByCustomerType.Checked = True Then
                    Me.GenerateByCustType()
                ElseIf Me.rboByMarketSegmentation.Checked = True Then
                    Me.GenerateByMarketSegmentation()
                ElseIf Me.rboByMarketSegmentationWithReturn.Checked = True Then
                    Me.GenerateByMarketSegmentationWithReturn()
                ElseIf Me.rboSalesAreaDetail.Checked = True Then
                    Me.GenerateBySalesAreaDetail()
                ElseIf Me.rboSalesAreaDetailCategory.Checked = True Then
                    Me.GenerateBySalesAreaDetail2()
                ElseIf Me.rboSalesAreaRekap.Checked = True Then
                    Me.GenerateBySalesAreaRekap()
                ElseIf Me.rboProductCode.Checked = True Then
                    Me.GenerateByProductCode()
                ElseIf Me.rboProductCodeWithReturn.Checked = True Then
                    Me.GenerateByProductCodeWithReturn()
                    'ElseIf Me.rboByInvoiceNumber.Checked = True Then
                    '    Me.GenerateByInvoiceNumber()
                ElseIf Me.rboByBrandWithReturn.Checked = True Then
                    Me.GenerateByBrandWithReturn()
                ElseIf Me.rboByProductCategorywithReturn.Checked = True Then
                    Me.GenerateByCategoryWithReturn()
                ElseIf Me.rboByCustomerTypeWithReturn.Checked = True Then
                    Me.GenerateByCustTypeWithReturn()
                ElseIf Me.rboBySalesAreaWithReturn.Checked = True Then
                    Me.GenerateBySalesAreaWithReturn()
                End If

                Me.TabControl1.SelectedIndex = 1
            End If

            Me.crv1.ShowExportButton = True
            Me.crv1.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    'Private Sub MasukkanNilaiCost()
    '    Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosing.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosing.EnumTipeTranskasiAsal.OutSales
    '    Dim TipeItem As ClsTransaksiMonthlyClosing.EnumTipeItemDetail = ClsTransaksiMonthlyClosing.EnumTipeItemDetail.OutItem
    '    Dim DTA_CL As New DseLaporanSalesAnalysisTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
    '    Dim TBL_CL As New DseLaporanSalesAnalysis.tbltrans_purchase_closing_item_detailDataTable
    '    Dim row_CL As DseLaporanSalesAnalysis.tbltrans_purchase_closing_item_detailRow
    '    Dim Filter As String = ""
    '    Dim JmlDV As Double = 0
    '    'Dim DV As DataRowView
    '    Dim rowItem As DseLaporanSalesAnalysis.DataTable1Row
    '    Dim CostDummy As Decimal = 0
    '    Dim DefPrec As Integer = 0

    '    Try
    '        DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
    '        DTA_CL.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_CL.Fill(TBL_CL, Me.EndPeriod, Me.StartPeriodYTD, TipeTransaksiAsal, TipeItem)

    '        For Each row_CL In TBL_CL.Rows
    '            Filter = "[Primary Invoice] = " + row_CL.primary_transaksi_asal.ToString + _
    '            " AND [Primary Product] = " + row_CL.primary_code_no.ToString

    '            Me.TBL.DefaultView.RowFilter = Filter
    '            JmlDV = 0
    '            CostDummy = 0

    '            If Me.TBL.DefaultView.Count > 0 Then
    '                For JmlDV = 0 To Me.TBL.DefaultView.Count - 1
    '                    rowItem = Me.TBL.DefaultView(JmlDV).Row

    '                    If JmlDV = Me.TBL.DefaultView.Count - 1 Then
    '                        rowItem.Cost = row_CL.amount - CostDummy
    '                    Else
    '                        rowItem.Cost = ObjCD.DigitGrouping((row_CL.amount * rowItem.Total_Qty) / row_CL.qty, DefPrec)
    '                        CostDummy += rowItem.Cost
    '                    End If
    '                Next
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA_CL.Dispose()
    '        TBL_CL.Dispose()
    '    End Try
    'End Sub

    Private Sub KosongkanTabel()
        Try
            Me.DS.DataTable1.Rows.Clear()
            Me.DS.DataTable2.Rows.Clear()
            Me.DS.DataTable3.Rows.Clear()
            Me.DS.Group_By_Customer_Market_Segmentation.Rows.Clear()
            Me.DS.Group_By_Customer_Type.Rows.Clear()
            Me.DS.Group_By_Product_Brand.Rows.Clear()
            Me.DS._Group_By_Market_Area_Detail.Rows.Clear()
            Me.DS._Group_By_Market_Area_Rekap.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GenerateByBrand()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
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
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount

                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Brand) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupByBrand.Rows.Add(row1.Brand, row1.Brand_Description, 0, 0, 0, _
                        0, SalesAmount, 0, row1.Cost, 0)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

                For Each row2 In Me.TBLGroupByBrand.Rows
                    With row2
                        .Group_Name = row2.Product_Brand & " - " & row2.Product_Brand_Name
                        'hitung % total sales
                        If TotalSales <> 0 Then
                            .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                        Else
                            .Total_Sales_Percentage = 0
                        End If
                        'hitung gm percentage
                        If .Sales_Amount = 0 Then
                            .GM_Percentage = 0
                        Else
                            .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                            If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                                .GM_Percentage = .GM_Percentage * (-1)
                            End If
                        End If
                        'hitung % total sales YTD
                        If TotalSalesYTD <> 0 Then
                            .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                        Else
                            .Total_Sales_YTD_Percentage = 0
                        End If
                        'hitung GM% YTD
                        If .Sales_Amount_YTD = 0 Then
                            .GM_Percentage_YTD = 0
                        Else
                            .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                            .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                            If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                                .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                            End If
                        End If
                    .IsColor = Me.GetIsColor(.GM_Percentage, .Sales_Amount)
                    End With
                Next

                'generate report
                Dim CustomerCode As String = ""
                Dim CustomerType As String = ""
                Dim Segmentation As String = ""
                Dim SalesArea As String = ""
                Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
                CustomerType, Segmentation, SalesArea)
                ReportByBrand.SetDataSource(DS)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerType.ParameterFieldName, CustomerType)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_SalesArea.ParameterFieldName, SalesArea)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Segmentation.ParameterFieldName, Segmentation)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY PRODUCT BRAND")
                ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Group_Name.ParameterFieldName, "PRODUCT BRAND")

                Me.crv1.ReportSource = ReportByBrand
                Me.crv1.DisplayGroupTree = False
                Me.crv1.DisplayStatusBar = True
                Me.crv1.DisplayToolbar = True
                Me.crv1.ShowGroupTreeButton = False
                Me.crv1.ShowExportButton = True
                Me.crv1.ShowRefreshButton = False
                Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
        End Try
    End Sub

    Private Sub IsikanReturnCost_AmountNonIDR(ByVal vTBLSR_Raw As DseLaporanSalesAnalysis.DataTable5DataTable, ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim DTASR As New DseLaporanSalesAnalysisTableAdapters.DataTable6TableAdapter
        Dim TBLSR As New DseLaporanSalesAnalysis.DataTable6DataTable
        Dim rowSR As DseLaporanSalesAnalysis.DataTable6Row
        'Dim Cost As Decimal = 0
        Dim DV As DataRowView
        Dim DTASR1 As New DseLaporanSalesPerformanceTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBLSR1 As New DseLaporanSalesPerformance.tbltrans_salesreturnDataTable
        Dim rowSR1 As DseLaporanSalesPerformance.tbltrans_salesreturnRow
        Dim SubTotal As Decimal = 0
        Dim DefPrec As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim Kombinasi As String = ""
        Dim QtyAllocated As Decimal = 0
        Dim CostAllocated As Decimal = 0
        Dim CostTotal As Decimal = 0
        Dim QtyTotal As Decimal = 0

        Try
            DTASR.Connection.ConnectionString = Core_Module.GetConstring
            DTASR1.Connection.ConnectionString = DTASR.Connection.ConnectionString

            DTASR.Fill(TBLSR, Me.ObjClosing.GetPrimaryClosing(Me.ObjCD.GetFirstDateofTheMonth(vStartDate)), _
            ClsTransaksiMonthlyClosing.EnumTipeTranskasiAsal.InSR)
            For Each rowSR_Raw In vTBLSR_Raw.Rows
                If Kombinasi <> rowSR_Raw.Primary_Stock_In_Sales_Return.ToString + "-" + rowSR_Raw.Primary_Product.ToString Then
                    CostTotal = 0
                    QtyTotal = 0
                    Kombinasi = rowSR_Raw.Primary_Stock_In_Sales_Return.ToString + "-" + rowSR_Raw.Primary_Product.ToString
                    CostAllocated = 0
                    QtyAllocated = 0

                    TBLSR.DefaultView.RowFilter = "[Primary Stock In Sales Return] = " + rowSR_Raw.Primary_Stock_In_Sales_Return.ToString + " AND " + _
                                    "[Primary Product] = " + rowSR_Raw.Primary_Product.ToString
                    If TBLSR.DefaultView.Count > 0 Then
                        For Each DV In TBLSR.DefaultView
                            rowSR = DV.Row
                            CostTotal += rowSR.Cost
                            QtyTotal += rowSR.Quantity
                        Next
                    End If
                End If
                
                If QtyAllocated + rowSR_Raw.Total_Qty = QtyTotal Then
                    rowSR_Raw.Cost = CostTotal - CostAllocated
                Else
                    If QtyTotal = 0 Then
                        'qty bisa =0 klo jasa
                        rowSR_Raw.Cost = Math.Round(0, DefPrec, MidpointRounding.AwayFromZero)
                    Else
                        rowSR_Raw.Cost = Math.Round((CostTotal * rowSR_Raw.Total_Qty) / QtyTotal, DefPrec, MidpointRounding.AwayFromZero)
                    End If
                End If
                CostAllocated += rowSR_Raw.Cost
                QtyAllocated += rowSR_Raw.Total_Qty
            Next

            'isikan data return amount yang currencynya bukan IDR
            DTASR1.Fill(TBLSR1, vEndDate, vStartDate, Core_Module.PrimaryBaseCurrency, CInt(TipeSR))
            For Each rowSR1 In TBLSR1.Rows
                rowSR1.Sub_Total_Base = Math.Round(rowSR1.Sub_Total * rowSR1.Tax_Exchange_Rate, DefPrec, MidpointRounding.AwayFromZero)

                vTBLSR_Raw.DefaultView.RowFilter = "[Primary SR] = " + rowSR1.Primary.ToString
                SubTotal = 0
                If vTBLSR_Raw.DefaultView.Count > 0 Then
                    For ij As Integer = 0 To vTBLSR_Raw.DefaultView.Count - 1
                        rowSR_Raw = vTBLSR_Raw.DefaultView(ij).Row
                        If ij = vTBLSR_Raw.DefaultView.Count - 1 Then
                            rowSR_Raw.Amount = rowSR1.Sub_Total_Base - SubTotal
                        Else
                            'bukan baris terakhir, hitung biasa, catatkan sub total
                            If rowSR_Raw.Primary_Currency <> Core_Module.PrimaryBaseCurrency Then
                                rowSR_Raw.Amount = Math.Round(rowSR_Raw.Amount * rowSR_Raw.Exchange_Rate, DefPrec, _
                                    MidpointRounding.AwayFromZero)
                                SubTotal += rowSR_Raw.Amount
                            End If
                        End If
                    Next
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR.Dispose()
            TBLSR.Dispose()
            DTASR1.Dispose()
            TBLSR1.Dispose()
        End Try
    End Sub

    Private Sub GenerateByCategoryWithReturn()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView

        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        'Dim TipeStockInSR As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim TotalNettSales, TotalNettCOGS As Decimal
        Dim TblDummy As New DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable
        Dim rowDummy As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim DV As DataRowView

        Try
            TotalNettCOGS = 0
            TotalNettSales = 0

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount

                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Category_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = row1.Cost
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = row1.Category_Code
                            .Product_Brand_Name = row1.Category_Desc
                            .Retur_Sales = 0
                            .Return_COGS = 0
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = SalesAmount
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                End With
            Next

            'isikan data SR
            DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.EndPeriod, Me.StartPeriode)
            Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, Me.StartPeriode, Me.EndPeriod)
            TBLSR_Raw.DefaultView.RowFilter = Me.FilterByDgvItem2
            For i As Integer = 0 To TBLSR_Raw.DefaultView.Count - 1
                rowSR_Raw = TBLSR_Raw.DefaultView(i).Row
                If Me.ApakahMasukKriteria(rowSR_Raw) = True Then
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(rowSR_Raw.Category_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = 0
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = rowSR_Raw.Category_Code
                            .Product_Brand_Name = rowSR_Raw.Category_Desc
                            .Retur_Sales = rowSR_Raw.Amount
                            .Return_COGS = rowSR_Raw.Cost
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = 0
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Retur_Sales += rowSR_Raw.Amount
                        row2.Return_COGS += rowSR_Raw.Cost
                    End If
                End If
            Next

            'hitung data yang belum lengkap
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_SalesNull = True Then
                    row2.Nett_Sales = 0
                End If
                If row2.IsRetur_SalesNull Then
                    row2.Retur_Sales = 0
                End If
                'hitung nett sales
                row2.Nett_Sales = row2.Sales_Amount - row2.Retur_Sales
                'hitung nett cogs
                If row2.IsNett_COGSNull = True Then
                    row2.Nett_COGS = 0
                End If
                If row2.IsReturn_COGSNull = True Then
                    row2.Return_COGS = 0
                End If
                row2.Nett_COGS = row2.Cost_Of_Sales - row2.Return_COGS

                TotalNettSales += row2.Nett_Sales
                TotalNettCOGS += row2.Nett_COGS
            Next
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_Sales_PercentageNull = True Then
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett sales percentage
                If row2.Nett_Sales_Percentage > 0 Then
                    row2.Nett_Sales_Percentage = Math.Round((row2.Nett_Sales * 100) / TotalNettSales, 2, MidpointRounding.AwayFromZero)
                Else
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett cogs percentage
                If row2.IsNett_COGS_PercentageNull = True Then
                    row2.Nett_COGS_Percentage = 0
                End If
                If row2.Nett_Sales <> 0 Then
                    row2.Nett_COGS_Percentage = Math.Round(((row2.Nett_Sales - row2.Nett_COGS) * 100) / row2.Nett_Sales, 2, MidpointRounding.AwayFromZero)
                    If row2.Nett_Sales < 0 And row2.Nett_COGS < 0 Then
                        row2.Nett_COGS_Percentage = row2.Nett_COGS_Percentage * (-1)
                    End If
                Else
                    row2.Nett_COGS_Percentage = 0
                End If
                row2.IsColor = Me.GetIsColor(row2.Nett_COGS_Percentage, row2.Nett_Sales)
            Next

            'urutkan sesuai nama brand
            TblDummy.Rows.Clear()
            For Each row2 In TBLGroupByBrand.Rows
                rowDummy = TblDummy.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ijk) = row2.Item(ijk)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            TBLGroupByBrand.Rows.Clear()
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "Product Brand"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    row2.Item(ijk) = rowDummy.Item(ijk)
                Next
                TBLGroupByBrand.Rows.Add(row2)
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand_Return.SetDataSource(DS)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY PRODUCT CATEGORY W/ RETURN")
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Group_Name.ParameterFieldName, "PRODUCT CATEGORY")

            Me.crv1.ReportSource = ReportByBrand_Return
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
        End Try
    End Sub

    Private Sub GenerateByProductCodeWithReturn()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView

        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        'Dim TipeStockInSR As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim TotalNettSales, TotalNettCOGS As Decimal
        Dim TblDummy As New DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable
        Dim rowDummy As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim DV As DataRowView

        Try
            TotalNettCOGS = 0
            TotalNettSales = 0

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount

                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Product_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = row1.Cost
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = row1.Product_Code
                            .Product_Brand_Name = row1.Description
                            .Retur_Sales = 0
                            .Return_COGS = 0
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = SalesAmount
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                End With
            Next

            'isikan data SR
            DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.EndPeriod, Me.StartPeriode)
            Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, Me.StartPeriode, Me.EndPeriod)
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To TBLSR_Raw.DefaultView.Count - 1
                rowSR_Raw = TBLSR_Raw.DefaultView(i).Row
                If Me.ApakahMasukKriteria(rowSR_Raw) = True Then
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(rowSR_Raw.Product_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = 0
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = rowSR_Raw.Product_Code
                            .Product_Brand_Name = rowSR_Raw.Description
                            .Retur_Sales = rowSR_Raw.Amount
                            .Return_COGS = rowSR_Raw.Cost
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = 0
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Retur_Sales += rowSR_Raw.Amount
                        row2.Return_COGS += rowSR_Raw.Cost
                    End If
                End If
            Next

            'hitung data yang belum lengkap
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_SalesNull = True Then
                    row2.Nett_Sales = 0
                End If
                If row2.IsRetur_SalesNull Then
                    row2.Retur_Sales = 0
                End If
                'hitung nett sales
                row2.Nett_Sales = row2.Sales_Amount - row2.Retur_Sales
                'hitung nett cogs
                If row2.IsNett_COGSNull = True Then
                    row2.Nett_COGS = 0
                End If
                If row2.IsReturn_COGSNull = True Then
                    row2.Return_COGS = 0
                End If
                row2.Nett_COGS = row2.Cost_Of_Sales - row2.Return_COGS

                TotalNettSales += row2.Nett_Sales
                TotalNettCOGS += row2.Nett_COGS
            Next
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_Sales_PercentageNull = True Then
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett sales percentage

                If row2.Nett_Sales_Percentage > 0 Then
                    row2.Nett_Sales_Percentage = Math.Round((row2.Nett_Sales * 100) / TotalNettSales, 2, MidpointRounding.AwayFromZero)
                Else
                    row2.Nett_Sales_Percentage = 0
                End If

                'hitung nett cogs percentage
                If row2.IsNett_COGS_PercentageNull = True Then
                    row2.Nett_COGS_Percentage = 0
                End If
                If row2.Nett_Sales <> 0 Then
                    row2.Nett_COGS_Percentage = Math.Round(((row2.Nett_Sales - row2.Nett_COGS) * 100) / row2.Nett_Sales, 2, MidpointRounding.AwayFromZero)
                    If row2.Nett_Sales < 0 And row2.Nett_COGS < 0 Then
                        row2.Nett_COGS_Percentage = row2.Nett_COGS_Percentage * (-1)
                    End If
                Else
                    row2.Nett_COGS_Percentage = 0
                End If
                row2.IsColor = Me.GetIsColor(row2.Nett_COGS_Percentage, row2.Nett_Sales)
            Next

            'urutkan sesuai nama brand
            TblDummy.Rows.Clear()
            For Each row2 In TBLGroupByBrand.Rows
                rowDummy = TblDummy.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ijk) = row2.Item(ijk)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            TBLGroupByBrand.Rows.Clear()
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "Product Brand"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    row2.Item(ijk) = rowDummy.Item(ijk)
                Next
                TBLGroupByBrand.Rows.Add(row2)
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand_Return.SetDataSource(DS)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY PRODUCT CODE W/ RETURN")
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Group_Name.ParameterFieldName, "PRODUCT CODE")

            Me.crv1.ReportSource = ReportByBrand_Return
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
        End Try
    End Sub

    Private Sub GenerateByBrandWithReturn()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView

        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        'Dim TipeStockInSR As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim TotalNettSales, TotalNettCOGS As Decimal
        Dim TblDummy As New DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable
        Dim rowDummy As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim DV As DataRowView

        Try
            TotalNettCOGS = 0
            TotalNettSales = 0

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To TBL.DefaultView.Count - 1
                row1 = TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount

                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Brand) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = row1.Cost
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = row1.Brand
                            .Product_Brand_Name = row1.Brand_Description
                            .Retur_Sales = 0
                            .Return_COGS = 0
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = SalesAmount
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                End With
            Next

            'isikan data SR
            DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.EndPeriod, Me.StartPeriode)
            Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, Me.StartPeriode, Me.EndPeriod)
            For Each rowSR_Raw In TBLSR_Raw.Rows
                If Me.ApakahMasukKriteria(rowSR_Raw) = True Then
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(rowSR_Raw.Brand_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = 0
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = rowSR_Raw.Brand_Code
                            .Product_Brand_Name = rowSR_Raw.Brand_Desc
                            .Retur_Sales = rowSR_Raw.Amount
                            .Return_COGS = rowSR_Raw.Cost
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = 0
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Retur_Sales += rowSR_Raw.Amount
                        row2.Return_COGS += rowSR_Raw.Cost
                    End If
                End If
            Next

            'hitung data yang belum lengkap
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_SalesNull = True Then
                    row2.Nett_Sales = 0
                End If
                If row2.IsRetur_SalesNull Then
                    row2.Retur_Sales = 0
                End If
                'hitung nett sales
                row2.Nett_Sales = row2.Sales_Amount - row2.Retur_Sales
                'hitung nett cogs
                If row2.IsNett_COGSNull = True Then
                    row2.Nett_COGS = 0
                End If
                If row2.IsReturn_COGSNull = True Then
                    row2.Return_COGS = 0
                End If
                row2.Nett_COGS = row2.Cost_Of_Sales - row2.Return_COGS

                TotalNettSales += row2.Nett_Sales
                TotalNettCOGS += row2.Nett_COGS
            Next
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_Sales_PercentageNull = True Then
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett sales percentage
                row2.Nett_Sales_Percentage = Math.Round((row2.Nett_Sales * 100) / TotalNettSales, 2, MidpointRounding.AwayFromZero)
                'hitung nett cogs percentage
                If row2.IsNett_COGS_PercentageNull = True Then
                    row2.Nett_COGS_Percentage = 0
                End If
                If row2.Nett_Sales <> 0 Then
                    row2.Nett_COGS_Percentage = Math.Round(((row2.Nett_Sales - row2.Nett_COGS) * 100) / row2.Nett_Sales, 2, MidpointRounding.AwayFromZero)
                    If row2.Nett_Sales < 0 And row2.Nett_COGS < 0 Then
                        row2.Nett_COGS_Percentage = row2.Nett_COGS_Percentage * (-1)
                    End If
                Else
                    row2.Nett_COGS_Percentage = 0
                End If
                row2.IsColor = Me.GetIsColor(row2.Nett_COGS_Percentage, row2.Nett_Sales)
            Next

            'urutkan sesuai nama brand
            TblDummy.Rows.Clear()
            For Each row2 In TBLGroupByBrand.Rows
                rowDummy = TblDummy.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ijk) = row2.Item(ijk)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            TBLGroupByBrand.Rows.Clear()
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "Product Brand"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    row2.Item(ijk) = rowDummy.Item(ijk)
                Next
                TBLGroupByBrand.Rows.Add(row2)
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand_Return.SetDataSource(DS)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY PRODUCT BRAND W/ RETURN")
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Group_Name.ParameterFieldName, "PRODUCT BRAND")

            Me.crv1.ReportSource = ReportByBrand_Return
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
        End Try
    End Sub

    Private Sub GenerateBySalesAreaWithReturn()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView

        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        'Dim TipeStockInSR As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Dim TotalNettSales, TotalNettCOGS As Decimal
        Dim TblDummy As New DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable
        Dim rowDummy As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim DV As DataRowView

        Try
            TotalNettCOGS = 0
            TotalNettSales = 0

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount

                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Sales_Area_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = row1.Cost
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = row1.Sales_Area_Code
                            .Product_Brand_Name = row1.Sales_Area_Description
                            .Retur_Sales = 0
                            .Return_COGS = 0
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = SalesAmount
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                End With
            Next

            'isikan data SR
            DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.EndPeriod, Me.StartPeriode)
            Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, Me.StartPeriode, Me.EndPeriod)
            TBLSR_Raw.DefaultView.RowFilter = Me.FilterByDgvItem2
            For i As Integer = 0 To TBLSR_Raw.DefaultView.Count - 1
                rowSR_Raw = TBLSR_Raw.DefaultView(i).Row
                If Me.ApakahMasukKriteria(rowSR_Raw) = True Then
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(rowSR_Raw.Sales_Area_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = 0
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = rowSR_Raw.Sales_Area_Code
                            .Product_Brand_Name = rowSR_Raw.Sales_Area_Desc
                            .Retur_Sales = rowSR_Raw.Amount
                            .Return_COGS = rowSR_Raw.Cost
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = 0
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Retur_Sales += rowSR_Raw.Amount
                        row2.Return_COGS += rowSR_Raw.Cost
                    End If
                End If
            Next

            'hitung data yang belum lengkap
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_SalesNull = True Then
                    row2.Nett_Sales = 0
                End If
                If row2.IsRetur_SalesNull Then
                    row2.Retur_Sales = 0
                End If
                'hitung nett sales
                row2.Nett_Sales = row2.Sales_Amount - row2.Retur_Sales
                'hitung nett cogs
                If row2.IsNett_COGSNull = True Then
                    row2.Nett_COGS = 0
                End If
                If row2.IsReturn_COGSNull = True Then
                    row2.Return_COGS = 0
                End If
                row2.Nett_COGS = row2.Cost_Of_Sales - row2.Return_COGS

                TotalNettSales += row2.Nett_Sales
                TotalNettCOGS += row2.Nett_COGS
            Next
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_Sales_PercentageNull = True Then
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett sales percentage

                If row2.Nett_Sales_Percentage > 0 Then
                    row2.Nett_Sales_Percentage = Math.Round((row2.Nett_Sales * 100) / TotalNettSales, 2, MidpointRounding.AwayFromZero)
                Else
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett cogs percentage
                If row2.IsNett_COGS_PercentageNull = True Then
                    row2.Nett_COGS_Percentage = 0
                End If
                If row2.Nett_Sales <> 0 Then
                    row2.Nett_COGS_Percentage = Math.Round(((row2.Nett_Sales - row2.Nett_COGS) * 100) / row2.Nett_Sales, 2, MidpointRounding.AwayFromZero)
                    If row2.Nett_Sales < 0 And row2.Nett_COGS < 0 Then
                        row2.Nett_COGS_Percentage = row2.Nett_COGS_Percentage * (-1)
                    End If
                Else
                    row2.Nett_COGS_Percentage = 0
                End If
                row2.IsColor = Me.GetIsColor(row2.Nett_COGS_Percentage, row2.Nett_Sales)
            Next

            'urutkan sesuai nama brand
            TblDummy.Rows.Clear()
            For Each row2 In TBLGroupByBrand.Rows
                rowDummy = TblDummy.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ijk) = row2.Item(ijk)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            TBLGroupByBrand.Rows.Clear()
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "Product Brand"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    row2.Item(ijk) = rowDummy.Item(ijk)
                Next
                TBLGroupByBrand.Rows.Add(row2)
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand_Return.SetDataSource(DS)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY SALES AREA W/ RETURN")
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Group_Name.ParameterFieldName, "SALES AREA")

            Me.crv1.ReportSource = ReportByBrand_Return
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
        End Try
    End Sub

    Private Sub GenerateByProductCode()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView
        Dim dummy As Double = 0
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Product_Code) + "'"


                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupByBrand.Rows.Add(row1.Product_Code, row1.Description, 0, 0, 0, _
                        0, SalesAmount, 0, row1.Cost, 0)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    .Group_Name = row2.Product_Brand & " - " & row2.Product_Brand_Name
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                    .IsColor = Me.GetIsColor(.GM_Percentage, .Sales_Amount)
                End With
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand.SetDataSource(DS)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY PRODUCT CODE")
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Group_Name.ParameterFieldName, "PRODUCT CODE")

            Me.crv1.ReportSource = ReportByBrand
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

    Private Sub ReOrderByInvoiceNumber(ByVal TBL As DseLaporanSalesAnalysis.DataTable1DataTable)
        Dim TBLDummy As New DseLaporanSalesAnalysis.DataTable1DataTable
        Dim row As DseLaporanSalesAnalysis.DataTable1Row
        Dim Sort As String = "[No Transaksi]"
        Dim DummySort As String = TBL.DefaultView.Sort
        Dim rowView As DataRowView

        Try
            TBL.DefaultView.Sort = Sort
            For Each rowView In TBL.DefaultView
                TBLDummy.ImportRow(rowView.Row)
            Next
            TBL.Rows.Clear()
            For Each row In TBLDummy.Rows
                TBL.ImportRow(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLDummy = Nothing
        End Try
    End Sub

    'Private Sub GenerateByInvoiceNumber()
    '    Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
    '    Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
    '    Dim SalesAmount, TotalSales, TotalSalesYTD As Double
    '    Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
    '    Dim cmd As String = ""
    '    Dim TBL2View As DataView = Me.TBL2.DefaultView
    '    Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView
    '    Dim dummy As Double = 0
    '    Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
    '    Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
    '    Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
    '    Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
    '    Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
    '    Dim DummyDate As DateTime = #1/1/1950#

    '    Try
    '        DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
    '        If Me.chkIncludeSalesReturn.Checked = True Then
    '            DummyDate = Me.StartPeriodYTD
    '            While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
    '                DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
    '                Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
    '                For Each rowSR_Raw In TBLSR_Raw.Rows
    '                    row1 = TBL.NewDataTable1Row
    '                    With row1
    '                        .Amount = rowSR_Raw.Amount * -1
    '                        .Brand = rowSR_Raw.Brand_Code
    '                        .Brand_Description = rowSR_Raw.Brand_Desc
    '                        .Calculates_By = 0
    '                        .Category_Code = rowSR_Raw.Category_Code
    '                        .Category_Desc = rowSR_Raw.Category_Desc
    '                        .Cost = rowSR_Raw.Cost * -1
    '                        .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
    '                        .Customer_Code = rowSR_Raw.Customer_Code
    '                        .Customer_Name = rowSR_Raw.Customer_Name
    '                        .Default_Precision = rowSR_Raw.Default_Precision
    '                        .Description = rowSR_Raw.Description
    '                        .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
    '                        .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
    '                        .Primary_Currency = rowSR_Raw.Primary_Currency
    '                        .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
    '                        .Primary_Invoice = rowSR_Raw.Primary_SR
    '                        .Primary_Product = rowSR_Raw.Primary_Product
    '                        .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
    '                        .Product_Code = rowSR_Raw.Product_Code
    '                        .Sales_Amount = Math.Round(rowSR_Raw.Total_Qty * rowSR_Raw.Nett_Price, rowSR_Raw.Default_Precision, MidpointRounding.AwayFromZero) * -1
    '                        .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
    '                        .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
    '                        .Segmentation_Code = rowSR_Raw.Segmentation_Code
    '                        .Segmentation_Description = rowSR_Raw.Segementasi_Desc
    '                        .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
    '                        .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
    '                        .Type_Code = rowSR_Raw.Cust_Type_Code
    '                        .Primary_SO = rowSR_Raw.Primary_SO
    '                        .Nett_Price = rowSR_Raw.Nett_Price
    '                        .No_Transaksi = rowSR_Raw.No_Transaksi
    '                        .Price_Exclude = 1
    '                        .Primary_Brand = rowSR_Raw.Primary_Brand
    '                        .Primary_Category = rowSR_Raw.Primary_Category
    '                        .Total_Qty = rowSR_Raw.Total_Qty
    '                    End With
    '                    TBL.Rows.Add(row1)
    '                Next
    '                DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

    '                'MsgBox("")
    '            End While
    '        End If

    '        'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
    '        Me.ReOrderByInvoiceNumber(Me.TBL)
    '        For Each row1 In Me.TBL.Rows
    '            If Me.ApakahMasukKriteria(row1) = True Then
    '                SalesAmount = row1.Sales_Amount
    '                'masukkan data total sales ytd
    '                TotalSalesYTD += SalesAmount
    '                'masukkan hasilnya ke total sales ytd + baris
    '                cmd = "[Product Brand] ='" + row1.No_Transaksi + "'"


    '                TBLGroupByBrandView.RowFilter = cmd
    '                If TBLGroupByBrandView.Count = 0 Then
    '                    'tambahkan baris baru
    '                    TBLGroupByBrand.Rows.Add(row1.No_Transaksi, row1.Description, 0, 0, 0, _
    '                    0, SalesAmount, 0, row1.Cost, 0)
    '                    'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
    '                    If row1.Invoice_Date >= Me.StartPeriode Then
    '                        row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
    '                        row2.Sales_Amount += SalesAmount
    '                        row2.Cost_Of_Sales += row1.Cost
    '                        'masukkan data total sales
    '                        TotalSales += SalesAmount
    '                    End If
    '                Else
    '                    row2 = TBLGroupByBrandView.Item(0).Row
    '                    'tambahan sales amount ytd
    '                    row2.Sales_Amount_YTD += SalesAmount
    '                    row2.Cost_Of_Sales_YTD += row1.Cost
    '                    'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
    '                    If row1.Invoice_Date >= Me.StartPeriode Then
    '                        row2.Sales_Amount += SalesAmount
    '                        row2.Cost_Of_Sales += row1.Cost
    '                        'masukkan data total sales
    '                        TotalSales += SalesAmount
    '                    End If
    '                End If
    '            End If
    '        Next

    '        For Each row2 In Me.TBLGroupByBrand.Rows
    '            With row2
    '                'hitung % total sales
    '                If TotalSales <> 0 Then
    '                    .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
    '                Else
    '                    .Total_Sales_Percentage = 0
    '                End If
    '                'hitung gm percentage
    '                If .Sales_Amount = 0 Then
    '                    .GM_Percentage = 0
    '                Else
    '                    .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
    '                End If
    '                'hitung % total sales YTD
    '                If TotalSalesYTD <> 0 Then
    '                    .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
    '                Else
    '                    .Total_Sales_YTD_Percentage = 0
    '                End If
    '                'hitung GM% YTD
    '                If .Sales_Amount_YTD = 0 Then
    '                    .GM_Percentage_YTD = 0
    '                Else
    '                    .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
    '                    .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
    '                End If
    '            End With
    '        Next

    '        'generate report
    '        Dim CustomerCode As String = ""
    '        Dim CustomerType As String = ""
    '        Dim Segmentation As String = ""
    '        Dim SalesArea As String = ""
    '        Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
    '        CustomerType, Segmentation, SalesArea)
    '        ReportByBrand.SetDataSource(DS)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerType.ParameterFieldName, CustomerType)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_SalesArea.ParameterFieldName, SalesArea)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Segmentation.ParameterFieldName, Segmentation)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY INVOICE NUMBER")
    '        ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Group_Name.ParameterFieldName, "INVOICE NUMBER")

    '        Me.crv1.ReportSource = ReportByBrand
    '        Me.crv1.DisplayGroupTree = False
    '        Me.crv1.DisplayStatusBar = True
    '        Me.crv1.DisplayToolbar = True
    '        Me.crv1.ShowGroupTreeButton = False
    '        Me.crv1.ShowExportButton = True
    '        Me.crv1.ShowRefreshButton = False
    '        Me.crv1.Show()

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub GenerateByCategory()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Category_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupByBrand.Rows.Add(row1.Category_Code, row1.Category_Desc, 0, 0, 0, _
                        0, SalesAmount, 0, row1.Cost, 0)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    .Group_Name = row2.Product_Brand & " - " & row2.Product_Brand_Name

                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                    .IsColor = Me.GetIsColor(row2.GM_Percentage, .Sales_Amount)
                End With
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand.SetDataSource(DS)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY PRODUCT CATEGORY")
            ReportByBrand.SetParameterValue(ReportByBrand.Parameter_Group_Name.ParameterFieldName, "PRODUCT CATEGORY")

            Me.crv1.ReportSource = ReportByBrand
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

    Function GetIsColor(ByVal vGrossMarginPercent As Double, ByVal vSalesAmount As Decimal) As Integer
        Try
            If Me.ChkMarkGM.Checked = True Then
                If vSalesAmount <> 0 Then
                    If Me.CboFilterGM.SelectedIndex = 0 Then
                        If vGrossMarginPercent < Me.txtGMPercentage.Text Then
                            Return 1
                        End If
                    ElseIf Me.CboFilterGM.SelectedIndex = 1 Then
                        If vGrossMarginPercent <= Me.txtGMPercentage.Text Then
                            Return 1
                        End If
                    ElseIf Me.CboFilterGM.SelectedIndex = 2 Then
                        If vGrossMarginPercent = Me.txtGMPercentage.Text Then
                            Return 1
                        End If
                    ElseIf Me.CboFilterGM.SelectedIndex = 3 Then
                        If vGrossMarginPercent >= Me.txtGMPercentage.Text Then
                            Return 1
                        End If
                    ElseIf Me.CboFilterGM.SelectedIndex = 4 Then
                        If vGrossMarginPercent > Me.txtGMPercentage.Text Then
                            Return 1
                        End If
                    End If
                End If
            End If

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahBolehCetak() As Boolean
        Try
            If Me.ObjClosing.GetPrimaryClosing(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpSalesInvoicePeriod.Value)) = -1 Then
                MsgBox("Please do monthly closing for period: " + Format(Me.dtpSalesInvoicePeriod.Value.Date, "MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'If Me.clbBrand.CheckedItems.Count <= 0 Then
            '    MsgBox("Please at least 1 product brand.", MsgBoxStyle.Information, "Information")
            '    Me.clbBrand.Focus()
            '    Return False
            'End If
            'If Me.clbCustomer.CheckedItems.Count <= 0 Then
            '    MsgBox("Please at least 1 customer.", MsgBoxStyle.Information, "Information")
            '    Me.clbCustomer.Focus()
            '    Return False
            'End If
            'If Me.clbCustomerType.CheckedItems.Count <= 0 Then
            '    MsgBox("Please at least 1 customer type.", MsgBoxStyle.Information, "Information")
            '    Me.clbCustomerType.Focus()
            '    Return False
            'End If
            'If Me.clbMarketSegmentation.CheckedItems.Count <= 0 Then
            '    MsgBox("Please at least 1 market segmentation.", MsgBoxStyle.Information, "Information")
            '    Me.clbMarketSegmentation.Focus()
            '    Return False
            'End If
            'If Me.clbSalesArea.CheckedItems.Count <= 0 Then
            '    MsgBox("Please at least 1 sales area.", MsgBoxStyle.Information, "Information")
            '    Me.clbSalesArea.Focus()
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

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
            'If Me.chkCustomer.Checked = True Then
            '    CustomerCode = "All"
            'Else
            '    Dim DRV As DataRowView
            '    For Each DRV In Me.clbCustomer.CheckedItems
            '        If CustomerCode = "" Then
            '            CustomerCode = DRV.Item(1)
            '        Else
            '            CustomerCode = CustomerCode + ", " + DRV.Item(1)
            '        End If
            '    Next
            'End If

            'If Me.chkCustType.Checked = True Then
            '    CustomerType = "All"
            'Else
            '    Dim DRV As DataRowView
            '    For Each DRV In Me.clbCustomerType.CheckedItems
            '        If CustomerType = "" Then
            '            CustomerType = DRV.Item(1)
            '        Else
            '            CustomerType = CustomerType + ", " + DRV.Item(1)
            '        End If
            '    Next
            'End If

            'If Me.chkMarketSegmentation.Checked = True Then
            '    Segmentation = "All"
            'Else
            '    Dim DRV As DataRowView
            '    For Each DRV In Me.clbMarketSegmentation.CheckedItems
            '        If Segmentation = "" Then
            '            Segmentation = DRV.Item(1)
            '        Else
            '            Segmentation = Segmentation + ", " + DRV.Item(1)
            '        End If
            '    Next
            'End If

            'If Me.chkSalesArea.Checked = True Then
            '    SalesArea = "All"
            'Else
            '    Dim DRV As DataRowView
            '    For Each DRV In Me.clbSalesArea.CheckedItems
            '        If SalesArea = "" Then
            '            SalesArea = DRV.Item(1)
            '        Else
            '            SalesArea = SalesArea + ", " + DRV.Item(1)
            '        End If
            '    Next
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetCostOfSales(ByVal KriteriaPerhitungan As EnumKriteriaPerhitunganCostOfSales, _
    ByVal StartDate As Date, ByVal EndDate As Date, _
    ByVal Kriteria1 As String, Optional ByVal Kriteria2 As String = "") As Double
        Dim cmd As String = ""
        Dim cmd2 As String = ""
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.DataTable3Row
        Dim TBLView As DataView = Me.TBL.DefaultView
        Dim TBLView2 As DataView = Me.TBL3.DefaultView
        Dim TotalCost As Double = 0
        Dim TanggalClosingPertama As Date = Me.ObjPurchaseClosing.GetPurchaseClosingMinimumDate

        Try
            'filter data penjualan, sesuai dengan kriteria yang diberikan
            Select Case KriteriaPerhitungan
                Case EnumKriteriaPerhitunganCostOfSales.ByBrand
                    cmd = "[Brand] ='" + Kriteria1 + "'"
                    Exit Select
                Case EnumKriteriaPerhitunganCostOfSales.ByCategory
                    cmd = "[Category Code] ='" + Kriteria1 + "'"
                    Exit Select
                Case EnumKriteriaPerhitunganCostOfSales.ByCustType
                    cmd = "[Type Code] ='" + Kriteria1 + "'"
                    Exit Select
                Case EnumKriteriaPerhitunganCostOfSales.ByMarketSegmentation
                    cmd = "[Segmentation Code] ='" + Kriteria1 + "'"
                    Exit Select
                Case EnumKriteriaPerhitunganCostOfSales.BySalesAreaDetail
                    cmd = "[Sales Area Code] ='" + Kriteria1 + "' and Brand = '" + Kriteria2 + "'"
                    Exit Select
                Case EnumKriteriaPerhitunganCostOfSales.BySalesAreaRekap
                    cmd = "[Sales Area Code] ='" + Kriteria1 + "'"
                    Exit Select
                Case EnumKriteriaPerhitunganCostOfSales.BySalesAreaDetailCategory
                    cmd = "[Sales Area Code] ='" + Kriteria1 + "' and [Category Code] = '" + Kriteria2 + "'"
                    Exit Select
            End Select
            TBLView.RowFilter = cmd

            'hitung total cost
            For Each RowView In TBLView
                row1 = RowView.Row
                'total cost, yang dihitung sebagai cost adalah harga inventory pada bulan terjadinya stock out 
                'selain itu cost pada closing pertama tidak dianggap
                If row1.Invoice_Date >= StartDate And Me.ObjCD.GetFirstDateofTheMonth(row1.Tanggal_Stock_Out) _
                > TanggalClosingPertama Then
                    If row1.IsTanggal_Stock_OutNull = False Then
                        cmd2 = "[Primary Product] =" + CStr(row1.Primary_Product) + " and " + _
                        "[Closing Period] =#" + Me.ObjCD.GetFirstDateofTheMonth(row1.Tanggal_Stock_Out.ToShortDateString) _
                        + "#"
                        TBLView2.RowFilter = cmd2
                        If TBLView2.Count <> 0 Then
                            row2 = TBLView2.Item(0).Row
                            If row2.Ending_Qty <> 0 Then
                                TotalCost += Math.Round(row2.Ending_Amount / row2.Ending_Qty, 0, MidpointRounding.AwayFromZero) _
                                * row1.Total_Qty
                            ElseIf row2.Out_Qty <> 0 Then
                                TotalCost += Math.Round(row2.Out_Amount / row2.Out_Qty, 0, MidpointRounding.AwayFromZero) * _
                                row1.Total_Qty
                            End If
                        End If
                    End If
                End If
            Next

            Return TotalCost
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GenerateByMarketSegmentationWithReturn()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView

        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        'Dim TipeStockInSR As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Dim TotalNettSales, TotalNettCOGS As Decimal
        Dim TblDummy As New DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable
        Dim rowDummy As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim DV As DataRowView

        Try
            TotalNettCOGS = 0
            TotalNettSales = 0

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount

                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Segmentation_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = row1.Cost
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = row1.Segmentation_Code
                            .Product_Brand_Name = row1.Segmentation_Description
                            .Retur_Sales = 0
                            .Return_COGS = 0
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = SalesAmount
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                End With
            Next

            'isikan data SR
            DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.EndPeriod, Me.StartPeriode)
            Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, Me.StartPeriode, Me.EndPeriod)
            TBLSR_Raw.DefaultView.RowFilter = Me.FilterByDgvItem2
            For i As Integer = 0 To TBLSR_Raw.DefaultView.Count - 1
                rowSR_Raw = TBLSR_Raw.DefaultView(i).Row
                If Me.ApakahMasukKriteria(rowSR_Raw) = True Then
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(rowSR_Raw.Segmentation_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = 0
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = rowSR_Raw.Segmentation_Code
                            .Product_Brand_Name = rowSR_Raw.Segementasi_Desc
                            .Retur_Sales = rowSR_Raw.Amount
                            .Return_COGS = rowSR_Raw.Cost
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = 0
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Retur_Sales += rowSR_Raw.Amount
                        row2.Return_COGS += rowSR_Raw.Cost
                    End If
                End If
            Next

            'hitung data yang belum lengkap
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_SalesNull = True Then
                    row2.Nett_Sales = 0
                End If
                If row2.IsRetur_SalesNull Then
                    row2.Retur_Sales = 0
                End If
                'hitung nett sales
                row2.Nett_Sales = row2.Sales_Amount - row2.Retur_Sales
                'hitung nett cogs
                If row2.IsNett_COGSNull = True Then
                    row2.Nett_COGS = 0
                End If
                If row2.IsReturn_COGSNull = True Then
                    row2.Return_COGS = 0
                End If
                row2.Nett_COGS = row2.Cost_Of_Sales - row2.Return_COGS

                TotalNettSales += row2.Nett_Sales
                TotalNettCOGS += row2.Nett_COGS
            Next
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_Sales_PercentageNull = True Then
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett sales percentage
                If TotalNettSales > 0 Then
                    row2.Nett_Sales_Percentage = Math.Round((row2.Nett_Sales * 100) / TotalNettSales, 2, MidpointRounding.AwayFromZero)
                End If

                'hitung nett cogs percentage
                If row2.IsNett_COGS_PercentageNull = True Then
                    row2.Nett_COGS_Percentage = 0
                End If
                If row2.Nett_Sales <> 0 Then
                    row2.Nett_COGS_Percentage = Math.Round(((row2.Nett_Sales - row2.Nett_COGS) * 100) / row2.Nett_Sales, 2, MidpointRounding.AwayFromZero)
                    If row2.Nett_Sales < 0 And row2.Nett_COGS < 0 Then
                        row2.Nett_COGS_Percentage = row2.Nett_COGS_Percentage * (-1)
                    End If
                Else
                    row2.Nett_COGS_Percentage = 0
                End If
                row2.IsColor = Me.GetIsColor(row2.Nett_COGS_Percentage, row2.Nett_Sales)
            Next

            'urutkan sesuai nama brand
            TblDummy.Rows.Clear()
            For Each row2 In TBLGroupByBrand.Rows
                rowDummy = TblDummy.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ijk) = row2.Item(ijk)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            TBLGroupByBrand.Rows.Clear()
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "Product Brand"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    row2.Item(ijk) = rowDummy.Item(ijk)
                Next
                TBLGroupByBrand.Rows.Add(row2)
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand_Return.SetDataSource(DS)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY MARKET SEGMENTATION W/ RETURN")
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Group_Name.ParameterFieldName, "MARKET SEGMENTATION")

            Me.crv1.ReportSource = ReportByBrand_Return
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
        End Try
    End Sub

    Private Sub GenerateByMarketSegmentation()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Customer_Market_SegmentationRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByMarketSegmentationView As DataView = Me.TBLGroupBySegmentation.DefaultView
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
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
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Customer Market Segmentation] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Segmentation_Code) + "'"
                    TBLGroupByMarketSegmentationView.RowFilter = cmd
                    If TBLGroupByMarketSegmentationView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupBySegmentation.Rows.Add(row1.Segmentation_Code, row1.Segmentation_Description, _
                        0, 0, 0, _
                        0, SalesAmount, 0, row1.Cost, 0)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupBySegmentation.Item(TBLGroupBySegmentation.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByMarketSegmentationView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next
            For Each row2 In Me.TBLGroupBySegmentation.Rows
                With row2
                    .Group_Name = row2.Customer_Market_Segmentation & " - " & row2.Segmentation_Desc
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung cost of sales
                    '.Cost_Of_Sales = Me.GetCostOfSales(EnumKriteriaPerhitunganCostOfSales.ByMarketSegmentation, _
                    'Me.StartPeriode, Me.EndPeriod, .Customer_Market_Segmentation)
                    'hitung GM %
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) _
                        / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung cost of sales YTD
                    '.Cost_Of_Sales_YTD = Me.GetCostOfSales(EnumKriteriaPerhitunganCostOfSales.ByMarketSegmentation, _
                    'Me.StartPeriodYTD, Me.EndPeriod, .Customer_Market_Segmentation)
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                    .IsColor = Me.GetIsColor(.GM_Percentage, .Sales_Amount)
                End With
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByMarketSegmentation.SetDataSource(DS)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByMarketSegmentation.SetParameterValue(ReportByMarketSegmentation.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            Me.crv1.ReportSource = ReportByMarketSegmentation
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

    Private Sub GenerateByCustType()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Customer_TypeRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByCustTypeView As DataView = Me.TBLGroupByCustType.DefaultView
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Customer Type] = '" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Type_Code) + " - " + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Cust_Type_Description) + "'"
                    TBLGroupByCustTypeView.RowFilter = cmd
                    If TBLGroupByCustTypeView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupByCustType.Rows.Add(row1.Type_Code + " - " + row1.Cust_Type_Description, _
                        0, 0, 0, 0, SalesAmount, 0, row1.Cost, 0, row1.Type_Code)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByCustType.Item(TBLGroupByCustType.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByCustTypeView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next
            For Each row2 In Me.TBLGroupByCustType.Rows
                With row2
                    .Group_Name = row2.Customer_Type
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                    .IsColor = Me.GetIsColor(.GM_Percentage, .Sales_Amount)
                End With
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByCustType.SetDataSource(DS)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByCustType.SetParameterValue(ReportByCustType.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            Me.crv1.ReportSource = ReportByCustType
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

    Private Sub GenerateByCustTypeWithReturn()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView

        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        'Dim TipeStockInSR As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Dim TotalNettSales, TotalNettCOGS As Decimal
        Dim TblDummy As New DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable
        Dim rowDummy As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim DV As DataRowView

        Try
            TotalNettCOGS = 0
            TotalNettSales = 0

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            Me.TBL.DefaultView.RowFilter = Me.FilterByDgvItem
            For i As Integer = 0 To Me.TBL.DefaultView.Count - 1
                row1 = Me.TBL.DefaultView(i).Row
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount

                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Type_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = row1.Cost
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = row1.Type_Code
                            .Product_Brand_Name = row1.Cust_Type_Description
                            .Retur_Sales = 0
                            .Return_COGS = 0
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = SalesAmount
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                End With
            Next

            'isikan data SR
            DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.EndPeriod, Me.StartPeriode)
            Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, Me.StartPeriode, Me.EndPeriod)
            TBLSR_Raw.DefaultView.RowFilter = Me.FilterByDgvItem2
            For i As Integer = 0 To TBLSR_Raw.DefaultView.Count - 1
                rowSR_Raw = TBLSR_Raw.DefaultView(i).Row
                If Me.ApakahMasukKriteria(rowSR_Raw) = True Then
                    cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(rowSR_Raw.Cust_Type_Code) + "'"
                    TBLGroupByBrandView.RowFilter = cmd
                    If TBLGroupByBrandView.Count = 0 Then
                        'tambahkan baris baru
                        row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                        With row2
                            .Cost_Of_Sales = 0
                            .Cost_Of_Sales_YTD = 0
                            .GM_Percentage = 0
                            .GM_Percentage_YTD = 0
                            .Nett_COGS = 0
                            .Nett_COGS_Percentage = 0
                            .Nett_Sales = 0
                            .Nett_Sales_Percentage = 0
                            .Product_Brand = rowSR_Raw.Cust_Type_Code
                            .Product_Brand_Name = rowSR_Raw.Cust_Type_Desc
                            .Retur_Sales = rowSR_Raw.Amount
                            .Return_COGS = rowSR_Raw.Cost
                            .Sales_Amount = 0
                            .Sales_Amount_YTD = 0
                            .Total_Sales_Percentage = 0
                            .Total_Sales_YTD_Percentage = 0
                        End With
                        TBLGroupByBrand.Rows.Add(row2)
                    Else
                        row2 = TBLGroupByBrandView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Retur_Sales += rowSR_Raw.Amount
                        row2.Return_COGS += rowSR_Raw.Cost
                    End If
                End If
            Next

            'hitung data yang belum lengkap
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_SalesNull = True Then
                    row2.Nett_Sales = 0
                End If
                If row2.IsRetur_SalesNull Then
                    row2.Retur_Sales = 0
                End If
                'hitung nett sales
                row2.Nett_Sales = row2.Sales_Amount - row2.Retur_Sales
                'hitung nett cogs
                If row2.IsNett_COGSNull = True Then
                    row2.Nett_COGS = 0
                End If
                If row2.IsReturn_COGSNull = True Then
                    row2.Return_COGS = 0
                End If
                row2.Nett_COGS = row2.Cost_Of_Sales - row2.Return_COGS

                TotalNettSales += row2.Nett_Sales
                TotalNettCOGS += row2.Nett_COGS
            Next
            For Each row2 In Me.TBLGroupByBrand.Rows
                If row2.IsNett_Sales_PercentageNull = True Then
                    row2.Nett_Sales_Percentage = 0
                End If
                'hitung nett sales percentage
                If row2.Nett_Sales_Percentage > 0 Then
                    row2.Nett_Sales_Percentage = Math.Round((row2.Nett_Sales * 100) / TotalNettSales, 2, MidpointRounding.AwayFromZero)
                Else
                    row2.Nett_Sales_Percentage = 0
                End If

                'hitung nett cogs percentage
                If row2.IsNett_COGS_PercentageNull = True Then
                    row2.Nett_COGS_Percentage = 0
                End If
                If row2.Nett_Sales <> 0 Then
                    row2.Nett_COGS_Percentage = Math.Round(((row2.Nett_Sales - row2.Nett_COGS) * 100) / row2.Nett_Sales, 2, MidpointRounding.AwayFromZero)
                    If row2.Nett_Sales < 0 And row2.Nett_COGS < 0 Then
                        row2.Nett_COGS_Percentage = row2.Nett_COGS_Percentage * (-1)
                    End If
                Else
                    row2.Nett_COGS_Percentage = 0
                End If
                row2.IsColor = Me.GetIsColor(row2.Nett_COGS_Percentage, row2.Nett_Sales)
            Next

            'urutkan sesuai nama brand
            TblDummy.Rows.Clear()
            For Each row2 In TBLGroupByBrand.Rows
                rowDummy = TblDummy.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ijk) = row2.Item(ijk)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            TBLGroupByBrand.Rows.Clear()
            TblDummy.DefaultView.RowFilter = ""
            TblDummy.DefaultView.Sort = "Product Brand"
            For Each DV In TblDummy.DefaultView
                rowDummy = DV.Row
                row2 = TBLGroupByBrand.NewGroup_By_Product_BrandRow
                For ijk As Integer = 0 To TblDummy.Columns.Count - 1
                    row2.Item(ijk) = rowDummy.Item(ijk)
                Next
                TBLGroupByBrand.Rows.Add(row2)
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportByBrand_Return.SetDataSource(DS)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_ReportName.ParameterFieldName, "SALES ANALYSIS BY CUSTOMER TYPE W/ RETURN")
            ReportByBrand_Return.SetParameterValue(ReportByBrand_Return.Parameter_Group_Name.ParameterFieldName, "CUSTOMER TYPE")

            Me.crv1.ReportSource = ReportByBrand_Return
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR_Raw.Dispose()
            TBLSR_Raw.Dispose()
        End Try
    End Sub

    Private Sub GenerateBySalesAreaDetail()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis._Group_By_Market_Area_DetailRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupBySalesAreaDetailView As DataView = Me.TBLGroupBySalesAreaDetail.DefaultView
        Dim TBLDummy As New DseLaporanSalesAnalysis._Group_By_Market_Area_DetailDataTable
        Dim Sort As String = "[Market Area], [Product Brand]"
        Dim rowView As DataRowView
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            For Each row1 In Me.TBL.Rows
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Market Area] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Sales_Area_Code) + "' and [Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Brand) + "'"
                    TBLGroupBySalesAreaDetailView.RowFilter = cmd
                    If TBLGroupBySalesAreaDetailView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupBySalesAreaDetail.Rows.Add(row1.Sales_Area_Code, row1.Sales_Area_Description, _
                        row1.Brand, row1.Brand_Description, 0, 0, 0, 0, SalesAmount, 0, row1.Cost, 0)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupBySalesAreaDetail.Item(TBLGroupBySalesAreaDetail.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupBySalesAreaDetailView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            'urutan ulang
            TBLGroupBySalesAreaDetailView.RowFilter = ""
            TBLGroupBySalesAreaDetailView.Sort = Sort
            For Each rowView In TBLGroupBySalesAreaDetailView
                row2 = rowView.Row
                With row2
                    TBLDummy.Rows.Add(.Market_Area, .Market_Area_Desc, .Product_Brand, .Product_Brand_Name, _
                    .Sales_Amount, .Total_Sales_Percentage, .Cost_Of_Sales, _
                    .GM_Percentage, .Sales_Amount_YTD, .Total_Sales_YTD_Percentage, .Cost_Of_Sales_YTD, _
                    .GM_Percentage_YTD)
                End With
            Next
            Me.TBLGroupBySalesAreaDetail.Rows.Clear()
            For Each row2 In TBLDummy.Rows
                With row2
                    Me.TBLGroupBySalesAreaDetail.Rows.Add(.Market_Area, .Market_Area_Desc, .Product_Brand, .Product_Brand_Name, _
                    .Sales_Amount, .Total_Sales_Percentage, .Cost_Of_Sales, _
                    .GM_Percentage, .Sales_Amount_YTD, .Total_Sales_YTD_Percentage, .Cost_Of_Sales_YTD, _
                    .GM_Percentage_YTD)
                End With
            Next

            For Each row2 In Me.TBLGroupBySalesAreaDetail.Rows
                With row2
                    .Group_Name = row2.Product_Brand & " - " & row2.Product_Brand_Name
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                    .IsColor = Me.GetIsColor(.GM_Percentage, .Sales_Amount)
                End With
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportBySalesAreaDetail.SetDataSource(DS)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            Me.crv1.ReportSource = ReportBySalesAreaDetail
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

    Private Sub GenerateBySalesAreaDetail2()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis._Group_By_Market_Area_DetailRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupBySalesAreaDetailView As DataView = Me.TBLGroupBySalesAreaDetail.DefaultView
        Dim TBLDummy As New DseLaporanSalesAnalysis._Group_By_Market_Area_DetailDataTable
        Dim Sort As String = "[Market Area], [Product Brand]"
        Dim rowView As DataRowView
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            For Each row1 In Me.TBL.Rows
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Market Area] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Sales_Area_Code) + "' and [Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Category_Code) + "'"
                    TBLGroupBySalesAreaDetailView.RowFilter = cmd
                    If TBLGroupBySalesAreaDetailView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupBySalesAreaDetail.Rows.Add(row1.Sales_Area_Code, row1.Sales_Area_Description, _
                        row1.Category_Code, row1.Category_Desc, 0, 0, 0, 0, SalesAmount, 0, row1.Cost, 0)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupBySalesAreaDetail.Item(TBLGroupBySalesAreaDetail.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupBySalesAreaDetailView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            'urutan ulang
            TBLGroupBySalesAreaDetailView.RowFilter = ""
            TBLGroupBySalesAreaDetailView.Sort = Sort
            For Each rowView In TBLGroupBySalesAreaDetailView
                row2 = rowView.Row
                With row2
                    TBLDummy.Rows.Add(.Market_Area, .Market_Area_Desc, .Product_Brand, .Product_Brand_Name, _
                    .Sales_Amount, .Total_Sales_Percentage, .Cost_Of_Sales, _
                    .GM_Percentage, .Sales_Amount_YTD, .Total_Sales_YTD_Percentage, .Cost_Of_Sales_YTD, _
                    .GM_Percentage_YTD)
                End With
            Next
            Me.TBLGroupBySalesAreaDetail.Rows.Clear()
            For Each row2 In TBLDummy.Rows
                With row2
                    Me.TBLGroupBySalesAreaDetail.Rows.Add(.Market_Area, .Market_Area_Desc, .Product_Brand, .Product_Brand_Name, _
                    .Sales_Amount, .Total_Sales_Percentage, .Cost_Of_Sales, _
                    .GM_Percentage, .Sales_Amount_YTD, .Total_Sales_YTD_Percentage, .Cost_Of_Sales_YTD, _
                    .GM_Percentage_YTD)
                End With
            Next

            For Each row2 In Me.TBLGroupBySalesAreaDetail.Rows
                With row2
                    .Group_Name = .Product_Brand & " - " & .Product_Brand_Name
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                    .IsColor = Me.GetIsColor(.GM_Percentage, .Sales_Amount)
                End With
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportBySalesAreaDetail.SetDataSource(DS)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportBySalesAreaDetail.SetParameterValue(ReportBySalesAreaDetail.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            Me.crv1.ReportSource = ReportBySalesAreaDetail
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

    Private Sub GenerateBySalesAreaRekap()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis._Group_By_Market_Area_RekapRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupBySalesAreaRekapView As DataView = Me.TBLGroupBySalesAreaRekap.DefaultView
        Dim DTASR_Raw As New DseLaporanSalesAnalysisTableAdapters.DataTable5TableAdapter
        Dim TBLSR_Raw As New DseLaporanSalesAnalysis.DataTable5DataTable
        Dim rowSR_Raw As DseLaporanSalesAnalysis.DataTable5Row
        Dim TipeSR As ClsTransaksiSalesReturn.EnumTipeSR = ClsTransaksiSalesReturn.EnumTipeSR.Nota_Retur_Pajak
        Dim TipeStockInSR As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
        Dim DummyDate As DateTime = #1/1/1950#

        Try
            DTASR_Raw.Connection.ConnectionString = Core_Module.GetConstring
            If Me.chkIncludeSalesReturn.Checked = True Then
                DummyDate = Me.StartPeriodYTD
                While Me.ObjCD.GetLasDateOfTheMonth(DummyDate) <= Me.EndPeriod
                    DTASR_Raw.Fill(TBLSR_Raw, TipeStockInSR, TipeSR, Me.ObjCD.GetLasDateOfTheMonth(DummyDate), DummyDate)
                    Me.IsikanReturnCost_AmountNonIDR(TBLSR_Raw, DummyDate, Me.ObjCD.GetLasDateOfTheMonth(DummyDate))
                    For Each rowSR_Raw In TBLSR_Raw.Rows
                        row1 = TBL.NewDataTable1Row
                        With row1
                            .Amount = rowSR_Raw.Amount * -1
                            .Brand = rowSR_Raw.Brand_Code
                            .Brand_Description = rowSR_Raw.Brand_Desc
                            .Calculates_By = 0
                            .Category_Code = rowSR_Raw.Category_Code
                            .Category_Desc = rowSR_Raw.Category_Desc
                            .Cost = rowSR_Raw.Cost * -1
                            .Cust_Type_Description = rowSR_Raw.Cust_Type_Desc
                            .Customer_Code = rowSR_Raw.Customer_Code
                            .Customer_Name = rowSR_Raw.Customer_Name
                            .Default_Precision = rowSR_Raw.Default_Precision
                            .Description = rowSR_Raw.Description
                            .Exchange_Rate_To_IDR = rowSR_Raw.Exchange_Rate
                            .Invoice_Date = rowSR_Raw.Tanggal_Transaksi
                            .Primary_Currency = rowSR_Raw.Primary_Currency
                            .Primary_DO = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Primary_Invoice = rowSR_Raw.Primary_SR
                            .Primary_Product = rowSR_Raw.Primary_Product
                            .Primary_Stock_Out_Sales = rowSR_Raw.Primary_Stock_In_Sales_Return
                            .Product_Code = rowSR_Raw.Product_Code
                            .Sales_Amount = rowSR_Raw.Amount * -1
                            .Sales_Area_Code = rowSR_Raw.Sales_Area_Code
                            .Sales_Area_Description = rowSR_Raw.Sales_Area_Desc
                            .Segmentation_Code = rowSR_Raw.Segmentation_Code
                            .Segmentation_Description = rowSR_Raw.Segementasi_Desc
                            .Tanggal_Stock_Out = rowSR_Raw.Tanggal_Transaksi
                            .Tax_Exchange_Rate = rowSR_Raw.Exchange_Rate
                            .Type_Code = rowSR_Raw.Cust_Type_Code
                            .Primary_SO = rowSR_Raw.Primary_SO
                            .Nett_Price = rowSR_Raw.Nett_Price
                            .No_Transaksi = rowSR_Raw.No_Transaksi
                            .Price_Exclude = 1
                            .Primary_Brand = rowSR_Raw.Primary_Brand
                            .Primary_Category = rowSR_Raw.Primary_Category
                            .Total_Qty = rowSR_Raw.Total_Qty
                        End With
                        TBL.Rows.Add(row1)
                    Next
                    DummyDate = Me.ObjCD.GetFirstDateofTheMonth(DummyDate + TimeSpan.FromDays(32))

                    'MsgBox("")
                End While
            End If

            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            For Each row1 In Me.TBL.Rows
                If Me.ApakahMasukKriteria(row1) = True Then
                    'hitung sales amount berdasarkan rate tsb
                    SalesAmount = row1.Sales_Amount
                    'masukkan data total sales ytd
                    TotalSalesYTD += SalesAmount
                    'masukkan hasilnya ke total sales ytd + baris
                    cmd = "[Market Area] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Sales_Area_Code) + "'"
                    TBLGroupBySalesAreaRekapView.RowFilter = cmd
                    If TBLGroupBySalesAreaRekapView.Count = 0 Then
                        'tambahkan baris baru
                        TBLGroupBySalesAreaRekap.Rows.Add(row1.Sales_Area_Code, row1.Sales_Area_Description, _
                        0, 0, 0, 0, SalesAmount, 0, row1.Cost, 0)
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2 = TBLGroupBySalesAreaRekap.Item(TBLGroupBySalesAreaRekap.Rows.Count - 1)
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    Else
                        row2 = TBLGroupBySalesAreaRekapView.Item(0).Row
                        'tambahan sales amount ytd
                        row2.Sales_Amount_YTD += SalesAmount
                        row2.Cost_Of_Sales_YTD += row1.Cost
                        'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                        If row1.Invoice_Date >= Me.StartPeriode Then
                            row2.Sales_Amount += SalesAmount
                            row2.Cost_Of_Sales += row1.Cost
                            'masukkan data total sales
                            TotalSales += SalesAmount
                        End If
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupBySalesAreaRekap.Rows
                With row2
                    .Group_Name = .Market_Area & " - " & .Market_Area_Desc
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                    .IsColor = Me.GetIsColor(.GM_Percentage, .Sales_Amount)
                End With
            Next

            'generate report
            Dim CustomerCode As String = ""
            Dim CustomerType As String = ""
            Dim Segmentation As String = ""
            Dim SalesArea As String = ""
            Me.IsikanCustomerCodeCustomerTypeSegmentationSalesArea(CustomerCode, _
            CustomerType, Segmentation, SalesArea)
            ReportBySalesAreaRekap.SetDataSource(DS)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_CustomerCode.ParameterFieldName, CustomerCode)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_CustomerType.ParameterFieldName, CustomerType)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_SalesArea.ParameterFieldName, SalesArea)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_Segmentation.ParameterFieldName, Segmentation)
            ReportBySalesAreaRekap.SetParameterValue(ReportBySalesAreaRekap.Parameter_Tgl.ParameterFieldName, Me.dtpSalesInvoicePeriod.Value)
            Me.crv1.ReportSource = ReportBySalesAreaRekap
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

    'Private Sub chkCustType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Me.chkCustType.Checked = True Then
    '            Me.clbCustomerType.Enabled = False
    '            For i As Integer = 0 To Me.clbCustomerType.Items.Count - 1
    '                Me.clbCustomerType.SetItemCheckState(i, CheckState.Checked)
    '            Next
    '        Else
    '            For i As Integer = 0 To Me.clbCustomerType.Items.Count - 1
    '                Me.clbCustomerType.SetItemCheckState(i, CheckState.Unchecked)
    '            Next
    '            Me.clbCustomerType.Enabled = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub chkCustomer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Me.chkCustomer.Checked = True Then
    '            Me.clbCustomer.Enabled = False
    '            For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
    '                Me.clbCustomer.SetItemCheckState(i, CheckState.Checked)
    '            Next
    '        Else
    '            For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
    '                Me.clbCustomer.SetItemCheckState(i, CheckState.Unchecked)
    '            Next
    '            Me.clbCustomer.Enabled = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub chkBrand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Me.chkBrand.Checked = True Then
    '            Me.clbBrand.Enabled = False
    '            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
    '                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
    '            Next
    '        Else
    '            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
    '                Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
    '            Next
    '            Me.clbBrand.Enabled = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub chkSalesArea_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Me.chkSalesArea.Checked = True Then
    '            Me.clbSalesArea.Enabled = False
    '            For i As Integer = 0 To Me.clbSalesArea.Items.Count - 1
    '                Me.clbSalesArea.SetItemCheckState(i, CheckState.Checked)
    '            Next
    '        Else
    '            For i As Integer = 0 To Me.clbSalesArea.Items.Count - 1
    '                Me.clbSalesArea.SetItemCheckState(i, CheckState.Unchecked)
    '            Next
    '            Me.clbSalesArea.Enabled = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub chkMarketSegmentation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If Me.chkMarketSegmentation.Checked = True Then
    '            Me.clbMarketSegmentation.Enabled = False
    '            For i As Integer = 0 To Me.clbMarketSegmentation.Items.Count - 1
    '                Me.clbMarketSegmentation.SetItemCheckState(i, CheckState.Checked)
    '            Next
    '        Else
    '            For i As Integer = 0 To Me.clbMarketSegmentation.Items.Count - 1
    '                Me.clbMarketSegmentation.SetItemCheckState(i, CheckState.Unchecked)
    '            Next
    '            Me.clbMarketSegmentation.Enabled = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Sub CekOut(ByVal vStartPeriod As Date, ByVal vEndPeriod As Date)
        Dim DTA As New DseLaporanSalesAnalysisTableAdapters.DataTable1TableAdapter
        Dim DTA2 As New DseLaporanSalesAnalysisTableAdapters.DataTable2TableAdapter
        Dim DTA3 As New DseLaporanSalesAnalysisTableAdapters.DataTable3TableAdapter

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.KosongkanTabel()
            'set awal dan akhir periode perhitungan
            Me.StartPeriode = vStartPeriod
            Me.EndPeriod = vEndPeriod
            Me.StartPeriodYTD = CDate("1/1/" + CStr(vStartPeriod.Year))

            'isikan tabel yang akan digunakan untuk perhitungan
            DTA.Fill_TanggalSampaiDari(TBL, Me.EndPeriod, Me.StartPeriodYTD)
            DTA2.Fill(TBL2, Me.EndPeriod, Me.StartPeriodYTD)
            DTA3.FillBy_TanggalSampaiDari(Me.TBL3, Me.EndPeriod, Me.StartPeriodYTD)
            Me.DTAReturn.Fill(Me.TBLReturn, Me.EndPeriod, Me.StartPeriodYTD)
            Me.KurangkanItemReturn()

            'set nilai maksimum progress bar
            Me.pbr1.Value = 0
            Me.pbr1.Maximum = TBL.Rows.Count

            Me.GenerateByProductCode_CekOut()

            Me.TabControl1.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Sub

    Private Sub GenerateByProductCode_CekOut()
        Dim row1 As DseLaporanSalesAnalysis.DataTable1Row
        Dim row2 As DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim SalesAmount, TotalSales, TotalSalesYTD As Double
        Dim CalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order
        Dim cmd As String = ""
        Dim TBL2View As DataView = Me.TBL2.DefaultView
        Dim TBLGroupByBrandView As DataView = Me.TBLGroupByBrand.DefaultView

        Try
            'untuk setiap baris pada tabel satu, pindahkan ke tabel GroupByBrand
            For Each row1 In Me.TBL.Rows
                SalesAmount = row1.Sales_Amount
                'masukkan data total sales ytd
                TotalSalesYTD += SalesAmount
                'masukkan hasilnya ke total sales ytd + baris
                cmd = "[Product Brand] ='" + Me.ObjCD.FormatSpecialChar_DataFilter(row1.Product_Code) + "'"
                TBLGroupByBrandView.RowFilter = cmd
                If TBLGroupByBrandView.Count = 0 Then
                    'tambahkan baris baru
                    TBLGroupByBrand.Rows.Add(row1.Product_Code, row1.Description, 0, 0, 0, _
                    0, SalesAmount, 0, row1.Cost, 0)
                    'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                    If row1.Invoice_Date >= Me.StartPeriode Then
                        row2 = TBLGroupByBrand.Item(TBLGroupByBrand.Rows.Count - 1)
                        row2.Sales_Amount += SalesAmount
                        row2.Cost_Of_Sales += row1.Cost
                        'masukkan data total sales
                        TotalSales += SalesAmount
                    End If
                Else
                    row2 = TBLGroupByBrandView.Item(0).Row
                    'tambahan sales amount ytd
                    row2.Sales_Amount_YTD += SalesAmount
                    row2.Cost_Of_Sales_YTD += row1.Cost
                    'jika termasuk pada periode, masukkan hasilnya ke total sales + baris
                    If row1.Invoice_Date >= Me.StartPeriode Then
                        row2.Sales_Amount += SalesAmount
                        row2.Cost_Of_Sales += row1.Cost
                        'masukkan data total sales
                        TotalSales += SalesAmount
                    End If
                End If
            Next

            For Each row2 In Me.TBLGroupByBrand.Rows
                With row2
                    'hitung % total sales
                    If TotalSales <> 0 Then
                        .Total_Sales_Percentage = Math.Round(.Sales_Amount / TotalSales * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_Percentage = 0
                    End If
                    'hitung gm percentage
                    If .Sales_Amount = 0 Then
                        .GM_Percentage = 0
                    Else
                        .GM_Percentage = Math.Round((.Sales_Amount - .Cost_Of_Sales) / .Sales_Amount * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount < 0 And .Cost_Of_Sales < 0 Then
                            .GM_Percentage = .GM_Percentage * (-1)
                        End If
                    End If
                    'hitung % total sales YTD
                    If TotalSalesYTD <> 0 Then
                        .Total_Sales_YTD_Percentage = Math.Round(.Sales_Amount_YTD / TotalSalesYTD * 100, 2, MidpointRounding.AwayFromZero)
                    Else
                        .Total_Sales_YTD_Percentage = 0
                    End If
                    'hitung GM% YTD
                    If .Sales_Amount_YTD = 0 Then
                        .GM_Percentage_YTD = 0
                    Else
                        .GM_Percentage_YTD = Math.Round((.Sales_Amount_YTD - .Cost_Of_Sales_YTD) / _
                        .Sales_Amount_YTD * 100, 2, MidpointRounding.AwayFromZero)
                        If .Sales_Amount_YTD < 0 And .Cost_Of_Sales_YTD < 0 Then
                            .GM_Percentage_YTD = .GM_Percentage_YTD * (-1)
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboSalesAreaRekap_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboSalesAreaRekap.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboSalesAreaDetailCategory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboSalesAreaDetailCategory.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboSalesAreaDetail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboSalesAreaDetail.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboBySalesAreaWithReturn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboBySalesAreaWithReturn.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SettingRbo()
        Try
            If Me.rboBySalesAreaWithReturn.Checked = True Or Me.rboSalesAreaDetail.Checked = True Or _
               Me.rboSalesAreaDetailCategory.Checked = True Or Me.rboSalesAreaRekap.Checked = True Then
                If Me.cboDeliveryType.Visible = False Then
                    Me.cboDeliveryType.SelectedIndex = 0
                End If
                Me.cboDeliveryType.Visible = True
                Me.cboDeliveryType.TabStop = True
            Else
                Me.cboDeliveryType.Visible = False
                Me.cboDeliveryType.TabStop = False
            End If

            If Me.rboByMarketSegmentationWithReturn.Checked = True Or _
            Me.rboByBrandWithReturn.Checked = True Or _
            Me.rboByCustomerTypeWithReturn.Checked = True Or _
            Me.rboByProductCategorywithReturn.Checked = True Or _
            Me.rboBySalesAreaWithReturn.Checked = True Or _
            Me.rboProductCodeWithReturn.Checked = True Then
                Me.chkIncludeSalesReturn.Checked = True
                Me.chkIncludeSalesReturn.Enabled = False
                'Me.ChkMarkGM.Checked = False
                'Me.ChkMarkGM.Enabled = False
            Else
                Me.chkIncludeSalesReturn.Enabled = True
                'Me.ChkMarkGM.Checked = False
                'Me.ChkMarkGM.Enabled = True
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByProductCategorywithReturn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByProductCategorywithReturn.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByMarketSegmentationWithReturn_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboByMarketSegmentationWithReturn.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByBrand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByBrand.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByProductCategory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByProductCategory.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByMarketSegmentation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByMarketSegmentation.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByCustomerType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByCustomerType.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByCustomerTypeWithReturn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByCustomerTypeWithReturn.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByBrandWithReturn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByBrandWithReturn.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboProductCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboProductCode.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboProductCodeWithReturn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboProductCodeWithReturn.CheckedChanged
        Try
            Me.SettingRbo()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkMarkGM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkMarkGM.CheckedChanged
        Try
            If Me.ChkMarkGM.Checked = True Then
                Me.CboFilterGM.Enabled = True
                Me.txtGMPercentage.Enabled = True
                Me.txtGMPercentage.Text = "0"
            Else
                Me.CboFilterGM.Enabled = False
                Me.txtGMPercentage.Enabled = False
                Me.txtGMPercentage.Text = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGMPercentage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGMPercentage.Click
        Try
            Me.ObjCD.Number_Click(Me.txtGMPercentage)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGMPercentage_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGMPercentage.Enter
        Try
            Me.ObjCD.Number_Enter(Me.txtGMPercentage, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGMPercentage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGMPercentage.GotFocus
        Me.txtGMPercentage.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtGMPercentage_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGMPercentage.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtGMPercentage, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGMPercentage_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtGMPercentage.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtGMPercentage, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGMPercentage_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGMPercentage.Leave
        Try
            Me.ObjCD.Number_Leave(Me.txtGMPercentage, sender, e, 2, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtGMPercentage_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGMPercentage.LostFocus
        Me.txtGMPercentage.BackColor = Color.White
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

    Private Function FilterByDgvItem2() As String
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
                            FilterBy_CustomerType = "[Cust Type Codee]='" & row.Cells(Me.ColCT_CustType.Index).Value & "'"
                        Else
                            FilterBy_CustomerType &= " OR [Cust Type Code]='" & row.Cells(Me.ColCT_CustType.Index).Value & "'"
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
                            FilterBy_Brand = "[Brand Code]='" & row.Cells(Me.ColB_Brand.Index).Value & "'"
                        Else
                            FilterBy_Brand &= " OR [Brand Code]='" & row.Cells(Me.ColB_Brand.Index).Value & "'"
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