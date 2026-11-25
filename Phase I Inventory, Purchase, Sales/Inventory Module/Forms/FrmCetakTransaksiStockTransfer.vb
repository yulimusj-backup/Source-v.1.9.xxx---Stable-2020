Option Explicit On
Option Strict Off
Public Class FrmCetakTransaksiStockTransfer
    Public primary As Double = -1
    Public FormatLaporan As EnumFormatLaporan = EnumFormatLaporan.A4
    Public TipeLaporan As FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockBundling
    Public CustomReport As Boolean = False
    Public UserCode As String = ""

    Public Enum EnumFormatLaporan As Integer
        A4
        A5_Landscape
    End Enum

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub FrmCetakTransaksiStockTransfer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockBundling Then
                Me.Text = "Print Stock Bundling"
                If Me.CetakStockBundling(Me.crv1, Me.primary) = False Then
                    Me.Close()
                End If
            ElseIf Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInAdjust Or _
                    Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutAdjust Then
                Me.Text = "Print Stock Adjust"
                If Me.CustomReport = False Then
                    If Me.CetakStockInOutAdjust(Me.crv1, Me.primary) = False Then
                        Me.Close()
                    End If
                End If
            ElseIf Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInPurchase Or _
                    Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutPurchaseReturn Then
                If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInPurchase Then
                    Me.Text = "Print Stock In Purchase"
                Else
                    Me.Text = "Print Stock Out Purchase Return"
                End If
                If Me.CustomReport = False Then
                    If Me.CetakStockInOutPurchase(Me.crv1, Me.primary) = False Then
                        Me.Close()
                    End If
                End If
            ElseIf Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInSalesReturn Or _
        Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutSales Then
                Me.Text = "Print Stock Sales"
                If Me.CustomReport = False Then
                    If Me.CetakStockInOutSales(Me.crv1, Me.primary) = False Then
                        Me.Close()
                    End If
                End If
            ElseIf Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockTransfer Then

                Me.Text = "Print Stock Transfer"
                If Me.CustomReport = False Then
                    If Me.CetakStockTransfer(Me.crv1, Me.primary) = False Then
                        Me.Close()
                    End If
                End If
                
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function CetakStockBundling(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objCustom As New CustomSMESS.ClsTransaksiStockBundling
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim Report As New CrpCetakStockBundling
        Dim Ds As New dsLaporanStockBundling
        Dim DtaBef As New dsLaporanStockBundlingTableAdapters.DataTable1TableAdapter
        Dim DtaAft As New dsLaporanStockBundlingTableAdapters.DataTable2TableAdapter
        Dim TblBef As dsLaporanStockBundling.DataTable1DataTable = Ds.DataTable1
        Dim TblAft As dsLaporanStockBundling.DataTable2DataTable = Ds.DataTable2
        Dim RowBef As dsLaporanStockBundling.DataTable1Row
        Dim RowAft As dsLaporanStockBundling.DataTable2Row
        Dim RowCetak As dsLaporanStockBundling.TblCetakRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                ObjSetting = Nothing
                Return objCustom.CetakStockBundling(vPrimary, CRV, Me.UserCode)
            End If

            DtaBef.FillBy_PrimaryTransaksi(TblBef, vPrimary)
            DtaAft.FillBy_PrimaryTransaksi(TblAft, vPrimary)

            For Each RowBef In TblBef.Rows
                RowCetak = Ds.TblCetak.NewRow
                With RowCetak
                    .Primary = RowBef.Primary
                    .No_Transaksi = RowBef.No_Transaksi
                    .Tanggal_Transaksi = RowBef.Tanggal_Transaksi
                    .Remarks = RowBef.Remarks
                    .Created_By = RowBef.Created_By
                    .Product_Code = RowBef.Product_Code
                    .Description = RowBef.Description
                    .Qty = RowBef.Qty
                    .Unit_Code = RowBef.Unit_Code
                    .Secondary_Qty = RowBef.Secondary_Qty
                    .Unit_Code_2 = RowBef.Unit_Code_2
                    .Batch_No = RowBef.Batch_No
                    .Expiry_Date = RowBef.Expiry_Date
                    .Warehouse_Code = RowBef.Warehouse
                    .Tipe = "Before"
                    .RemarksItem = RowBef.RemarksItem
                    .Cost_Factor = -1
                    ._Decimal = RowBef._Decimal
                    If RowBef.Batch_No <> "" Then
                        .Qty_Per_Batch = RowBef.Qty
                    Else
                        .Qty_Per_Batch = 0
                    End If
                End With
                Ds.TblCetak.Rows.Add(RowCetak)
            Next

            For Each RowAft In TblAft.Rows
                RowCetak = Ds.TblCetak.NewRow
                With RowCetak
                    .Primary = RowAft.Primary
                    .No_Transaksi = RowAft.No_Transaksi
                    .Tanggal_Transaksi = RowAft.Tanggal_Transaksi
                    .Remarks = RowAft.Remarks
                    .Created_By = RowAft.Created_By
                    .Product_Code = RowAft.Product_Code
                    .Description = RowAft.Description
                    .Qty = RowAft.Qty
                    .Unit_Code = RowAft.Unit_Code
                    .Secondary_Qty = RowAft.Secondary_Qty
                    .Unit_Code_2 = RowAft.Unit_Code_2
                    .Batch_No = RowAft.Batch_No
                    .Expiry_Date = RowAft.Expiry_Date
                    .Warehouse_Code = RowAft.Warehouse
                    .Tipe = "After"
                    .RemarksItem = RowAft.RemarksItem
                    .Cost_Factor = RowAft.Faktor_Pemberat
                    ._Decimal = RowAft._Decimal
                    If RowAft.Batch_No <> "" Then
                        .Qty_Per_Batch = RowAft.Qty
                    Else
                        .Qty_Per_Batch = 0
                    End If
                End With
                Ds.TblCetak.Rows.Add(RowCetak)
            Next

            If Me.FormatLaporan = EnumFormatLaporan.A4 Then
                Report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
            ElseIf Me.FormatLaporan = EnumFormatLaporan.A5_Landscape Then
                Report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA5
                Report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
            End If

            Report.SetDataSource(Ds)
            Report.SetParameterValue(Report.Parameter_City.ParameterFieldName, ObjCompany.GetCompany_City)

            CRV.ReportSource = Report
            CRV.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaBef.Dispose()
            DtaAft.Dispose()
            TblBef.Dispose()
            TblAft.Dispose()
            Ds.Dispose()
            ObjCompany = Nothing
            objCustom = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Private Function CetakStockInOutAdjust(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objCustom As New CustomSMESS.ClsTransaksiStockInOut
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim ShowDetail As Integer = 0

        Dim report As New CrpCetakStockInOutAdjust
        Dim dta As New dsLaporanStockInAdjustTableAdapters.DataTable1TableAdapter
        Dim TBL As New dsLaporanStockInAdjust.DataTable1DataTable
        Dim row As dsLaporanStockInAdjust.DataTable1Row

        Dim ds As New dsLaporanStockInAdjust
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowlaporan As dsLaporanStockInAdjust.TblCetakRow
        Dim TipeAdjustment As String = ""

        Try
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInAdjust Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Adjust, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Return objCustom.CetakStockInAdjust(CRV, vPrimary)
                End If
            End If
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutAdjust Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Adjust, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Return objCustom.CetakStockOutAdjust(CRV, vPrimary)
                End If
            End If

            dta.FillBy_PrimaryTransaksi(TBL, vPrimary)
            For Each row In TBL.Rows
                rowlaporan = ds.TblCetak.NewRow()
                rowlaporan.Stock_In_No = row.No_Transaksi
                rowlaporan._Date = row.Tanggal_Transaksi
                rowlaporan.Code_No = row.Product_Code
                rowlaporan.Description = row.Description
                rowlaporan.Remark = row.Remarks
                rowlaporan.Remarks1 = row.Remarks1
                rowlaporan.Created = row.Created_By
                rowlaporan.Batch = row.Batch_No
                rowlaporan.Expired_Date = row.Expiry_Date
                rowlaporan.Warehouse = row.WarehouseCode
                rowlaporan.Location = row.StorageAreaCode
                If row.Batch_No <> "" Then
                    rowlaporan.Quantity_Per_Batch = row.Qty
                Else
                    rowlaporan.Quantity_Per_Batch = 0
                End If
                rowlaporan.Quantity = row.Qty
                rowlaporan.Unit_ = row.Unit_Code
                rowlaporan.SecondaryUnit = row.SecUnit
                rowlaporan.SecondaryQuantity = row.Secondary_Qty
                rowlaporan.Stock_In_No = row.No_Transaksi
                rowlaporan._Decimal = row._Decimal
                ds.TblCetak.Rows.Add(rowlaporan)

                If ShowDetail = 0 Then
                    If row.Batch_No <> "" Or row.WarehouseCode <> "" Then
                        ShowDetail = 1
                    End If
                End If
            Next

            If Me.FormatLaporan = EnumFormatLaporan.A4 Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
            ElseIf Me.FormatLaporan = EnumFormatLaporan.A5_Landscape Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA5
                report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
            End If

            report.SetDataSource(ds)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                TipeAdjustment = row.Adjustment_Type
            End If
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInAdjust Then
                If TipeAdjustment = "" Then
                    report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK IN ADJUST")
                Else
                    report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK IN ADJUST - " + TipeAdjustment)
                End If

                Me.Text = "Print Stock In Adjust"
            ElseIf Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutAdjust Then
                If TipeAdjustment = "" Then
                    report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK OUT ADJUST")
                Else
                    report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK OUT ADJUST - " + TipeAdjustment)
                End If
                Me.Text = "Print Stock Out Adjust"
            End If
            report.SetParameterValue(report.Parameter_City.ParameterFieldName, ObjCompany.GetCompany_City)
            report.SetParameterValue(report.Parameter_ShowDetail.ParameterFieldName, ShowDetail)

            CRV.ReportSource = report
            CRV.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            TBL.Dispose()
            ds.Dispose()
            ObjCompany = Nothing
        End Try
    End Function

    Private Function CetakStockInOutPurchase(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double) As Boolean
        Dim ShowDetail As Integer = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objCustom As New CustomSMESS.ClsTransaksiStockInOut
        Dim ObjCompanyData As New Core_Module.ClsMasterCompanyData

        Dim report As New CrpCetakStockInOutPurchase
        Dim dta As New dsLaporanStockInPurchaseTableAdapters.DataTable2TableAdapter
        Dim TBL As New dsLaporanStockInPurchase.DataTable2DataTable
        Dim row As dsLaporanStockInPurchase.DataTable2Row
        Dim ds As New dsLaporanStockInPurchase
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowlaporan As dsLaporanStockInPurchase.TblCetakRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim SupplierDONumber As String = ""

        Try

            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInPurchase Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_Purchase, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Return objCustom.CetakStockInPurchase(CRV, vPrimary)
                End If
            End If
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutPurchaseReturn Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Purchase_Return, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Return objCustom.CetakStockOutPurchaseReturn(CRV, vPrimary)
                End If

                Dim ObjStockOut As New clsTransaksiStockInOut
                Dim ObjPurchaseReturn As New ClsTransaksiPurchaseReturn
                Dim ObjGRN As New ClsTransaksiGRN
                SupplierDONumber = ObjStockOut.GetSupplierDONumber(ObjGRN.GetPrimaryStockInPurchase((ObjPurchaseReturn.GetPrimaryGRN(ObjStockOut.GetPrimaryPurchaseReturn(vPrimary)))))
                ObjStockOut = Nothing
                ObjPurchaseReturn = Nothing
                ObjGRN = Nothing
            End If

            dta.FillBy_primarytransaksi(TBL, vPrimary)

            For Each row In TBL.Rows
                rowlaporan = ds.TblCetak.NewRow()

                If row.IsNo_TransaksiNull Then
                    rowlaporan.No_PO = ""
                Else
                    rowlaporan.No_PO = row.No_Transaksi
                End If
                
                rowlaporan.Product_Code = row.Product_Code
                If row.Primary_Multi_Units = row.Primary_Units Then
                    rowlaporan.Description = row.Description
                Else
                    rowlaporan.Description = row.Description & " (" & ObjCD.digitgrouping(row.Qty_Multi_Unit, row.Decimal2) & row.Unit_Code_Multi & ")"
                End If
                rowlaporan.Batch = row.Batch_No
                rowlaporan.Expired_Date = row.Expiry_Date
                If row.Batch_No <> "" Then
                    rowlaporan.Qty_Per_Batch = row.Qty
                Else
                    rowlaporan.Qty_Per_Batch = 0
                End If
                rowlaporan.Warehouse = row.WarehouseCode
                rowlaporan.Location = row.StorageAreaCode
                rowlaporan._Stock_In_No_ = row.NoTransStockInOut
                rowlaporan._Supplier_DO_No_ = row.Supplier_DO_No
                If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutPurchaseReturn Then
                    rowlaporan._Supplier_DO_No_ = SupplierDONumber
                End If
                rowlaporan.Remark = row.Expr1
                rowlaporan.Supplier_Code = row.Supplier_Code
                rowlaporan.Supplier_Name = row.Supplier_Name
                rowlaporan._Date = row.Tanggal_Transaksi
                rowlaporan.Quantity = row.Qty
                rowlaporan.UnitCode = row.Unit_Code
                rowlaporan.Remarks1 = row.Remarks
                rowlaporan.Created = row.Created_By
                rowlaporan._Decimal = row._Decimal
                rowlaporan.Qty_Multi_Unit = row.Qty_Multi_Unit
                rowlaporan.Unit_Code_Multi = row.Unit_Code_Multi
                rowlaporan.Decimal2 = row.Decimal2
                ds.TblCetak.Rows.Add(rowlaporan)

                If ShowDetail = 0 Then
                    If row.Batch_No <> "" Or row.WarehouseCode <> "" Then
                        ShowDetail = 1
                    End If
                End If
            Next

            If Me.FormatLaporan = EnumFormatLaporan.A4 Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
            ElseIf Me.FormatLaporan = EnumFormatLaporan.A5_Landscape Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA5
                report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
            End If

            report.SetDataSource(ds)
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInPurchase Then
                report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK IN PURCHASE")
                Me.Text = "Print Stock In Purchase"
            ElseIf Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutPurchaseReturn Then
                report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK OUT PURCHASE RETURN")
                Me.Text = "Print Stock Out Purchase Return"
            End If
            report.SetParameterValue(report.Parameter_City.ParameterFieldName, ObjCompanyData.GetCompany_City)
            report.SetParameterValue(report.Parameter_ShowDetail.ParameterFieldName, ShowDetail)

            CRV.ReportSource = report
            CRV.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            TBL.Dispose()
            ds.Dispose()
            ObjCD = Nothing
            ObjCompanyData = Nothing
        End Try
    End Function


    Private Function CetakStockInOutSales(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double) As Boolean
        Dim ShowDetail As Integer = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objCustom As New CustomSMESS.ClsTransaksiStockInOut
        Dim ObjCompanyData As New Core_Module.ClsMasterCompanyData
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjCD As New Core_Module.ClsCekData

        Dim report As New CrpCetakStockInOutSales
        Dim dta As New dsLaporanStockInSalesReturnTableAdapters.DataTable21TableAdapter
        Dim TBL As New dsLaporanStockInSalesReturn.DataTable21DataTable
        Dim row As dsLaporanStockInSalesReturn.DataTable21Row
        Dim ds As New dsLaporanStockInSalesReturn

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowlaporan As dsLaporanStockInSalesReturn.TblCetakRow

        Try

            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInSalesReturn Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_In_SalesReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Return objCustom.CetakStockInSalesReturn(CRV, vPrimary)
                End If
            End If
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutSales Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Out_Sales, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                    ObjSetting = Nothing
                    Return objCustom.CetakStockOutSales(CRV, vPrimary)
                End If
            End If

            dta.FillBy(TBL, vPrimary)
            For Each row In TBL.Rows
                rowlaporan = ds.TblCetak.NewRow()
                rowlaporan.Customer_Code = row.Customer_Code
                rowlaporan.Customer_Name = row.Customer_Name
                rowlaporan.DO_No = row.DoNo
                rowlaporan.Stock_In_No = row.No_Transaksi
                rowlaporan._Date = row.Tanggal_Transaksi
                rowlaporan.Product_Code = row.Product_Code
                rowlaporan.Description = row.Description
                rowlaporan.Remark = row.Remarks
                rowlaporan.Batch = row.Batch_No
                rowlaporan.Expired_Date = row.Expiry_Date
                rowlaporan.Warehouse = row.WarehouseCode
                rowlaporan.Location = row.StorageArea
                If row.Batch_No <> "" Then
                    rowlaporan.Quantity_PerBatch = row.Qty
                Else
                    rowlaporan.Quantity_PerBatch = 0
                End If
                rowlaporan.Quantity = row.Qty
                rowlaporan.Unit_ = row.Unit_Code
                rowlaporan.UnitSecondary = row.SecUnit
                rowlaporan.SecondaryQTY = row.Secondary_Qty
                rowlaporan.Stock_In_No = row.No_Transaksi
                rowlaporan.Created = row.Created_By
                rowlaporan.Remarks = row.Remarks1
                rowlaporan.SO_No = row.SoNo
                rowlaporan._Decimal = row._Decimal

                If row.Primary_Units <> row.Primary_Multi_Units Then
                    rowlaporan.Decimal_Multi = ObjProd.GetAngkaDecimal(row.Primary_Product, row.Primary_Multi_Units)
                    rowlaporan.Qty_Multi = "(" & ObjCD.DigitGrouping(row.Qty_Multi_Unit, rowlaporan.Decimal_Multi).ToString & " " & row.Unit_Code_Multi & ")"
                Else
                    rowlaporan.Decimal_Multi = 0
                    rowlaporan.Qty_Multi = ""
                End If

                ds.TblCetak.Rows.Add(rowlaporan)
                If ShowDetail = 0 Then
                    If row.Batch_No <> "" Or row.WarehouseCode <> "" Then
                        ShowDetail = 1
                    End If
                End If
            Next

            If Me.FormatLaporan = EnumFormatLaporan.A4 Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
            ElseIf Me.FormatLaporan = EnumFormatLaporan.A5_Landscape Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA5
                report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
            End If

            report.SetDataSource(ds)
            If Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockInSalesReturn Then
                report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK IN SALES RETURN")
                Me.Text = "Print Stock In Sales Return"
            ElseIf Me.TipeLaporan = FrmLaporanFormatStockInOut_Bundling.EnumTipeLaporan.StockOutSales Then
                report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "STOCK OUT SALES")
                Me.Text = "Print Stock Out Sales"
            End If
            report.SetParameterValue(report.Parameter_City.ParameterFieldName, ObjCompanyData.GetCompany_City)
            report.SetParameterValue(report.Parameter_ShowDetail.ParameterFieldName, ShowDetail)

            CRV.ReportSource = report
            CRV.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            TBL.Dispose()
            ds.Dispose()
            ObjCompanyData = Nothing
            ObjProd = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function CetakStockTransfer(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objCustom As New CustomSMESS.clsTransaksiStockTransfer
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim report As New CrpCetakStockTransfer
        Dim dta As New dsLaporanStockInAdjustTableAdapters.DataTable2TableAdapter
        Dim TBL As New dsLaporanStockInAdjust.DataTable2DataTable
        Dim row As dsLaporanStockInAdjust.DataTable2Row
        Dim ShowDetail As Integer = 0
        Dim ds As New dsLaporanStockInAdjust
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowlaporan As dsLaporanStockInAdjust.TblCetakRow

        Try
            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Stock_Transfer, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
                ObjSetting = Nothing
                Return objCustom.CetakStockTransfer(CRV, vPrimary)
            End If

            dta.Fill(TBL, vPrimary)
            For Each row In TBL.Rows
                rowlaporan = ds.TblCetak.NewRow()
                rowlaporan.Stock_In_No = row.No_Transaksi
                rowlaporan._Date = row.Tanggal_Transaksi
                rowlaporan.Code_No = row.Product_Code
                rowlaporan.Description = row.Description
                rowlaporan.Remark = row.Remarks
                rowlaporan.Remarks1 = ""
                rowlaporan.Created = row.Created_By
                rowlaporan.Batch = row.Batch_No
                rowlaporan.Expired_Date = row.Expiry_Date
                rowlaporan.Warehouse = row.Warehouse_Code
                rowlaporan.Location = row.Location
                rowlaporan.Warehouse_Target = row.Target_Warehouse
                rowlaporan.Location_Target = row.Target_Warehouse
                If row.Batch_No <> "" Then
                    rowlaporan.Quantity_Per_Batch = row.Qty
                Else
                    rowlaporan.Quantity_Per_Batch = 0
                End If
                rowlaporan.Quantity = row.Qty
                rowlaporan.Unit_ = row.Unit_Code
                rowlaporan.SecondaryUnit = ""
                rowlaporan.SecondaryQuantity = row.Secondary_Qty
                rowlaporan.Stock_In_No = row.No_Transaksi
                rowlaporan._Decimal = row._Decimal
                ds.TblCetak.Rows.Add(rowlaporan)

                If ShowDetail = 0 Then
                    If row.Batch_No <> "" Or row.Warehouse_Code <> "" Then
                        ShowDetail = 1
                    End If
                End If
            Next

            If Me.FormatLaporan = EnumFormatLaporan.A4 Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
            ElseIf Me.FormatLaporan = EnumFormatLaporan.A5_Landscape Then
                report.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA5
                report.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
            End If

            report.SetDataSource(ds)
            Me.Text = "Print Location Transfer"
            report.SetParameterValue(report.Parameter_City.ParameterFieldName, ObjCompany.GetCompany_City)
            report.SetParameterValue(report.Parameter_ShowDetail.ParameterFieldName, ShowDetail)
            report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "Location Transfer")

            CRV.ReportSource = report
            CRV.Show()
            Windows.Forms.Cursor.Current = Cursors.Arrow

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSetting = Nothing
            objCustom = Nothing
            ObjCompany = Nothing
            dta.Dispose()
            TBL.Dispose()
            ds.Dispose()
        End Try
    End Function
End Class