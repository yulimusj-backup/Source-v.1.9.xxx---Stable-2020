Option Explicit On
Option Strict Off

Public Class FrmLaporanInventoryMovementReports
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim RefVal As Decimal = 0

    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDecimal As Integer = 0
    Dim Goods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

    Public UserCode As String = ""
    Dim ObjLicense As New Core_Module.ClsLicense

    Public Enum EnumFormat As Integer
        General
        Detail
    End Enum

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Function ApakahBolehCetak() As Boolean
        Dim PrimaryClosing As Double = -1

        Try
            If Me.clbBrand.CheckedItems.Count <= 0 Then
                MsgBox("Please select at least 1 Brand.", MsgBoxStyle.Information, "Information")
                Me.clbBrand.Focus()
                Return False
            End If

            PrimaryClosing = Me.GetPrimaryPurchaseClosing(Me.dtpFrom.Value)
            If PrimaryClosing = -1 Then
                MsgBox("Please do inventory first for: " + Format(Me.dtpFrom.Value.Date, "MMM-yyyy") + ".", MsgBoxStyle.Information, "Information")
                Me.crvReport.ReportSource = Nothing
                Me.crvReport.Show()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.ApakahBolehCetak = True Then
                If Me.chkShowTypeDetail.Checked = False Then
                    If Me.cboTipe.SelectedIndex = 2 Or Me.cboTipe.SelectedIndex = 3 Then
                        Me.Cetak()
                    Else
                        If Me.cboTipe.SelectedIndex = 0 Then
                            Me.CetakRekap()
                        Else
                            Me.CetakRekap_ByAccount()
                        End If
                    End If
                Else
                    Me.Show_TransaksiDetail2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub CetakRekap_ByAccount()
        Dim DS As New DseLaporanInventoryMovementReport
        'Dim DTA As New DseLaporanInventoryMovementReportTableAdapters.laporan_inventory_movement_purchase_view2TableAdapter
        '     Dim row As DseLaporanInternalPurchaseOrder.internal_purchase_order_report_viewRow
        Dim report As New CrpLaporanInventoryMovementRekaps
        Dim report2 As New CrpLaporanInventoryMovementRekaps1
        Dim JmlBegining, JmlIn, JmlOut, JmlEnding As Decimal

        Dim counter As Double = 0
        Dim row As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2Row
        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim DV As DataRowView
        Dim primaryBrand As Double = 0
        Dim Brand As String = ""

        Dim DTA_CL As New DseLaporanInventoryMovementReportTableAdapters.tbltrans_purchase_closing_item_detail1TableAdapter
        Dim TBL_CL As New DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1DataTable

        Try
            DTA_CL.Connection.ConnectionString = Core_Module.GetConstring
            DS.laporan_inventory_movement_purchase_view2.Rows.Clear()

            'isikan data
            If Me.clbBrand.Items.Count = Me.clbBrand.CheckedItems.Count Then
                'begining balance
                DTA_CL.FillBy_Account(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem)
                Me.IsikanBeginingAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                'in amount
                DTA_CL.FillBy_Account(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem)
                Me.IsikanInAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                'out amount
                DTA_CL.FillBy_Account(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem)
                Me.IsikanOutAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                    'begining balance
                    DTA_CL.FillBy_PrimaryBrang_GroupByAccount(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem)
                    Me.IsikanBeginingAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                    'in amount
                    DTA_CL.FillBy_PrimaryBrang_GroupByAccount(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem)
                    Me.IsikanInAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                    'out amount
                    DTA_CL.FillBy_PrimaryBrang_GroupByAccount(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem)
                    Me.IsikanOutAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                Next
            End If

            'hitung ending balance
            For Each rowItem As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2Row In _
            DS.laporan_inventory_movement_purchase_view2.Rows
                rowItem.Ending_Amount = rowItem.Begin_Amount + rowItem.In_Amount - rowItem.Out_Amount
            Next

            If Me.clbBrand.Items.Count = Me.clbBrand.CheckedItems.Count Then
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

            'filter laporan yang kosong semua tidak usah ditampilkan
            counter = 0
            JmlBegining = 0
            JmlIn = 0
            JmlOut = 0
            JmlEnding = 0
            While counter < DS.laporan_inventory_movement_purchase_view2.Rows.Count
                row = DS.laporan_inventory_movement_purchase_view2.Rows(counter)
                If row.Begin_Amount = 0 And row.Ending_Amount = 0 And _
                row.In_Amount = 0 And row.Out_Amount = 0 Then
                    DS.laporan_inventory_movement_purchase_view2.Rows.RemoveAt(counter)
                Else
                    JmlBegining += row.Begin_Amount
                    JmlIn += row.In_Amount
                    JmlOut += row.Out_Amount
                    JmlEnding += row.Ending_Amount

                    counter += 1
                End If
            End While

            If JmlBegining > Me.RefVal Or JmlIn > Me.RefVal Or JmlOut > Me.RefVal Or JmlEnding > Me.RefVal Then
                report2.SetDataSource(DS)
                report2.SetParameterValue(report2.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report2.SetParameterValue(report2.Parameter_Brand.ParameterFieldName, Brand)
                report2.SetParameterValue(report2.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))
                report2.SetParameterValue(report2.Parameter_Code.ParameterFieldName, "Account")

                Me.crvReport.ReportSource = report2
                Me.crvReport.Show()
            Else
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
                report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))
                report.SetParameterValue(report.Parameter_Code.ParameterFieldName, "Account")

                Me.crvReport.ReportSource = report
                Me.crvReport.Show()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            objCD = Nothing
            'DTA.Dispose()
        End Try
    End Sub

    Private Sub CetakRekap()
        Dim DS As New DseLaporanInventoryMovementReport
        'Dim DTA As New DseLaporanInventoryMovementReportTableAdapters.laporan_inventory_movement_purchase_view2TableAdapter
        '     Dim row As DseLaporanInternalPurchaseOrder.internal_purchase_order_report_viewRow
        Dim report As New CrpLaporanInventoryMovementRekaps
        Dim report2 As New CrpLaporanInventoryMovementRekaps1
        Dim JmlBegining, JmlIn, JmlOut, JmlEnding As Decimal

        Dim counter As Double = 0
        Dim row As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2Row
        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim DV As DataRowView
        Dim primaryBrand As Double = 0
        Dim Brand As String = ""

        Dim DTA_CL As New DseLaporanInventoryMovementReportTableAdapters.tbltrans_purchase_closing_item_detail1TableAdapter
        Dim TBL_CL As New DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1DataTable

        Dim rowItem, rowItemDummy As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2Row
        Dim TBLItemDummy As New DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2DataTable

        Try
            DTA_CL.Connection.ConnectionString = Core_Module.GetConstring
            DS.laporan_inventory_movement_purchase_view2.Rows.Clear()

            'isikan data
            If Me.clbBrand.Items.Count = Me.clbBrand.CheckedItems.Count Then
                'begining balance
                DTA_CL.Fill(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem)
                Me.IsikanBeginingAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                'in amount
                DTA_CL.Fill(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem)
                Me.IsikanInAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                'out amount
                DTA_CL.Fill(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem)
                Me.IsikanOutAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                    'begining balance
                    DTA_CL.FillBy_PrimaryBrand(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem)
                    Me.IsikanBeginingAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                    'in amount
                    DTA_CL.FillBy_PrimaryBrand(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem)
                    Me.IsikanInAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                    'out amount
                    DTA_CL.FillBy_PrimaryBrand(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem)
                    Me.IsikanOutAmount(DS.laporan_inventory_movement_purchase_view2, TBL_CL)
                Next
            End If

            'hitung ending balance
            For Each rowItem In _
            DS.laporan_inventory_movement_purchase_view2.Rows
                rowItem.Ending_Amount = rowItem.Begin_Amount + rowItem.In_Amount - rowItem.Out_Amount
            Next

            If Me.clbBrand.Items.Count = Me.clbBrand.CheckedItems.Count Then
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

            'urutkan ulang berdasarkan brand
            TBLItemDummy.Rows.Clear()
            For Each rowItem In _
            DS.laporan_inventory_movement_purchase_view2.Rows
                rowItemDummy = TBLItemDummy.Newlaporan_inventory_movement_purchase_view2Row
                For ijk As Integer = 0 To TBLItemDummy.Columns.Count - 1
                    rowItemDummy.Item(ijk) = rowItem.Item(ijk)
                Next
                TBLItemDummy.Rows.Add(rowItemDummy)
            Next
            DS.laporan_inventory_movement_purchase_view2.Rows.Clear()
            TBLItemDummy.DefaultView.Sort = "[Brand Code] ASC"
            TBLItemDummy.DefaultView.RowFilter = ""
            For Each DV In TBLItemDummy.DefaultView
                rowItemDummy = DV.Row
                rowItem = DS.laporan_inventory_movement_purchase_view2.Newlaporan_inventory_movement_purchase_view2Row
                For ijk As Integer = 0 To TBLItemDummy.Columns.Count - 1
                    rowItem.Item(ijk) = rowItemDummy.Item(ijk)
                Next
                DS.laporan_inventory_movement_purchase_view2.Rows.Add(rowItem)
            Next

            If Me.clbBrand.Items.Count = Me.clbBrand.CheckedItems.Count Then
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

            'filter laporan yang kosong semua tidak usah ditampilkan
            counter = 0
            JmlBegining = 0
            JmlIn = 0
            JmlOut = 0
            JmlEnding = 0
            While counter < DS.laporan_inventory_movement_purchase_view2.Rows.Count
                row = DS.laporan_inventory_movement_purchase_view2.Rows(counter)
                If row.Begin_Amount = 0 And row.Ending_Amount = 0 And _
                row.In_Amount = 0 And row.Out_Amount = 0 Then
                    DS.laporan_inventory_movement_purchase_view2.Rows.RemoveAt(counter)
                Else
                    JmlBegining += row.Begin_Amount
                    JmlIn += row.In_Amount
                    JmlOut += row.Out_Amount
                    JmlEnding += row.Ending_Amount
                    counter += 1
                End If
            End While

            If JmlBegining > Me.RefVal Or JmlIn > Me.RefVal Or JmlOut > Me.RefVal Or JmlEnding > Me.RefVal Then
                report2.SetDataSource(DS)
                report2.SetParameterValue(report2.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report2.SetParameterValue(report2.Parameter_Brand.ParameterFieldName, Brand)
                report2.SetParameterValue(report2.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))
                report2.SetParameterValue(report2.Parameter_Code.ParameterFieldName, "Brand")

                Me.crvReport.ReportSource = report2
                Me.crvReport.Show()
            Else
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
                report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))
                report.SetParameterValue(report.Parameter_Code.ParameterFieldName, "Brand")

                Me.crvReport.ReportSource = report
                Me.crvReport.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            objCD = Nothing
            TBLItemDummy = Nothing
            'DTA.Dispose()
        End Try
    End Sub

    Private Sub IsikanBeginingAmount(ByVal vTBL As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2DataTable, _
    ByVal vTBL_Amount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1DataTable)
        Dim rowAmount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1Row
        Dim vrow As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2Row

        Try
            For Each rowAmount In vTBL_Amount.Rows
                vrow = vTBL.Newlaporan_inventory_movement_purchase_view2Row

                vrow.Begin_Amount = rowAmount.Amount
                vrow.Brand_Code = rowAmount.Brand_Code
                vrow.Brand_Desc = rowAmount.Brand_Desc
                vrow.Ending_Amount = 0
                vrow.In_Amount = 0
                vrow.Out_Amount = 0

                vTBL.Rows.Add(vrow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanInAmount(ByVal vTBL As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2DataTable, _
    ByVal vTBL_Amount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1DataTable)
        Dim Filter As String = ""
        Dim rowAmount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1Row
        Dim vrow As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2Row

        Try
            For Each rowAmount In vTBL_Amount.Rows
                Filter = "[Brand Code] = '" + rowAmount.Brand_Code + "'"
                vTBL.DefaultView.RowFilter = Filter

                If vTBL.DefaultView.Count = 0 Then
                    vrow = vTBL.Newlaporan_inventory_movement_purchase_view2Row

                    vrow.Begin_Amount = 0
                    vrow.Brand_Code = rowAmount.Brand_Code
                    vrow.Brand_Desc = rowAmount.Brand_Desc
                    vrow.Ending_Amount = 0
                    vrow.In_Amount = rowAmount.Amount
                    vrow.Out_Amount = 0

                    vTBL.Rows.Add(vrow)
                Else
                    vrow = vTBL.DefaultView(0).Row
                    vrow.In_Amount += rowAmount.Amount
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanOutAmount(ByVal vTBL As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2DataTable, _
    ByVal vTBL_Amount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1DataTable)
        Dim Filter As String = ""
        Dim rowAmount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail1Row
        Dim vrow As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_view2Row

        Try
            For Each rowAmount In vTBL_Amount.Rows
                Filter = "[Brand Code] = '" + rowAmount.Brand_Code + "'"
                vTBL.DefaultView.RowFilter = Filter

                If vTBL.DefaultView.Count = 0 Then
                    vrow = vTBL.Newlaporan_inventory_movement_purchase_view2Row

                    vrow.Begin_Amount = 0
                    vrow.Brand_Code = rowAmount.Brand_Code
                    vrow.Brand_Desc = rowAmount.Brand_Desc
                    vrow.Ending_Amount = 0
                    vrow.In_Amount = 0
                    vrow.Out_Amount = rowAmount.Amount

                    vTBL.Rows.Add(vrow)
                Else
                    vrow = vTBL.DefaultView(0).Row
                    vrow.Out_Amount += rowAmount.Amount
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak()
        Dim DS As New DseLaporanInventoryMovementReport
        'Dim DTA As New DseLaporanInventoryMovementReportTableAdapters.laporan_inventory_movement_purchase_viewTableAdapter
        '     Dim row As DseLaporanInternalPurchaseOrder.internal_purchase_order_report_viewRow
        Dim report As New CrpLaporanInventoryMovementReports3
        Dim report2 As New CrpLaporanInventoryMovementReports3_1
        Dim JmlBegining, JmlIn, JmlOut, JmlEnding As Decimal

        Dim counter As Double = 0
        Dim row, row1 As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewRow
        Dim TBLDummy As New DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewDataTable

        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim DV As DataRowView
        Dim primaryBrand As Double = 0
        Dim Brand As String = ""

        Dim DTA_CL As New DseLaporanInventoryMovementReportTableAdapters.tbltrans_purchase_closing_item_detail3TableAdapter
        Dim TBL_CL As New DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail3DataTable
        Dim DTA_Account As New DseLaporanInventoryMovementReportTableAdapters.TBL_AccountTableAdapter
        Dim TBL_Account As New DseLaporanInventoryMovementReport.TBL_AccountDataTable
        Dim row_Account As DseLaporanInventoryMovementReport.TBL_AccountRow

        Try
            DTA_CL.Connection.ConnectionString = Core_Module.GetConstring
            DS.laporan_inventory_movement_purchase_view.Rows.Clear()

            If Me.clbBrand.Items.Count = Me.clbBrand.CheckedItems.Count Then
                DTA_CL.Fill(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem)
                Me.IsikanBeginingAmount1(DS.laporan_inventory_movement_purchase_view, TBL_CL)

                DTA_CL.Fill(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem)
                Me.IsikanInAmount1(DS.laporan_inventory_movement_purchase_view, TBL_CL)

                DTA_CL.Fill(TBL_CL, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), _
                ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem)
                Me.IsikanOutAmount1(DS.laporan_inventory_movement_purchase_view, TBL_CL)
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)

                    DTA_CL.FillBy_PrimaryBrand(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem)
                    Me.IsikanBeginingAmount1(DS.laporan_inventory_movement_purchase_view, TBL_CL)

                    DTA_CL.FillBy_PrimaryBrand(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem)
                    Me.IsikanInAmount1(DS.laporan_inventory_movement_purchase_view, TBL_CL)

                    DTA_CL.FillBy_PrimaryBrand(TBL_CL, primaryBrand, objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value), _
                    ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem)
                    Me.IsikanOutAmount1(DS.laporan_inventory_movement_purchase_view, TBL_CL)
                Next
            End If

            For Each rowItem As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewRow In _
            DS.laporan_inventory_movement_purchase_view.Rows
                rowItem.Ending_Amount = rowItem.Begin_Amount + rowItem.In_Amount - rowItem.Out_Amount
                rowItem.Ending_Qty = rowItem.Begin_Qty + rowItem.In_Qty - rowItem.Out_Qty
            Next

            If Me.clbBrand.Items.Count = Me.clbBrand.CheckedItems.Count Then
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

            'filter laporan yang kosong semua tidak usah ditampilkan
            counter = 0
            JmlBegining = 0
            JmlIn = 0
            JmlOut = 0
            JmlEnding = 0
            While counter < DS.laporan_inventory_movement_purchase_view.Rows.Count
                row = DS.laporan_inventory_movement_purchase_view.Rows(counter)
                If row.Begin_Amount = 0 And row.Begin_Qty = 0 And row.Ending_Amount = 0 And _
                row.Ending_Qty = 0 And row.In_Amount = 0 And row.In_Qty = 0 And _
                row.Out_Amount = 0 And row.Out_Qty = 0 Then
                    DS.laporan_inventory_movement_purchase_view.Rows.RemoveAt(counter)
                Else
                    JmlBegining += row.Begin_Amount
                    JmlIn += row.In_Amount
                    JmlOut += row.Out_Amount
                    JmlEnding += row.Ending_Amount
                    counter += 1
                End If
            End While

            'jika group by account code, maka gantikan brand dengan account code
            If Me.cboTipe.SelectedIndex = 3 Then
                DTA_Account.Connection.ConnectionString = Core_Module.GetConstring
                DTA_Account.Fill(TBL_Account)
                counter = 0
                While counter < DS.laporan_inventory_movement_purchase_view.Rows.Count
                    row = DS.laporan_inventory_movement_purchase_view.Rows(counter)
                    TBL_Account.DefaultView.RowFilter = "[Primary Product] = " + row.Primary_Product.ToString
                    If TBL_Account.DefaultView.Count = 0 Then
                        MsgBox("Logic Error, please contact your vendor if you see this message")
                        Exit Sub
                    Else
                        row_Account = TBL_Account.DefaultView(0).Row
                        row.Brand_Code = row_Account.Account_Code
                        row.Brand_Desc = row_Account.Description
                    End If
                    counter += 1
                End While
            End If

            'sorting ulang, diurutkan berdasarkan brand code dan product code
            For Each row In DS.laporan_inventory_movement_purchase_view.Rows
                row1 = TBLDummy.Newlaporan_inventory_movement_purchase_viewRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    row1.Item(ij) = row.Item(ij)
                Next
                TBLDummy.Rows.Add(row1)
            Next
            TBLDummy.DefaultView.Sort = "[Brand Code], [Product Code]"
            DS.laporan_inventory_movement_purchase_view.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                row1 = DV.Row
                row = DS.laporan_inventory_movement_purchase_view.Newlaporan_inventory_movement_purchase_viewRow
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    row.Item(ij) = row1.Item(ij)
                Next
                DS.laporan_inventory_movement_purchase_view.Rows.Add(row)
            Next

            If JmlBegining > Me.RefVal Or JmlIn > Me.RefVal Or JmlOut > Me.RefVal Or JmlEnding > Me.RefVal Then
                report2.SetDataSource(DS)
                report2.SetParameterValue(report2.Parameter_Brand.ParameterFieldName, Brand)
                If Me.cboTipe.SelectedIndex = 2 Then
                    report2.SetParameterValue(report2.Parameter_Group.ParameterFieldName, "Brand")
                Else
                    report2.SetParameterValue(report2.Parameter_Group.ParameterFieldName, "Account")
                End If
                report2.SetParameterValue(report2.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
                report2.SetParameterValue(report2.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
                report2.SetParameterValue(report2.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
                report2.SetParameterValue(report2.Parameter_Tgl.ParameterFieldName, Me.dtpFrom.Value)
                report2.SetParameterValue(report2.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)

                Me.crvReport.ReportSource = report2
                Me.crvReport.Show()
            Else
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
                If Me.cboTipe.SelectedIndex = 2 Then
                    report.SetParameterValue(report.Parameter_Group.ParameterFieldName, "Brand")
                Else
                    report.SetParameterValue(report.Parameter_Group.ParameterFieldName, "Account")
                End If
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
                report.SetParameterValue(report.Parameter_Tgl.ParameterFieldName, Me.dtpFrom.Value)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)

                Me.crvReport.ReportSource = report
                Me.crvReport.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            'DTA.Dispose()
            objCD = Nothing
            DTA_CL.Dispose()
            TBL_CL.Dispose()
            DTA_Account.Dispose()
            TBL_Account.Dispose()
            TBLDummy.Dispose()
        End Try
    End Sub

    Private Sub IsikanBeginingAmount1(ByVal vTBL As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewDataTable, _
    ByVal vTBL_Amount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail3DataTable)
        Dim rowAmount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail3Row
        Dim vrow As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewRow

        Try
            For Each rowAmount In vTBL_Amount.Rows
                vrow = vTBL.Newlaporan_inventory_movement_purchase_viewRow

                vrow.Brand_Code = rowAmount.Brand_Code
                vrow.Brand_Desc = rowAmount.Brand_Desc
                vrow.Begin_Qty = rowAmount.Qty
                vrow.Begin_Amount = rowAmount.Amount
                vrow.In_Amount = 0
                vrow.In_Qty = 0
                vrow.Out_Amount = 0
                vrow.Out_Qty = 0
                vrow.Ending_Amount = 0
                vrow.Ending_Qty = 0

                vrow.Category_Code = rowAmount.Category_Code
                vrow.Deskripsi_Kategori = rowAmount.Deskripsi_Kategori
                vrow.Description = rowAmount.Description
                vrow.Primary_Product = rowAmount.primary_code_no
                vrow.Primary_Unit = rowAmount.Primary_Unit
                vrow.Product_Code = rowAmount.Product_Code
                vrow.Unit_Code = rowAmount.Unit_Code
                vrow._Decimal = rowAmount._Decimal

                vTBL.Rows.Add(vrow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanInAmount1(ByVal vTBL As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewDataTable, _
   ByVal vTBL_Amount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail3DataTable)
        Dim Filter As String = ""
        Dim rowAmount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail3Row
        Dim vrow As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewRow

        Try
            For Each rowAmount In vTBL_Amount.Rows
                Filter = "[Primary Product] = " + rowAmount.primary_code_no.ToString + ""
                vTBL.DefaultView.RowFilter = Filter

                If vTBL.DefaultView.Count = 0 Then
                    vrow = vTBL.Newlaporan_inventory_movement_purchase_viewRow

                    vrow.Brand_Code = rowAmount.Brand_Code
                    vrow.Brand_Desc = rowAmount.Brand_Desc
                    vrow.Begin_Qty = 0
                    vrow.Begin_Amount = 0
                    vrow.In_Amount = rowAmount.Amount
                    vrow.In_Qty = rowAmount.Qty
                    vrow.Out_Amount = 0
                    vrow.Out_Qty = 0
                    vrow.Ending_Amount = 0
                    vrow.Ending_Qty = 0

                    vrow.Category_Code = rowAmount.Category_Code
                    vrow.Deskripsi_Kategori = rowAmount.Deskripsi_Kategori
                    vrow.Description = rowAmount.Description
                    vrow.Primary_Product = rowAmount.primary_code_no
                    vrow.Primary_Unit = rowAmount.Primary_Unit
                    vrow.Product_Code = rowAmount.Product_Code
                    vrow.Unit_Code = rowAmount.Unit_Code
                    vrow._Decimal = rowAmount._Decimal

                    vTBL.Rows.Add(vrow)
                Else
                    vrow = vTBL.DefaultView(0).Row
                    vrow.In_Amount += rowAmount.Amount
                    vrow.In_Qty += rowAmount.Qty
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanOutAmount1(ByVal vTBL As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewDataTable, _
    ByVal vTBL_Amount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail3DataTable)
        Dim Filter As String = ""
        Dim rowAmount As DseLaporanInventoryMovementReport.tbltrans_purchase_closing_item_detail3Row
        Dim vrow As DseLaporanInventoryMovementReport.laporan_inventory_movement_purchase_viewRow

        Try
            For Each rowAmount In vTBL_Amount.Rows
                Filter = "[Primary Product] = " + rowAmount.primary_code_no.ToString + ""
                vTBL.DefaultView.RowFilter = Filter

                If vTBL.DefaultView.Count = 0 Then
                    vrow = vTBL.Newlaporan_inventory_movement_purchase_viewRow

                    vrow.Brand_Code = rowAmount.Brand_Code
                    vrow.Brand_Desc = rowAmount.Brand_Desc
                    vrow.Begin_Qty = 0
                    vrow.Begin_Amount = 0
                    vrow.In_Amount = 0
                    vrow.In_Qty = 0
                    vrow.Out_Amount = rowAmount.Amount
                    vrow.Out_Qty = rowAmount.Qty
                    vrow.Ending_Amount = 0
                    vrow.Ending_Qty = 0

                    vrow.Category_Code = rowAmount.Category_Code
                    vrow.Deskripsi_Kategori = rowAmount.Deskripsi_Kategori
                    vrow.Description = rowAmount.Description
                    vrow.Primary_Product = rowAmount.primary_code_no
                    vrow.Primary_Unit = rowAmount.Primary_Unit
                    vrow.Product_Code = rowAmount.Product_Code
                    vrow.Unit_Code = rowAmount.Unit_Code
                    vrow._Decimal = rowAmount._Decimal

                    vTBL.Rows.Add(vrow)
                Else
                    vrow = vTBL.DefaultView(0).Row
                    vrow.Out_Amount += rowAmount.Amount
                    vrow.Out_Qty += rowAmount.Qty
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub FrmLaporanInventoryMovementReports_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.cboTipe.SelectedIndex = 0
            Me.FormatDecimal = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.dtpFrom.Value = ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)

            'checked semua brand
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next

            Me.RefVal = 10 ^ (13 - Me.ObjCur.GetDecimal_DefaultCurrency)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBjBrand = Nothing
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetPrimaryPurchaseClosing(ByVal vClosingPeriod As Date) As Double
        Dim DTA As New DseLaporanInventoryMovementDetailTableAdapters.tbltrans_purchase_closingTableAdapter
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

    Private Sub Show_TransaksiDetail2()
        Dim Report As New CrpLaporanInventoryMovementReports4
        Dim Report2 As New CrpLaporanInventoryMovementReports4_3
        Dim Report3 As New CrpLaporanInventoryMovementReports5
        Dim Report4 As New CrpLaporanInventoryMovementReports5_3
        Dim Report_Brand As New CrpLaporanInventoryMovementReports4_1
        Dim Report_Brand2 As New CrpLaporanInventoryMovementReports4_2
        Dim Report_Brand3 As New CrpLaporanInventoryMovementReports5_1
        Dim Report_Brand4 As New CrpLaporanInventoryMovementReports5_2
        Dim DS As New DseLaporanInventoryMovementDetail
        Dim TBL As DseLaporanInventoryMovementDetail.DataTable1DataTable = DS.DataTable1
        Dim Row As DseLaporanInventoryMovementDetail.DataTable1Row
        Dim TBLDummy As New DseLaporanInventoryMovementDetail.DataTable1DataTable
        Dim RowDummy As DseLaporanInventoryMovementDetail.DataTable1Row
        Dim DTA1 As New DseLaporanInventoryMovementDetailTableAdapters.tbltrans_purchase_closing_item_detail1TableAdapter
        Dim TBL1 As New DseLaporanInventoryMovementDetail.tbltrans_purchase_closing_item_detail1DataTable
        Dim Row1 As DseLaporanInventoryMovementDetail.tbltrans_purchase_closing_item_detail1Row
        Dim DTA2 As New DseLaporanInventoryMovementDetailTableAdapters.DataTable2TableAdapter
        Dim TBL2 As New DseLaporanInventoryMovementDetail.DataTable2DataTable
        Dim Row2 As DseLaporanInventoryMovementDetail.DataTable2Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryClosing As Double = -1
        Dim PeriodClosing As String = Format(ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date), "dd MMM yyyy")
        Dim Filter As String = ""
        Dim Combination As String = ""
        Dim DV As Data.DataRowView
        Dim ListCombination As New List(Of String)
        Dim ListBarisVisible As New List(Of Integer)
        Dim AdaProduksi As Boolean = False
        Dim FormatDecimal As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim PrimaryBrand2 As Double = -1
        Dim RowBrand As Core_Module.dseMasterProductGroup.tblmaster_product_group2Row
        Dim i As Integer = 0
        Dim JmlBegining, JmlIn, JmlOut, JmlEnding As Decimal

        Try
            PrimaryClosing = Me.GetPrimaryPurchaseClosing(PeriodClosing)
            If PrimaryClosing = -1 Then
                Me.crvReport.ReportSource = Nothing
                Me.crvReport.Show()
                Exit Sub
            End If

            'isikan tabel acuan kode brand dan kode akun
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA2.Fill(TBL2)

            For Each DV_Brand As DataRowView In Me.clbBrand.CheckedItems
                RowBrand = DV_Brand.Row
                DTA1.FillBy_PrimaryClosing_PrimaryBrand(TBL1, PrimaryClosing, RowBrand.Primary)
                If TBL1.Rows.Count > 0 Then
                    For Each Row1 In TBL1.Rows
                        If Row1.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
                            Filter = "[Primary Code No] = " + Row1.primary_code_no.ToString + ""
                            TBL.DefaultView.RowFilter = Filter
                            If TBL.DefaultView.Count > 0 Then
                                Row = TBL.DefaultView(0).Row
                                Row.Begining += Row1.amount
                            Else
                                Row = TBL.NewDataTable1Row
                                Row.Begining = Row1.amount
                                Row.Primary_Code_No = Row1.primary_code_no
                                TBL.Rows.Add(Row)
                            End If
                        ElseIf Row1.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
                            Filter = "[Primary Code No] = " + Row1.primary_code_no.ToString + ""
                            TBL.DefaultView.RowFilter = Filter
                            If TBL.DefaultView.Count > 0 Then
                                If Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.In_Adjustment += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.In_Production += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.In_Purchase += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.In_Sales_Return += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.In_Stock_Take += Row1.amount
                                End If
                            Else
                                If Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                                    Row = TBL.NewDataTable1Row
                                    Row.In_Adjustment = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                                    Row = TBL.NewDataTable1Row
                                    Row.In_Production = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                                    Row = TBL.NewDataTable1Row
                                    Row.In_Purchase = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                                    Row = TBL.NewDataTable1Row
                                    Row.In_Sales_Return = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                                    Row = TBL.NewDataTable1Row
                                    Row.In_Stock_Take = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                End If
                            End If
                        ElseIf Row1.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
                            Filter = "[Primary Code No] = " + Row1.primary_code_no.ToString + ""
                            TBL.DefaultView.RowFilter = Filter
                            If TBL.DefaultView.Count > 0 Then
                                If Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.Out_Adjustment += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.Out_Production += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.Out_Purchase_Return += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.Out_Sales += Row1.amount
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                                    Row = TBL.DefaultView(0).Row
                                    Row.Out_Stock_Take += Row1.amount
                                End If
                            Else
                                If Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                                    Row = TBL.NewDataTable1Row
                                    Row.Out_Adjustment = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                                    Row = TBL.NewDataTable1Row
                                    Row.Out_Production = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                                    Row = TBL.NewDataTable1Row
                                    Row.Out_Purchase_Return = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                                    Row = TBL.NewDataTable1Row
                                    Row.Out_Sales = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                ElseIf Row1.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                                    Row = TBL.NewDataTable1Row
                                    Row.Out_Stock_Take = Row1.amount
                                    Row.Primary_Code_No = Row1.primary_code_no
                                    TBL.Rows.Add(Row)
                                End If
                            End If
                        ElseIf Row1.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
                            Filter = "[Primary Code No] = " + Row1.primary_code_no.ToString + ""
                            TBL.DefaultView.RowFilter = Filter
                            If TBL.DefaultView.Count > 0 Then
                                Row = TBL.DefaultView(0).Row
                                Row.Ending += Row1.amount
                            Else
                                Row = TBL.NewDataTable1Row
                                Row.Ending = Row1.amount
                                Row.Primary_Code_No = Row1.primary_code_no
                                TBL.Rows.Add(Row)
                            End If
                        End If
                    Next
                End If
            Next

            'Masukan product code, group code, ......
            If TBL.Rows.Count > 0 Then
                For Each Row In TBL.Rows
                    Filter = "[Primary] = " + Row.Primary_Code_No.ToString + ""
                    TBL2.DefaultView.RowFilter = Filter
                    If TBL2.DefaultView.Count > 0 Then
                        Row2 = TBL2.DefaultView(0).Row
                        Row.Code_No = Row2.Product_Code
                        Row.Description = Row2.Description
                        If Me.cboTipe.SelectedIndex = 0 Or Me.cboTipe.SelectedIndex = 2 Then
                            Row.Group_Code = Row2.Brand_Code
                            Row.Group_Description = Row2.Brand_Desc
                            Row.Combination = Row2.Brand_Code + " - " + Row2.Brand_Desc
                        Else
                            Row.Group_Code = Row2.Account_Code
                            Row.Group_Description = Row2.Account_Desc
                            Row.Combination = Row2.Account_Code + " - " + Row2.Account_Desc
                        End If

                    End If
                Next
            End If

            'Hapus urutkan dan tandai group terakhir
            For Each Row In TBL.Rows
                If ListCombination.Contains(Row.Combination) = False Then
                    ListCombination.Add(Row.Combination)
                End If
            Next

            TBL.DefaultView.Sort = "[Combination], [Code No]"
            If ListCombination.Count > 0 Then
                i = 0
                While i <= ListCombination.Count - 1
                    Combination = ListCombination(i)
                    Filter = "[Combination] = '" + ObjCD.FormatSpecialChar_DataFilter(Combination.ToString) + "'"
                    TBL.DefaultView.RowFilter = Filter
                    TBL.DefaultView.Sort = "[Combination], [Code No]"
                    If TBL.DefaultView.Count > 0 Then
                        If i = ListCombination.Count - 1 Then
                            ListBarisVisible.Add(0)
                        Else
                            ListBarisVisible.Add(1)
                        End If
                        For Each DV In TBL.DefaultView
                            Row = DV.Row
                            RowDummy = TBLDummy.NewDataTable1Row
                            For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                                RowDummy.Item(ij) = Row.Item(ij)
                            Next
                            TBLDummy.Rows.Add(RowDummy)
                        Next
                        i += 1
                    Else
                        If i = 0 Then
                            ListBarisVisible.Add(0)
                            i += 1
                        Else
                            ListBarisVisible.Item(i - 1) = 0
                            ListCombination.RemoveAt(i)
                        End If
                    End If
                End While
            End If
            TBL.Rows.Clear()
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = ""
            JmlBegining = 0
            JmlIn = 0
            JmlOut = 0
            JmlEnding = 0
            If TBLDummy.Rows.Count > 0 Then
                For Each RowDummy In TBLDummy.Rows
                    Row = TBL.NewDataTable1Row
                    For ij As Integer = 0 To TBL.Columns.Count - 1
                        If ij = 18 Then
                            If RowDummy.Combination = ListCombination.Item(ListCombination.Count - 1) _
                            And ListBarisVisible.Item(ListBarisVisible.Count - 1) = 0 Then
                                Row.Item(ij) = 0
                            End If
                        Else
                            Row.Item(ij) = RowDummy.Item(ij)
                        End If
                    Next
                    TBL.Rows.Add(Row)
                    JmlBegining += Row.Begining
                    JmlEnding += Row.Ending
                    JmlIn += (Row.In_Adjustment + Row.In_Production + Row.In_Purchase + Row.In_Sales_Return + Row.In_Stock_Take)
                    JmlOut += (Row.Out_Adjustment + Row.Out_Production + Row.Out_Purchase_Return + Row.Out_Sales + Row.Out_Stock_Take)
                Next
            End If

            AdaProduksi = False
            For Each Row In TBL.Rows
                If Row.Out_Production <> 0 Or Row.In_Production <> 0 Then
                    AdaProduksi = True
                    Exit For
                End If
            Next
            If JmlBegining > Me.RefVal Or JmlIn > Me.RefVal Or JmlOut > Me.RefVal Or JmlEnding > Me.RefVal Then
                If Me.cboTipe.SelectedIndex = 0 Or Me.cboTipe.SelectedIndex = 1 Then
                    If AdaProduksi = True Then
                        Report_Brand3.SetDataSource(DS)
                        Report_Brand3.SetParameterValue(Report_Brand3.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report_Brand3.SetParameterValue(Report_Brand3.Parameter_Decimal.ParameterFieldName, FormatDecimal)
                        If Me.cboTipe.SelectedIndex = 0 Then
                            Report_Brand3.SetParameterValue(Report_Brand3.Parameter_Group.ParameterFieldName, "Brand")
                        Else
                            Report_Brand3.SetParameterValue(Report_Brand3.Parameter_Group.ParameterFieldName, "Account")
                        End If
                        Me.crvReport.ReportSource = Report_Brand3
                        Me.crvReport.Show()
                    Else
                        Report_Brand4.SetDataSource(DS)
                        Report_Brand4.SetParameterValue(Report_Brand4.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report_Brand4.SetParameterValue(Report_Brand4.Parameter_Decimal.ParameterFieldName, FormatDecimal)
                        If Me.cboTipe.SelectedIndex = 0 Then
                            Report_Brand4.SetParameterValue(Report_Brand4.Parameter_Group.ParameterFieldName, "Brand")
                        Else
                            Report_Brand4.SetParameterValue(Report_Brand4.Parameter_Group.ParameterFieldName, "Account")
                        End If
                        Me.crvReport.ReportSource = Report_Brand4
                        Me.crvReport.Show()
                    End If
                Else
                    If AdaProduksi = True Then
                        Report3.SetDataSource(DS)
                        Report3.SetParameterValue(Report3.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report3.SetParameterValue(Report3.Parameter_Decimal.ParameterFieldName, FormatDecimal)
                        Me.crvReport.ReportSource = Report3
                        Me.crvReport.Show()
                    Else
                        Report4.SetDataSource(DS)
                        Report4.SetParameterValue(Report4.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report4.SetParameterValue(Report4.Parameter_Decimal.ParameterFieldName, FormatDecimal)

                        Me.crvReport.ReportSource = Report4
                        Me.crvReport.Show()
                    End If
                End If
            Else
                If Me.cboTipe.SelectedIndex = 0 Or Me.cboTipe.SelectedIndex = 1 Then
                    If AdaProduksi = True Then
                        Report_Brand.SetDataSource(DS)
                        Report_Brand.SetParameterValue(Report_Brand.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report_Brand.SetParameterValue(Report_Brand.Parameter_Decimal.ParameterFieldName, FormatDecimal)
                        If Me.cboTipe.SelectedIndex = 0 Then
                            Report_Brand.SetParameterValue(Report_Brand.Parameter_Group.ParameterFieldName, "Brand")
                        Else
                            Report_Brand.SetParameterValue(Report_Brand.Parameter_Group.ParameterFieldName, "Account")
                        End If
                        Me.crvReport.ReportSource = Report_Brand
                        Me.crvReport.Show()
                    Else
                        Report_Brand2.SetDataSource(DS)
                        Report_Brand2.SetParameterValue(Report_Brand2.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report_Brand2.SetParameterValue(Report_Brand2.Parameter_Decimal.ParameterFieldName, FormatDecimal)
                        If Me.cboTipe.SelectedIndex = 0 Then
                            Report_Brand2.SetParameterValue(Report_Brand2.Parameter_Group.ParameterFieldName, "Brand")
                        Else
                            Report_Brand2.SetParameterValue(Report_Brand2.Parameter_Group.ParameterFieldName, "Account")
                        End If
                        Me.crvReport.ReportSource = Report_Brand2
                        Me.crvReport.Show()
                    End If
                Else
                    If AdaProduksi = True Then
                        Report.SetDataSource(DS)
                        Report.SetParameterValue(Report.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report.SetParameterValue(Report.Parameter_Decimal.ParameterFieldName, FormatDecimal)
                        Me.crvReport.ReportSource = Report
                        Me.crvReport.Show()
                    Else
                        Report2.SetDataSource(DS)
                        Report2.SetParameterValue(Report2.Parameter_Period.ParameterFieldName, Me.dtpFrom.Value.Date)
                        Report2.SetParameterValue(Report2.Parameter_Decimal.ParameterFieldName, FormatDecimal)

                        Me.crvReport.ReportSource = Report2
                        Me.crvReport.Show()
                    End If
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
            TBLDummy.Dispose()
            ObjCD = Nothing
            ListCombination = Nothing
            ListBarisVisible = Nothing
        End Try
    End Sub
End Class