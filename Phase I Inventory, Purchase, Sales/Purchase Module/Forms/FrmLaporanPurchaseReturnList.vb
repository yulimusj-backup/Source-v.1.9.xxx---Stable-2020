Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseReturnList
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim TanggalClosing As Date = #1/1/1950#
    Dim objCD As New Core_Module.ClsCekData

    Dim DTAClosing As New DsePurchaseReturnListTableAdapters.DataTable2TableAdapter
    Dim TBLClosing As New DsePurchaseReturnList.DataTable2DataTable

    Public DSOut As New DsePurchaseReturnList

    Dim DefPrec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjClosing As New ClsTransaksiMonthlyClosing
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub Generate_Purchase_Return_List2()
        Dim DS As New DsePurchaseReturnList
        Dim report As New CrpLaporanPurchaseReturnListC
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim TBL As DsePurchaseReturnList.DataTable1DataTable = DS.DataTable1
        Dim row As DsePurchaseReturnList.DataTable1Row
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = 0
        Dim DTA1 As New DsePurchaseReturnListTableAdapters.DataTable3TableAdapter
        Dim TBL1 As New DsePurchaseReturnList.DataTable3DataTable
        Dim row1 As DsePurchaseReturnList.DataTable3Row
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
        Dim PrimaryClosing As Double = -1
        Dim TipeProduct As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods


        Try
            PrimaryClosing = Me.ObjClosing.GetPrimary(TglDari)
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            TBL1.Rows.Clear()
            DTA1.ClearBeforeFill = False

            If Me.chkGoods.Checked = True Then
                TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
            Else
                TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
            End If
            If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
                If Me.chkGoods.Checked = True And Me.chkService.Checked = True Then
                    DTA1.Fill(TBL1, TglSampai, TglDari, PrimaryClosing, TipeItemDetail, TipeTransaksiAsal)
                Else
                    DTA1.FillBy_TipeProduct1(TBL1, TglSampai, TglDari, TipeProduct, PrimaryClosing, TipeItemDetail, TipeTransaksiAsal)
                End If
            Else
                If Me.chkGoods.Checked = True And Me.chkService.Checked = True Then
                    For Each DV In Me.clbBrand.CheckedItems
                        primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                        DTA1.FillBy_PrimaryBrand(TBL1, TglSampai, TglDari, primaryBrand, PrimaryClosing, TipeItemDetail, TipeTransaksiAsal)
                    Next
                Else
                    DTA1.FillBy_TipeProduct2(TBL1, TglSampai, TglDari, primaryBrand, TipeProduct, PrimaryClosing, TipeItemDetail, TipeTransaksiAsal)
                End If
            End If

            DS.DataTable1.Rows.Clear()
            For Each row1 In TBL1.Rows
                row = TBL.NewDataTable1Row
                With row
                    .Brand = row1.Brand
                    .Cost = row1.amount
                    .No_Stock_Out_Purchase_Return = row1.No_Transaksi
                    .Primary = row1.primary
                    .Primary_Brand = row1.Primary_Brand
                    .Primary_Product = row1.primary_code_no
                    .Primary_Supplier = row1.primary_supplier
                    .Product_Code = row1.Product_Code
                    .Qty = row1.qty
                    .Supplier = row1.Supplier
                    .Tanggal_Transaksi = row1.Tanggal_Transaksi
                    .Unit_Code = row1.Unit_Code
                    ._Decimal = row1._Decimal
                    .Return_Price = row1.return_price
                End With
                TBL.Rows.Add(row)
            Next

            Dim Brand As String = ""
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
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.DefPrec)
            report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "Purchase Return Cost Report")
            report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    'Private Sub Generate_Purchase_Return_List()
    '    If Me.ApakahBolehCetak = False Then
    '        Exit Sub
    '    End If
    '    Me.Generate_Purchase_Return_List2()
    '    Exit Sub

    '    Dim DS As New DsePurchaseReturnList
    '    Dim DTA As New DsePurchaseReturnListTableAdapters.DataTable1TableAdapter
    '    Dim report As New CrpLaporanPurchaseReturnListB
    '    Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '    Dim row As DsePurchaseReturnList.DataTable1Row
    '    Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
    '    Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
    '    Dim DV As Data.DataRowView
    '    Dim primaryBrand As Double = 0

    '    Try
    '        'stock out purchase return
    '        DS.DataTable1.Rows.Clear()
    '        If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
    '            DTA.FillBy_All(DS.DataTable1, _
    '            TglSampai, TglDari)
    '        Else
    '            For Each DV In Me.clbBrand.CheckedItems
    '                primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
    '                'MsgBox(primaryBrand)
    '                DTA.ClearBeforeFill = False
    '                DTA.FillBy_Supplier(DS.DataTable1, _
    '                TglSampai, TglDari, primaryBrand)
    '            Next
    '        End If

    '        'hitung cost inventory
    '        For Each row In DS.DataTable1.Rows
    '            row.Cost = row.Qty * Me.GetCostOfSales(row.Primary_Product, _
    '            Me.objCD.GetFirstDateofTheMonth(row.Tanggal_Transaksi))
    '        Next

    '        Dim Brand As String = ""
    '        If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
    '            Brand = "All"
    '        Else
    '            For Each DV In Me.clbBrand.CheckedItems
    '                If Brand = "" Then
    '                    Brand = DV.Item(1)
    '                Else
    '                    Brand += ", " + DV.Item(1)
    '                End If
    '            Next
    '        End If

    '        report.SetDataSource(DS)
    '        report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
    '        report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.DefPrec)
    '        report.SetParameterValue(report.Parameter_NamaLaporan.ParameterFieldName, "Purchase Return List")
    '        report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))

    '        Me.CrystalReportViewer1.ReportSource = report
    '        Me.CrystalReportViewer1.Show()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DS.Dispose()
    '        DTA.Dispose()
    '    End Try
    'End Sub

    Private Sub FrmLaporanSalesReturnList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup

        Try
            Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.cboGrouping.SelectedIndex = 0
            Me.dtpFrom.Value = Me.objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Me.Text = "Purchase Return Cost Report"
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

    Private Function GetCostOfSales(ByVal vPrimaryProduct As Double, ByVal vClosingPeriod As Date) As Decimal
        Dim filter As String = ""
        Dim row As DsePurchaseReturnList.DataTable2Row

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
                        Return Math.Round(row.Out_Amount / row.Out_Qty, Me.DefPrec, MidpointRounding.AwayFromZero)
                    End If
                Else
                    If row.Ending_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(row.Ending_Amount / row.Ending_Qty, Me.DefPrec, MidpointRounding.AwayFromZero)
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

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Dim PrimaryClosing As Double = -1

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

            If Me.ApakahBolehCetak = True Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Me.Generate_Purchase_Return_List2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
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

    Public Sub CekOut1(ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            'Me.cboPurchaseReturnTipe.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBjBrand = Nothing
        End Try

        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        Dim DS As New DsePurchaseReturnList
        Dim DTA As New DsePurchaseReturnListTableAdapters.DataTable1TableAdapter
        Dim report As New CrpLaporanPurchaseReturnListB
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row As DsePurchaseReturnList.DataTable1Row
        Dim TglDari As Date = vStartDate
        Dim TglSampai As Date = vEndDate
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = 0

        Try
            'stock out purchase return
            For Each DV In Me.clbBrand.CheckedItems
                primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                DTA.ClearBeforeFill = False
                DTA.FillBy_Supplier(DS.DataTable1, _
                TglSampai, TglDari, primaryBrand)
            Next

            'hitung cost inventory
            For Each row In DS.DataTable1.Rows
                row.Cost = row.Qty * Me.GetCostOfSales(row.Primary_Product, _
                Me.objCD.GetFirstDateofTheMonth(row.Tanggal_Transaksi))
            Next

            Me.DSOut = DS
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub cboGrouping_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGrouping.SelectedIndexChanged
        Try
            If Me.cboGrouping.SelectedIndex = 1 Then
                Me.gboSupplier.BringToFront()
            Else
                Me.gboBrand.BringToFront()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class