Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseAuditList
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim TanggalClosing As Date = #1/1/1950#
    Dim objCD As New Core_Module.ClsCekData

    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Sub btnShowReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReport.Click
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

            Me.Generate_revisi_invoice_baru()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Generate_revisi_invoice()
        If Me.ApakahBolehCetak = False Then
            Exit Sub
        End If

        Dim DS As New DseLaporanPurchaseAuditList
        Dim DTA As New DseLaporanPurchaseAuditListTableAdapters.DataTable1TableAdapter
        Dim report As New CrpLaporanPurchaseCost
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row, row1 As DseLaporanPurchaseAuditList.DataTable1Row
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = 0

        Dim DTAReturn As New DseLaporanPurchaseAuditListTableAdapters.DataTable2TableAdapter
        Dim TBLReturn As New DseLaporanPurchaseAuditList.DataTable2DataTable
        Dim rowReturn As DseLaporanPurchaseAuditList.DataTable2Row
        Dim Filter As String = ""
        Dim TBLDummy As New DseLaporanPurchaseAuditList.DataTable1DataTable
        Dim Sort As String = ""
        Dim rowView As DataRowView

        Dim AOPDummy As String = ""
        Dim BrandDummy As String = ""
        Dim i As Double = 0

        Try
            'in purchase
            For Each DV In Me.clbBrand.CheckedItems
                primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                DTA.ClearBeforeFill = False
                DTA.FillBy1(TBLDummy, _
                TglSampai, TglDari, primaryBrand)
            Next

            'kurangkan dengan data return 
            DTAReturn.Fill(TBLReturn, TglSampai, TglDari)
            For Each rowReturn In TBLReturn.Rows
                Filter = "[Primary Product] =" + rowReturn.Primary_Product.ToString + _
                " AND [AOP Number] = '" + rowReturn.AOP_Number + "' AND [Primary GRN] = " + _
                rowReturn.Primary_GRN.ToString + " AND [Primary PO] = " + rowReturn.Primary_PO.ToString + _
                " AND [Secondary Qty] = " + rowReturn.Secondary_Qty.ToString
                TBLDummy.DefaultView.RowFilter = Filter

                i = 0
                While TBLDummy.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row = TBLDummy.DefaultView(i).Row
                    If row.Total_Qty >= rowReturn.Total_Qty Then
                        row.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row.Total_Qty
                        row.Total_Qty = 0
                    End If

                    If row.Total_Qty = 0 Then
                        TBLDummy.Rows.Remove(row)
                    End If
                End While
            Next


            'betulan data qty str dan hitung total unit price
            For Each row In TBLDummy.Rows
                row.Total_Unit_Price = row.Unit_Price * row.Total_Qty
                If row.Sec_Unit_Code <> "" Then
                    row.Qty_Str = Me.objCD.DigitGrouping(row.Total_Qty.ToString, row._Decimal) + " " + row.Unit_Code + " " + row.Secondary_Qty.ToString + row.Sec_Unit_Code
                Else
                    row.Qty_Str = Me.objCD.DigitGrouping(row.Total_Qty, row._Decimal) + " " + row.Unit_Code
                End If
            Next

            'urutkan berdasarkan Brand, AOP Number, Code Number
            Sort = "Brand,[AOP Number], [Product Code]"
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = Sort

            For Each rowView In TBLDummy.DefaultView
                row = rowView.Row
                row1 = DS.DataTable1.NewDataTable1Row

                With row
                    If Me.CheckBox1.Checked = True Then
                        If .Supplier_Invoice_No <> "" Then
                            row1.AOP_Number = .Supplier_Invoice_No
                        Else
                            row1.AOP_Number = .AOP_Number
                        End If
                    Else
                        row1.AOP_Number = .AOP_Number
                    End If
                    row1.Brand = .Brand
                    row1.Primary_Brand = .Primary_Brand
                    row1.Primary_GRN = .Primary_GRN
                    row1.Primary_PO = .Primary_PO
                    row1.Primary_Product = .Primary_Product
                    row1.Product_Code = .Product_Code
                    row1.Qty_Str = .Qty_Str
                    row1.Sec_Unit_Code = .Sec_Unit_Code
                    row1.Secondary_Qty = .Secondary_Qty
                    row1.Tanggal_Transaksi = .Tanggal_Transaksi
                    row1.Total_Qty = .Total_Qty
                    row1.Total_Unit_Price = .Total_Unit_Price
                    row1.Unit_Code = .Unit_Code
                    row1.Unit_Price = .Unit_Price
                    row1.Supplier_Invoice_No = .Supplier_Invoice_No
                    row1._Decimal = row._Decimal
                    row1.At_Amount = row.At_Amount

                    If BrandDummy <> row1.Brand Then
                        row1.Show_AOP_No = 1
                        AOPDummy = row1.AOP_Number
                        BrandDummy = row1.Brand
                    Else
                        If AOPDummy <> row1.AOP_Number Then
                            row1.Show_AOP_No = 1
                            AOPDummy = row1.AOP_Number
                        Else
                            row1.Show_AOP_No = 0
                        End If
                    End If
                End With

                DS.DataTable1.Rows.Add(row1)
            Next

            ' Me.Checking(DS.DataTable1)

            'tandai mana saja yang muncul tanggalnya, dan mana saja yang tidak muncul tanggalnya
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
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))
            If Me.CheckBox1.Checked = True Then
                report.SetParameterValue(report.Parameter_RefDocument.ParameterFieldName, "Supplier Invoice")
            Else
                report.SetParameterValue(report.Parameter_RefDocument.ParameterFieldName, "AOP")
            End If

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
        End Try
    End Sub

    Private Sub Generate_revisi_invoice_baru()
        If Me.ApakahBolehCetak = False Then
            Exit Sub
        End If

        Dim DS As New DseLaporanPurchaseAuditList
        Dim DTA As New DseLaporanPurchaseAuditListTableAdapters.DataTable1TableAdapter
        Dim report As New CrpLaporanPurchaseCost
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row, row1 As DseLaporanPurchaseAuditList.DataTable1Row
        Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = 0

        Dim DTAReturn As New DseLaporanPurchaseAuditListTableAdapters.DataTable2TableAdapter
        Dim TBLReturn As New DseLaporanPurchaseAuditList.DataTable2DataTable
        Dim rowReturn As DseLaporanPurchaseAuditList.DataTable2Row
        Dim Filter As String = ""
        Dim TBLDummy As New DseLaporanPurchaseAuditList.DataTable1DataTable
        Dim Sort As String = ""
        Dim rowView As DataRowView

        Dim AOPDummy As String = ""
        Dim BrandDummy As String = ""
        Dim i As Double = 0
        Dim TipeProduct As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            'in purchase
            For Each DV In Me.clbBrand.CheckedItems
                primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                DTA.ClearBeforeFill = False

                If Me.chkGoods.Checked = True And Me.chkService.Checked = True Then
                    DTA.FillBy1(TBLDummy, _
                    TglSampai, TglDari, primaryBrand)
                Else
                    If Me.chkGoods.Checked = True Then
                        TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
                    Else
                        TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
                    End If
                    DTA.FillBy3(TBLDummy, _
                    TglSampai, TglDari, primaryBrand, CInt(TipeProduct))
                End If
            Next

            ''@rosadinnn
            ''@MUaT
            Dim ObjProduct As New Core_Module.clsMasterProduct
            For Each row In TBLDummy
                With row
                    If .Primary_Units <> .Primary_Units_Multi Then
                        .Decimal_MultiUnit = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                    Else
                        .Decimal_MultiUnit = ._Decimal
                    End If
                End With
            Next

            'kurangkan dengan data return 
            DTAReturn.Fill(TBLReturn, TglSampai, TglDari)
            For Each rowReturn In TBLReturn.Rows
                Filter = "[Primary Product] =" + rowReturn.Primary_Product.ToString + _
                " AND [AOP Number] = '" + rowReturn.AOP_Number + "' AND [Primary GRN] = " + _
                rowReturn.Primary_GRN.ToString + " AND [Primary PO] = " + rowReturn.Primary_PO.ToString + _
                " AND [Secondary Qty] = " + rowReturn.Secondary_Qty.ToString
                TBLDummy.DefaultView.RowFilter = Filter

                i = 0
                While TBLDummy.DefaultView.Count > i And rowReturn.Total_Qty > 0
                    row = TBLDummy.DefaultView(i).Row
                    If row.Total_Qty >= rowReturn.Total_Qty Then
                        row.Total_Qty -= rowReturn.Total_Qty
                        rowReturn.Total_Qty = 0
                    Else
                        rowReturn.Total_Qty -= row.Total_Qty
                        row.Total_Qty = 0
                    End If

                    If row.Total_Qty = 0 Then
                        TBLDummy.Rows.Remove(row)
                    End If
                End While
            Next


            'betulan data qty str dan hitung total unit price
            For Each row In TBLDummy.Rows
                row.Total_Unit_Price = row.Unit_Price * row.Total_Qty
                row.Total_Unit_Price = row.At_Amount
                If row.Sec_Unit_Code <> "" Then
                    row.Qty_Str = Me.objCD.DigitGrouping(row.Total_Qty.ToString, row._Decimal) + " " + _
                    row.Unit_Code + " " + row.Secondary_Qty.ToString + row.Sec_Unit_Code
                Else

                    ''@rosadinnn
                    ''@MUaT
                    row.Qty_Str = Me.objCD.DigitGrouping(row.Total_Qty, row.Decimal_MultiUnit) + " " + row.Unit_Code
                    If Me.ChkShowRatio.Checked = True And row.Primary_Units <> row.Primary_Units_Multi Then
                        row.Product_Code += "(" & objCD.DigitGrouping(row.Ratio_Reference, row.Decimal_MultiUnit) & " " & _
                        row.Unit_Code & "=" & objCD.DigitGrouping(row.Ratio_Default, row._Decimal) & " " & row.Unit_Code_Default & ")"
                    End If

                End If
            Next

            'urutkan berdasarkan Brand, AOP Number, Code Number
            Sort = "Brand,[AOP Number], [Product Code]"
            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = Sort

            For Each rowView In TBLDummy.DefaultView
                row = rowView.Row
                row1 = DS.DataTable1.NewDataTable1Row

                With row
                    If Me.CheckBox1.Checked = True Then
                        If .Supplier_Invoice_No <> "" Then
                            row1.AOP_Number = .Supplier_Invoice_No
                        Else
                            row1.AOP_Number = .AOP_Number
                        End If
                    Else
                        row1.AOP_Number = .AOP_Number
                    End If
                    row1.Brand = .Brand
                    row1.Primary_Brand = .Primary_Brand
                    row1.Primary_GRN = .Primary_GRN
                    row1.Primary_PO = .Primary_PO
                    row1.Primary_Product = .Primary_Product
                    row1.Product_Code = .Product_Code
                    row1.Qty_Str = .Qty_Str
                    row1.Sec_Unit_Code = .Sec_Unit_Code
                    row1.Secondary_Qty = .Secondary_Qty
                    row1.Tanggal_Transaksi = .Tanggal_Transaksi
                    row1.Total_Qty = .Total_Qty
                    row1.Total_Unit_Price = .Total_Unit_Price
                    row1.Unit_Code = .Unit_Code
                    row1.At_Amount = row.At_Amount
                    If Me.dtpFrom.Value = #2/1/2014# Or Me.dtpFrom.Value = #3/1/2014# Then
                        row1.Unit_Price = Math.Round(row1.At_Amount / row1.Total_Qty, _
                        DefPrec, MidpointRounding.AwayFromZero)
                    Else
                        row1.Unit_Price = .Unit_Price
                    End If
                    row1.Supplier_Invoice_No = .Supplier_Invoice_No
                    row1._Decimal = row._Decimal

                    If BrandDummy <> row1.Brand Then
                        row1.Show_AOP_No = 1
                        AOPDummy = row1.AOP_Number
                        BrandDummy = row1.Brand
                    Else
                        If AOPDummy <> row1.AOP_Number Then
                            row1.Show_AOP_No = 1
                            AOPDummy = row1.AOP_Number
                        Else
                            row1.Show_AOP_No = 0
                        End If
                    End If
                End With

                DS.DataTable1.Rows.Add(row1)
            Next

            ' Me.Checking(DS.DataTable1)

            'tandai mana saja yang muncul tanggalnya, dan mana saja yang tidak muncul tanggalnya
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
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_Periode.ParameterFieldName, Format(Me.dtpFrom.Value, "MMMMM yyyy"))
            If Me.CheckBox1.Checked = True Then
                report.SetParameterValue(report.Parameter_RefDocument.ParameterFieldName, "Supplier Invoice")
            Else
                report.SetParameterValue(report.Parameter_RefDocument.ParameterFieldName, "AOP")
            End If

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
        End Try
    End Sub

    Private Sub FrmLaporanSalesReturnList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.dtpFrom.Value = Me.objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
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

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Function GetAmount(ByVal vPrimaryProduct As Double, ByVal vTBLDummy As DseLaporanPurchaseAuditList.DataTable1DataTable) As Decimal
        Dim row1 As DseLaporanPurchaseAuditList.DataTable1Row
        Dim retval As Decimal = 0

        Try
            For Each row1 In vTBLDummy.Rows
                If row1.Primary_Product = vPrimaryProduct Then
                    retval += row1.Total_Unit_Price
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetAmount(ByVal vPrimaryProduct As Double, ByVal vNoTransaksi As String, ByVal vPrimaryGRN As Double, _
    ByVal vPrimaryPO As Double, ByVal TBL As Core_Module.DseLaporanPembelian.tbltrans_aop_view2DataTable) As Decimal
        Try
            Dim row As Core_Module.DseLaporanPembelian.tbltrans_aop_view2Row
            Dim Retval As Double = 0

            For Each row In TBL.Rows
                If row.Primary_Product = vPrimaryProduct And row.No_Transaksi = vNoTransaksi And _
                row.Primary_GRN = vPrimaryGRN And row.Primary_PO = vPrimaryPO Then
                    Retval += row.Amount
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function
End Class