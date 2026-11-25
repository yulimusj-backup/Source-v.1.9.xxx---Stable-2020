Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesOrderReportByDateDetail
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim DS As New DseSalesOrderReportByDateDetail
    Dim report As New CrpSalesOrderReportByDateDetail
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjSales As New Core_Module.ClsMasterSalesPerson
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1 '
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim DummyIndex As Integer = -1
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Buffer As String = ""

    Private Function GetMaxMin(ByRef PrimaryMin As Double, ByRef PrimaryMax As Double) As Boolean
        PrimaryMin = 0
        PrimaryMax = 0
        Try
            Me.dgvItemProductGroup.Sort(Me.ColProductGroup_PrimaryAcuan, System.ComponentModel.ListSortDirection.Ascending)
            For Each row As DataGridViewRow In Me.dgvItemProductGroup.Rows
                If String.IsNullOrEmpty(row.Cells(Me.ColProductGroup_CodeNo.Index).Value) = False Then
                    If row.Index = 0 Then
                        PrimaryMin = row.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value
                    Else
                        If CDbl(row.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value) > PrimaryMax Then
                            PrimaryMax = row.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value
                        End If
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
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
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.dtpFrom.Focus()
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.CetakDetail(Me.dtpTo.Value, Me.dtpFrom.Value, CDbl(Me.txtPrimaryProduct.Text), CDbl(Me.txtPrimaryCustomer.Text), _
            CDbl(Me.txtPrimaryCurrency.Text), CDbl(Me.txtPrimarySalesPerson.Text))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDecimalMulti(ByVal TBLCetakDummy As DseSalesOrderReportByDateDetail.tblCetakDataTable, ByVal vPrimaryProduct As Double)
        Dim RowCetakDummy As DseSalesOrderReportByDateDetail.tblCetakRow
        Dim DTA As New DseSalesOrderReportByDateDetailTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim TBL As New DseSalesOrderReportByDateDetail.tblmaster_product_item_multiunitDataTable
        Dim row As DseSalesOrderReportByDateDetail.tblmaster_product_item_multiunitRow

        Try
            DTA.FillBy_Primary_Product(TBL, vPrimaryProduct)
            For Each RowCetakDummy In TBLCetakDummy.Rows
                If RowCetakDummy.Primary_Product = vPrimaryProduct Then
                    If RowCetakDummy.Primary_Units <> RowCetakDummy.Primary_Units_Multi Then
                        For Each row In TBL.Rows
                            If RowCetakDummy.Primary_Units_Multi = row.Primary_Unit Then
                                RowCetakDummy.Decimal_Multi = row.Default_Precision
                                RowCetakDummy.Unit_Code = row.Unit_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakDetail(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimaryCustomer As Double, ByVal PrimaryCurrency As Double, ByVal PrimarySalesPerson As Double)
        Dim Report As New CrpSalesOrderReportByDateDetail
        Dim DS As New DseSalesOrderReportByDateDetail

        Dim DTACetakDummy As New DseSalesOrderReportByDateDetailTableAdapters.tblCetakTableAdapter
        Dim TBLCetak As DseSalesOrderReportByDateDetail.tblCetakDataTable = DS.tblCetak
        Dim TBLCetakDummy As New DseSalesOrderReportByDateDetail.tblCetakDataTable
        Dim TBLTotalCurr As DseSalesOrderReportByDateDetail.tblTotalCurrDataTable = DS.tblTotalCurr
        Dim RowCetak As DseSalesOrderReportByDateDetail.tblCetakRow
        Dim RowCetakDummy As DseSalesOrderReportByDateDetail.tblCetakRow
        Dim RowTotalCurr As DseSalesOrderReportByDateDetail.tblTotalCurrRow
        Dim DV As DataRowView
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView

        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim filter As String = "" 

        Dim DTADel As New DseSalesOrderReportByDateDetailTableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBLDel As New DseSalesOrderReportByDateDetail.tbltrans_so_itemdeliveryDataTable
        Dim rowDel As DseSalesOrderReportByDateDetail.tbltrans_so_itemdeliveryRow
        Dim PrimaryDummy As Double = -1
        Dim JmlDel As Decimal = 0
        Dim Ketemu As Boolean = False
        Dim PrimaryMin As Double = -1
        Dim PrimaryMax As Double = -1
        Dim FilterGroup As String = ""
        Dim vPrimarySO As String = ""
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Try
            DTACetakDummy.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan tabel
            TBLCetak.Rows.Clear()
            TBLCetakDummy.Rows.Clear()

            paramStartDate = Me.dtpFrom.Value
            paramEndDate = Me.dtpTo.Value

            'Isi tblCetakDummy ----------------------------------------------------------------------------------------------------------------
            'If Me.ChkAllProductSO.Checked = True Then
            '    'If Me.ChkOS.Checked = True Then
            '    '    DTACetakDummy.FillBy_OS_AllProductSO(TBLCetakDummy, Me.dtpFrom.Value, Me.dtpTo.Value, PrimaryProduct)
            '    'Else

            '    DTACetakDummy.FillBy_AllProductSO(TBLCetakDummy, Me.dtpFrom.Value, Me.dtpTo.Value)
            '    'End If
            'Else
            '    'If Me.ChkOS.Checked = True Then
            '    '    DTACetakDummy.FillBy_OS(TBLCetakDummy, Me.dtpFrom.Value, Me.dtpTo.Value)
            '    'Else
            '    DTACetakDummy.Fill(TBLCetakDummy, Me.dtpFrom.Value, Me.dtpTo.Value)
            '    'End If
            'End If

            DTACetakDummy.Fill(TBLCetakDummy, Me.dtpFrom.Value, Me.dtpTo.Value)
            'Filter tblCetakDummy disini ------------------------------------------------------------------------------------------------------
            If Me.ChkOS.Checked = True Then
                filter = "[Primary SO] = [Primary SO] AND [OS] > 0"
            Else
                filter = "[Primary SO] = [Primary SO]"
            End If

            If PrimaryProduct <> -1 Then
                If Me.ChkAllProductSO.Checked = False Then
                    filter &= " AND [Primary Product] = " & PrimaryProduct & ""
                Else
                    filter = filter
                End If
            Else
                filter = filter
            End If
            If PrimaryCustomer <> -1 Then
                filter &= " AND [Primary Customer] = " & PrimaryCustomer & ""
            End If
            If PrimarySalesPerson <> -1 Then
                filter &= " AND [Primary Sales Person] = " & PrimarySalesPerson & ""
            End If
            If PrimaryCurrency <> -1 Then
                filter &= " AND [Primary Currency] = " & PrimaryCurrency & ""
            End If
            FilterGroup = Me.GetFilterProduct(TBLCetakDummy)
            If FilterGroup <> "" Then
                filter &= " AND " & FilterGroup
            End If

            TBLCetakDummy.DefaultView.RowFilter = Nothing
            TBLCetakDummy.DefaultView.RowFilter = filter
            If TBLCetakDummy.DefaultView.Count > 0 Then
                For Each DV In TBLCetakDummy.DefaultView
                    RowCetakDummy = DV.Row
                   
                    RowCetakDummy.Delivered = RowCetakDummy.Qty_Multi - RowCetakDummy.SOB_Multi
                    RowCetakDummy.Persen = (RowCetakDummy.Delivered / RowCetakDummy.Qty_Multi) * 100
                    If RowCetakDummy.Primary_Units <> RowCetakDummy.Primary_Units_Multi Then
                        If RowCetakDummy.Decimal_Multi = -1 Then
                            Me.IsikanDecimalMulti(TBLCetakDummy, RowCetakDummy.Primary_Product)
                        End If
                    Else
                        RowCetakDummy.Decimal_Multi = RowCetakDummy._Decimal
                    End If

                    'isi tblCetak dengan tblCetakDummy hasil filteran disini --------------------------------------------------------------------------
                    RowCetak = TBLCetak.NewtblCetakRow
                    RowCetak.Primary_SO = RowCetakDummy.Primary_SO
                    RowCetak.Primary_Product = RowCetakDummy.Primary_Product
                    RowCetak.Primary_Customer = RowCetakDummy.Primary_Customer
                    RowCetak.Primary_Units = RowCetakDummy.Primary_Units
                    RowCetak.primary_sales_person = RowCetakDummy.primary_sales_person
                    RowCetak.SO_No = RowCetakDummy.SO_No
                    RowCetak.SO_Date = RowCetakDummy.SO_Date

                    If RowCetakDummy.Ratio_Default <> RowCetakDummy.Ratio_Reference Then
                        RowCetak.Product_Code = RowCetakDummy.Product_Code & " (" & Me.ObjCD.DigitGrouping(RowCetakDummy.Ratio_Reference, RowCetakDummy.Decimal_Multi) & " " & ObjUnit.GetUnitCode(RowCetakDummy.Primary_Units_Multi) & "=" & _
                        Me.ObjCD.DigitGrouping(RowCetakDummy.Ratio_Default, RowCetakDummy._Decimal) & " " & ObjUnit.GetUnitCode(RowCetakDummy.Primary_Units) & ")"
                    Else
                        RowCetak.Product_Code = RowCetakDummy.Product_Code
                    End If
                    RowCetak.Product_Name = RowCetakDummy.Product_Name
                    RowCetak.Customer_Code = RowCetakDummy.Customer_Code
                    RowCetak.Customer_Name = RowCetakDummy.Customer_Name
                    RowCetak.Sales_Person_Code = RowCetakDummy.Sales_Person_Code
                    RowCetak.Sales_Person_Name = RowCetakDummy.Sales_Person_Name
                    RowCetak.Cust_PO_No = RowCetakDummy.Cust_PO_No
                    RowCetak.identifier = RowCetakDummy.identifier
                    RowCetak.Delivery_Date_Cont = RowCetakDummy.Delivery_Date_Cont

                    If IsNumeric(RowCetakDummy.Payment_Term) Then
                        RowCetak.Payment_Term = RowCetakDummy.Payment_Term & " DAYS"
                    Else
                        RowCetak.Payment_Term = RowCetakDummy.Payment_Term
                    End If
                    RowCetak.Unit_Code = RowCetakDummy.Unit_Code
                    RowCetak.Ordered = RowCetakDummy.Qty_Multi
                    RowCetak.Delivered = RowCetakDummy.Delivered
                    RowCetak.OS = RowCetakDummy.SOB_Multi
                    RowCetak.Persen = RowCetakDummy.Persen
                    RowCetak.Primary_Currency = RowCetakDummy.Primary_Currency
                    RowCetak.Currency_Code = RowCetakDummy.Currency_Code
                    RowCetak.Nett_Price = RowCetakDummy.Nett_Price
                    RowCetak.Amount = RowCetakDummy.Amount
                    RowCetak.Tax = RowCetakDummy.Tax
                    RowCetak.Default_Precision = RowCetakDummy.Default_Precision
                    RowCetak._Decimal = RowCetakDummy.Decimal_Multi
                    RowCetak.Decimal_Multi = RowCetak._Decimal

                    RowCetak.Sub_Total = RowCetakDummy.Sub_Total
                    RowCetak.VAT = RowCetakDummy.VAT
                    RowCetak.Total = RowCetakDummy.Total

                    TBLCetak.Rows.Add(RowCetak)
                Next
            End If

            For Each RowCetak In TBLCetak.Rows
                'Dim DecSubTotal As Decimal = 0
                TBLCetakDummy.DefaultView.RowFilter = Nothing
                If PrimaryProduct <> -1 Then
                    If Me.ChkAllProductSO.Checked = False Then
                        If Me.ChkOS.Checked = True Then
                            TBLCetakDummy.DefaultView.RowFilter = "[Primary So] = " & RowCetak.Primary_SO & " AND [Primary Product] = " & PrimaryProduct & " AND [OS] > 0"
                        Else
                            TBLCetakDummy.DefaultView.RowFilter = "[Primary So] = " & RowCetak.Primary_SO & " AND [Primary Product] = " & PrimaryProduct & ""
                        End If
                    Else
                        If Me.ChkOS.Checked = True Then
                            TBLCetakDummy.DefaultView.RowFilter = "[Primary So] = " & RowCetak.Primary_SO & " AND [OS] > 0"
                        Else
                            TBLCetakDummy.DefaultView.RowFilter = "[Primary So] = " & RowCetak.Primary_SO
                        End If
                    End If
                Else
                    If Me.ChkOS.Checked = True Then
                        TBLCetakDummy.DefaultView.RowFilter = "[Primary So] = " & RowCetak.Primary_SO & " AND [OS] > 0"
                    Else
                        TBLCetakDummy.DefaultView.RowFilter = "[Primary So] = " & RowCetak.Primary_SO
                    End If
                End If
                For Each DV1 In TBLCetakDummy.DefaultView
                    RowCetakDummy = DV1.Row
                    If IsNumeric(RowCetakDummy.Amount) Then
                        RowCetak.Sub_Total += CDec(RowCetakDummy.Amount)
                        RowCetak.VAT += RowCetakDummy.Amount * (CDec(RowCetakDummy.Tax) / 100)
                    End If
                Next
                'RowCetak.SubTotal = DecSubTotal
                RowCetak.Total = RowCetak.Sub_Total + RowCetak.VAT
            Next

            For Each RowCetak In TBLCetak.Rows
                TBLTotalCurr.DefaultView.RowFilter = Nothing
                TBLTotalCurr.DefaultView.RowFilter = "[Primary Currency ] = " & RowCetak.Primary_Currency
                If TBLTotalCurr.DefaultView.Count > 0 Then
                    For Each DV2 In TBLTotalCurr.DefaultView
                        RowTotalCurr = DV2.Row
                        If IsNumeric(RowCetak.Amount) Then
                            RowTotalCurr.SubTotal += CDec(RowCetak.Amount)
                            RowTotalCurr.VAT += RowCetak.Amount * (CDec(RowCetak.Tax) / 100)
                            RowTotalCurr.Total = RowTotalCurr.SubTotal + RowTotalCurr.VAT
                        End If
                    Next
                Else
                    RowTotalCurr = TBLTotalCurr.NewtblTotalCurrRow
                    RowTotalCurr.Primary_Currency_ = RowCetak.Primary_Currency
                    RowTotalCurr.Currency_Code = RowCetak.Currency_Code
                    RowTotalCurr.Default_Precision = RowCetak.Default_Precision
                    RowTotalCurr.SubTotal += CDec(RowCetak.Amount)
                    RowTotalCurr.VAT += RowCetak.Amount * (CDec(RowCetak.Tax) / 100)
                    RowTotalCurr.Total = RowTotalCurr.SubTotal + RowTotalCurr.VAT
                    TBLTotalCurr.Rows.Add(RowTotalCurr)
                End If
            Next

            'isikan tanggal delivery
            If TBLCetak.Rows.Count > 0 Then 
                DTADel.Connection.ConnectionString = Core_Module.GetConstring
                DTADel.Fill(TBLDel, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
                TBLDel.DefaultView.RowFilter = ""

                For Each RowCetak In TBLCetak.Rows
                    If RowCetak.Delivery_Date_Cont = True Then
                        TBLDel.DefaultView.RowFilter = "[primary transaksi] = " + RowCetak.Primary_SO.ToString + " AND identifier = '" + RowCetak.identifier + "'"
                        If TBLDel.DefaultView.Count = 0 Then
                            RowCetak.Delivery_Date = "-"
                        ElseIf TBLDel.DefaultView.Count = 1 Then
                            rowDel = TBLDel.DefaultView(0).Row
                            RowCetak.Delivery_Date = Format(rowDel.delivery_date, "dd MMM yyyy")
                            While TBLDel.DefaultView.Count > 0
                                TBLDel.Rows.Remove(TBLDel.DefaultView(0).Row)
                            End While
                        Else
                            JmlDel = 0
                            Ketemu = False
                            For iw As Integer = TBLDel.DefaultView.Count - 1 To 0 Step -1
                                rowDel = TBLDel.DefaultView(iw).Row
                                JmlDel += rowDel.delivery_qty_multi
                                If JmlDel >= RowCetak.OS Then
                                    RowCetak.Delivery_Date = Format(rowDel.delivery_date, "dd MMM yyyy")
                                    Ketemu = True
                                    Exit For
                                End If
                            Next
                            If Ketemu = False Then
                                RowCetak.Delivery_Date = Format(rowDel.delivery_date, "dd MMM yyyy")
                            End If
                            While TBLDel.DefaultView.Count > 0
                                TBLDel.Rows.Remove(TBLDel.DefaultView(0).Row)
                            End While
                        End If
                    Else
                        RowCetak.Delivery_Date = "-"
                    End If
                Next
            End If

            'generate report
            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_StartDate.ParameterFieldName, paramStartDate)
            Report.SetParameterValue(Report.Parameter_EndDate.ParameterFieldName, paramEndDate)
            crv.ReportSource = Report
            crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTACetakDummy.Dispose()
            TBLCetak.Dispose()
            TBLCetakDummy.Dispose()
            TBLTotalCurr.Dispose()
            DS.Dispose()
            DTADel.Dispose()
            TBLDel.Dispose()
        End Try
    End Sub

    Private Sub txtProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtProductCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtProductCode.Text <> "" Then
                    If Me.ObjProduct.ApakahSudahAda(Me.txtProductCode.Text) = True Then
                        Me.txtPrimaryProduct.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtProductCode.Text))
                        Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtCustomerCode.Focus()
                    Else
                        Me.btnFindProduct.PerformClick()
                    End If
                Else
                    Me.btnFindProduct.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.Leave
        Dim hj As New Core_Module.frmCariMasterProduct
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtProductCode.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.txtProductCode.Text) = True Then
                    Me.txtPrimaryProduct.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtProductCode.Text))
                    Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                    Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                Else
                    If Me.txtProductCode.Text <> "" Then
                        hj.txtCari.Text = Me.txtProductCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryProduct > 0 Then
                        Me.txtPrimaryProduct.Text = CStr(hj.PrimaryProduct)
                        Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                        Me.txtCustomerCode.Focus()
                    Else
                        Me.txtPrimaryProduct.Text = "-1"
                        Me.txtProductName.Clear()
                        Me.txtProductCode.Clear()
                        Me.txtProductCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtProductName.Clear()
                Me.txtProductCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtProductCode.Text <> "" Then
                hj.txtCari.Text = Me.txtProductCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryProduct > 0 Then
                Me.txtPrimaryProduct.Text = CStr(hj.PrimaryProduct)
                Me.txtProductCode.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryProduct.Text))
                Me.txtProductName.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryProduct.Text))
                Me.txtCustomerCode.Focus()
            Else
                Me.txtPrimaryProduct.Text = "-1"
                Me.txtProductName.Clear()
                Me.txtProductCode.Clear()
                Me.txtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtCustomerCode.Text <> "" Then
                    If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                        Me.txtPrimaryCustomer.Text = CStr(Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text))
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCustomerCode.Focus()
                    Else
                        Me.btnFindCustomer.PerformClick()
                    End If
                Else
                    Me.btnFindCustomer.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Dim hj As New Core_Module.FrmCariMasterCustomer
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = CStr(Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text))
                    Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                    Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        hj.txtCari.Text = Me.txtCustomerCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = CStr(hj.PrimaryCustomer)
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCustomerCode.Focus()
                    Else
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtCustomerName.Clear()
                        Me.txtCustomerCode.Clear()
                        Me.txtCustomerCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtCustomerName.Clear()
                Me.txtCustomerCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim hj As New Core_Module.FrmCariMasterCustomer
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtCustomerCode.Text <> "" Then
                hj.txtCari.Text = Me.txtCustomerCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryCustomer > 0 Then
                Me.txtPrimaryCustomer.Text = CStr(hj.PrimaryCustomer)
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                Me.txtCustomerCode.Focus()
            Else
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtCustomerName.Clear()
                Me.txtCustomerCode.Clear()
                Me.txtCustomerCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtCurrencyCode.Text <> "" Then
                    If Me.ObjCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                        Me.txtPrimaryCurrency.Text = CStr(Me.ObjCurr.GetPrimary(Me.txtCurrencyCode.Text))
                        Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                        Me.btnFilter.Focus()
                    Else
                        Me.btnFindCurrency.PerformClick()
                    End If
                Else
                    Me.btnFindCurrency.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim hj As New Core_Module.FrmCariMasterCurrency
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If Me.ObjCurr.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = CStr(Me.ObjCurr.GetPrimary(Me.txtCurrencyCode.Text))
                    Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        hj.txtCari.Text = Me.txtCurrencyCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = CStr(hj.PrimaryCurrency)
                        Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                        Me.btnFilter.Focus()
                    Else
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Clear()
                        Me.txtCurrencyCode.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim hj As New Core_Module.FrmCariMasterCurrency
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtCurrencyCode.Text <> "" Then
                hj.txtCari.Text = Me.txtCurrencyCode.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryCurrency > 0 Then
                Me.txtPrimaryCurrency.Text = CStr(hj.PrimaryCurrency)
                Me.txtCurrencyCode.Text = Me.ObjCurr.GetCurrencyCode(CDbl(Me.txtPrimaryCurrency.Text))
                Me.btnFilter.Focus()
            Else
                Me.txtPrimaryCurrency.Text = "-1"
                Me.txtCurrencyCode.Clear()
                Me.txtCurrencyCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtSalesPerson_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSalesPerson.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtSalesPerson.Text <> "" Then
                    If Me.ObjSales.ApakahSudahAda(Me.txtSalesPerson.Text) = True Then
                        Me.txtPrimarySalesPerson.Text = CStr(Me.ObjSales.GetPrimary(Me.txtSalesPerson.Text))
                        Me.txtSalesPerson.Text = Me.ObjSales.GetCode(CDbl(Me.txtPrimarySalesPerson.Text))
                        Me.TxtSalesPersonName.Text = Me.ObjSales.GetName(CDbl(Me.txtPrimarySalesPerson.Text))
                        Me.txtCurrencyCode.Focus()
                    Else
                        Me.btnFindSalesPerson.PerformClick()
                    End If
                Else
                    Me.btnFindSalesPerson.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSalesPerson_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesPerson.Leave
        Dim hj As New Core_Module.FrmMasterSalesPerson_Cari
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtSalesPerson.Text <> "" Then
                If Me.ObjSales.ApakahSudahAda(Me.txtSalesPerson.Text) = True Then
                    Me.txtPrimarySalesPerson.Text = CStr(Me.ObjSales.GetPrimary(Me.txtSalesPerson.Text))
                    Me.txtSalesPerson.Text = Me.ObjSales.GetCode(CDbl(Me.txtPrimarySalesPerson.Text))
                    Me.TxtSalesPersonName.Text = Me.ObjSales.GetName(CDbl(Me.txtPrimarySalesPerson.Text))
                Else
                    If Me.txtSalesPerson.Text <> "" Then
                        hj.txtCari.Text = Me.txtSalesPerson.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimarySalesPerson > 0 Then
                        Me.txtPrimarySalesPerson.Text = CStr(hj.PrimarySalesPerson)
                        Me.txtSalesPerson.Text = Me.ObjSales.GetCode(CDbl(Me.txtPrimarySalesPerson.Text))
                        Me.TxtSalesPersonName.Text = Me.ObjSales.GetName(CDbl(Me.txtPrimarySalesPerson.Text))
                        Me.txtCurrencyCode.Focus()
                    Else
                        Me.txtPrimarySalesPerson.Text = "-1"
                        Me.TxtSalesPersonName.Clear()
                        Me.txtSalesPerson.Clear()
                        Me.txtSalesPerson.Focus()
                    End If
                End If
            Else
                Me.txtPrimarySalesPerson.Text = "-1"
                Me.TxtSalesPersonName.Clear()
                Me.txtSalesPerson.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesPerson.Click
        Dim hj As New Core_Module.FrmMasterSalesPerson_Cari
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtSalesPerson.Text <> "" Then
                hj.txtCari.Text = Me.txtSalesPerson.Text
            End If

            hj.ShowDialog()
            If hj.PrimarySalesPerson > 0 Then
                Me.txtPrimarySalesPerson.Text = CStr(hj.PrimarySalesPerson)
                Me.txtSalesPerson.Text = Me.ObjSales.GetCode(CDbl(Me.txtPrimarySalesPerson.Text))
                Me.TxtSalesPersonName.Text = Me.ObjSales.GetName(CDbl(Me.txtPrimarySalesPerson.Text))
                Me.btnFilter.Focus()
            Else
                Me.txtPrimarySalesPerson.Text = "-1"
                Me.TxtSalesPersonName.Clear()
                Me.txtSalesPerson.Clear()
                Me.txtSalesPerson.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.GotFocus, txtCurrencyCode.GotFocus, txtCustomerCode.GotFocus, txtSalesPerson.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtProductName.BackColor
                        Else
                            Txt.BackColor = Me.BAckColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.LostFocus, txtCustomerCode.LostFocus, txtCurrencyCode.LostFocus, txtSalesPerson.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnFilter.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtProductName.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimaryProduct.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryProduct_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrimaryProduct.TextChanged
        Try
            If Me.txtPrimaryProduct.Text = -1 Then
                Me.ChkAllProductSO.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProductSO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProductSO.Click
        Try
            'If Me.txtPrimaryProduct.Text <> -1 Then

            'Else
            '    Me.ChkAllProductSO.Checked = False
            '    MsgBox("Product cannot be an empty string", MsgBoxStyle.Information, "Information")
            '    Me.txtProductCode.Focus()
            'End If
            If Me.ApakahPakeFilterProduct = False Then
                Me.ChkAllProductSO.Checked = False
                MsgBox("Product cannot be an empty string", MsgBoxStyle.Information, "Information")
                Me.txtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSalesOrderReportByDateDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ChkOS.Text = "Only show SO with sales orders balance"
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "Filteran Datagrid Product"
    Private Sub dgvItemProductGroup_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemProductGroup.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            'yang bisa diedit:
            '1. jika baris baru
            If e.RowIndex = Me.dgvItemProductGroup.Rows.Count - 1 Then
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'pastikan baris diatasnya sudah beres dahulu
                    If e.RowIndex > 0 Then
                        rowDGV = Me.dgvItemProductGroup.Rows(e.RowIndex - 1)
                        If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value.ToString.Trim = "" Then
                            MsgBox("Please specify " + rowDGV.Cells(Me.ColProductGroup_Level.Index).Value.ToString + _
                            " code on the previous row.", MsgBoxStyle.Information, "Information")
                            Me.dgvItemProductGroup.ClearSelection()
                            rowDGV.Selected = True
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.Buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemProductGroup.CellEndEdit
        Dim Level As Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Try
            If IsNothing(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True _
                       Or Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                If e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'Remarks
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = ""
                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    'kalo kode product group dihapus, maka hapus satu bari semuanya
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                    Exit Sub
                End If
            Else
                'level
                If e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                    'do nothing

                ElseIf e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Then
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If

                    'cek ke master
                    Level = Me.objItem.GetProductLevel(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
                    If Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Me.EndEditProductGroup(e, Me.ProductGroup_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Me.EndEditProductGroup(e, Me.Brand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Me.EndEditProductGroup(e, Me.SubBrand_Level)
                    ElseIf Level = Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Me.EndEditCategory(e)
                    Else
                        Me.EndEditProduct(e)
                    End If
                ElseIf e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                    'remarks
                    'pastikan ditrim dan  max 250 karakter
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 250 Then
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 250)
                    End If
                    'tidak boleh ada karakter aneh2
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("'") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("%") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("*") Or _
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains("#") Then
                        MsgBox("Please avoid using the following character: ', *, %, #.", MsgBoxStyle.Information, "Information")
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndex_ItemProductGroup(ByVal vCode As String, ByVal vRefIndex As Integer, ByVal vProductGroupLevelStr As String) As Integer
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = vCode And _
                rowDGV.Index <> vRefIndex And _
                rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = vProductGroupLevelStr Then
                    Return rowDGV.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub EndEditProductGroup(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal vProductGroup_Level As Integer)
        Try
            If Me.ObjProductGroup.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level) = True Then
                '1. kode product group ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, vProductGroup_Level)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product group yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.FrmCariMasterProductGroup
                ij.ProductGroupLevel = vProductGroup_Level
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProductGroup <> -1 Then
                    '2.a. product group yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProductGroup
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProductGroup.GetGroupCode(ij.PrimaryProductGroup)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(ij.PrimaryProductGroup)
                ElseIf ij.PrimaryProductGroup = -1 Then
                    '2.b. kode akun yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Trim(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If
            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.Buffer, vProductGroup_Level)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProductGroup.GetGroupDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCategory(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjCategory.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode category ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode category yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterCategory
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryCategory <> -1 Then
                    '2.a. category yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryCategory
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCategory.GetCategoryCode(ij.PrimaryCategory)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(ij.PrimaryCategory)
                ElseIf ij.PrimaryCategory = -1 Then
                    '2.b. kode category yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjCategory.GetDesc(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditProduct(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If Me.ObjProduct.ApakahSudahAda(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                '1. kode product ada  --> load data nama + primary
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value)
                Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
            Else
                '2. kode product yang diketik tidak ada --> tampilkan form cari
                Dim ij As New Core_Module.frmCariMasterProduct
                ij.txtCari.Text = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                ij.ShowDialog()
                If ij.PrimaryProduct <> -1 Then
                    '2.a. product yang dicari ada --> load data deskripsi + primary + kode
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = ij.PrimaryProduct
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                ElseIf ij.PrimaryProduct = -1 Then
                    '2.b. kode product yang dicari tidak ada --> kembalikan ke kondisi buffer
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    If Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = "" Then
                        Me.DummyIndex = e.RowIndex
                        Me.BackgroundWorker3.RunWorkerAsync()
                        ij = Nothing
                        Exit Sub
                    End If
                End If
                ij = Nothing
            End If

            'pastikan tidak double dengan yang lainnya
            DummyIndex = Me.GetIndex_ItemProductGroup(Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex, _
            Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value)
            If DummyIndex <> -1 Then
                MsgBox(Me.dgvItemProductGroup.Item(Me.ColProductGroup_Level.Index, e.RowIndex).Value + _
                " code duplication, cancelling data change.", MsgBoxStyle.Information, "Information")
                If Buffer = "" Then
                    Me.DummyIndex = e.RowIndex
                    Me.BackgroundWorker3.RunWorkerAsync()
                Else
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.Buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError

    End Sub

    Private Sub dgvItemProductGroup_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemProductGroup.RowsAdded
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgvItemProductGroup.Rows(Me.dgvItemProductGroup.Rows.Count - 1)
            rowDGV.Cells(Me.ColProductGroup_CodeNo.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Desc.Index).Value = ""
            rowDGV.Cells(Me.ColProductGroup_Level.Index).Value = Me.objItem.GetProductLevelStr(Inventory_Module.ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup)
            rowDGV.Cells(Me.ColProductGroup_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value = -1
            rowDGV.Cells(Me.ColProductGroup_Remarks.Index).Value = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemProductGroup.RowsRemoved
        Try
            If ApakahPakeFilterProduct() = False Then
                Me.ChkAllProductSO.Checked = False
            End If
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemProductGroup_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemProductGroup.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
    End Sub

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
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
            If Me.DummyIndex <> -1 Then
                Me.dgvItemProductGroup.Rows.RemoveAt(Me.DummyIndex)
                DummyIndex = -1
                Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Function GetFilterProduct(ByVal TBLCetak As DseSalesOrderReportByDateDetail.tblCetakDataTable) As String
        Dim rowDGV As New DataGridViewRow
        Dim FilterGroup As String = ""
        Dim Row As DseSalesOrderReportByDateDetail.tblCetakRow
        Dim ListPrimarySO As New List(Of Double)
        Dim FilterSO As String = ""
        Try
            For Each rowDGV In Me.dgvItemProductGroup.Rows
                If rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value <> -1 Then
                    TipeProductGroup = Me.ObjSettingProd.GetLevelInt(rowDGV.Cells(Me.ColProductGroup_Level.Index).Value)
                    If FilterGroup = "" Then
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " [Primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    Else
                        If TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup Then
                            FilterGroup &= " OR [Primary Group] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Brand Then
                            FilterGroup &= " OR [Primary Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.SubBrand Then
                            FilterGroup &= " OR [Primary Sub Brand] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        ElseIf TipeProductGroup = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.Category Then
                            FilterGroup &= " OR [Primary Category] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        Else
                            FilterGroup &= " OR [Primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                        End If
                    End If
                End If
            Next

            If FilterGroup <> "" Then
                If Me.ChkAllProductSO.Checked = True Then
                    TBLCetak.DefaultView.RowFilter = FilterGroup
                    For i As Integer = 0 To TBLCetak.DefaultView.Count - 1
                        Row = TBLCetak.DefaultView(i).Row
                        If ListPrimarySO.Contains(Row.Primary_SO) = False Then
                            ListPrimarySO.Add(Row.Primary_SO)
                        End If
                    Next

                    For Each LS As Double In ListPrimarySO
                        If FilterSO = "" Then
                            FilterSO = "[Primary SO]=" & LS
                        Else
                            FilterSO &= " OR [Primary SO]= " & LS
                        End If
                    Next


                    'If FilterSO <> "" Then
                    '    FilterGroup = "(" & FilterGroup & ") AND (" & FilterSO & ")"
                    'Else
                    '    FilterGroup = "(" & FilterGroup & ")"
                    'End If
                    FilterGroup = "(" & FilterSO & ")"
                Else
                    FilterGroup = "(" & FilterGroup & ")"
                End If
            End If

            Return FilterGroup
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function ApakahPakeFilterProduct() As Boolean
        Dim N As Boolean = False
        Try
            Me.dgvItemProductGroup.AllowUserToAddRows = False
            If Me.dgvItemProductGroup.Rows.Count > 0 Then
                If String.IsNullOrEmpty(Me.dgvItemProductGroup.Item(Me.ColProductGroup_CodeNo.Index, 0).Value) = False Then
                    Return True
                End If
            End If
            Me.dgvItemProductGroup.AllowUserToAddRows = True
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class