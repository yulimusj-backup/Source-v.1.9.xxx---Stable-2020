Option Explicit On
Option Strict Off

Public Class FrmLaporanSalesReturnDetail
    Public BAckColorAll As Color = Color.FromArgb(170, 255, 190)
    Public UserCode As String = ""

    Dim ObjCd As New Core_Module.ClsCekData

    Dim buffer As String = ""
    Dim DummyIndex As Integer = -1
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1

    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim objCustomer As New Core_Module.clsMasterCustomer
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim ObjCategory As New Core_Module.ClsMasterCategory
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim EraseRowIndex As Integer = 0
    Dim ObjLicense As New Core_Module.ClsLicense

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
            Me.CetakDetail(Me.dtpTo.Value, Me.dtpFrom.Value, CDbl(Me.txtPrimaryProduct.Text), CDbl(Me.txtPrimaryCustomer.Text), CDbl(Me.txtPrimaryCurrency.Text))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakDetail(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimaryCustomer As Double, ByVal PrimaryCurrency As Double)
        Dim Report As New CrpLaporanSalesReturnDetail
        Dim Report2 As New CrpLaporanSalesReturnDetailGroup
        Dim DS As New DseLaporanSalesReturnDetail

        Dim DTACetakDummy As New DseLaporanSalesReturnDetailTableAdapters.Tbl_CetakTableAdapter
        Dim TBLCetak As DseLaporanSalesReturnDetail.Tbl_CetakDataTable = DS.Tbl_Cetak
        Dim TBLCetakDummy As New DseLaporanSalesReturnDetail.Tbl_CetakDataTable
        Dim TBLFilter As New DseLaporanSalesReturnDetail.Tbl_CetakDataTable
        Dim TBLTotalCurr As DseLaporanSalesReturnDetail.Total_CurrDataTable = DS.Total_Curr
        Dim RowCetakDummy As DseLaporanSalesReturnDetail.Tbl_CetakRow
        Dim RowCetak As DseLaporanSalesReturnDetail.Tbl_CetakRow
        Dim RowFilter As DseLaporanSalesReturnDetail.Tbl_CetakRow
        Dim RowTotalCurr As DseLaporanSalesReturnDetail.Total_CurrRow

        Dim DV As DataRowView
        Dim DV1 As DataRowView
        Dim DV2 As DataRowView
        Dim ListCombinasi As New List(Of Double)

        Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim filter As String = ""

        Dim Decimal_multi As Integer = 0
        Try
            DTACetakDummy.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            TBLCetak.Rows.Clear()
            TBLCetakDummy.Rows.Clear()

            paramStartDate = Me.dtpFrom.Value
            paramEndDate = Me.dtpTo.Value


            DTACetakDummy.Fill(TBLCetakDummy, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            DTACetakDummy.Fill(TBLFilter, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            filter = "[Primary SO] = [Primary SO]"

            If PrimaryCustomer <> -1 Then
                filter &= " AND [Primary Customer] = " & PrimaryCustomer & " OR [Primary Customer Quick] = " & PrimaryCustomer & ""
            End If


            'If CboTipePRTN.SelectedIndex = 1 Then
            '    filter &= " AND [Tipe PR] = 0"
            'ElseIf CboTipePRTN.SelectedIndex = 2 Then
            '    filter &= " AND [Tipe PR] = 1 AND [Invoice Cancellation] = 0"
            'ElseIf CboTipePRTN.SelectedIndex = 3 Then
            '    filter &= " AND [Tiper PR] = 1 AND [Invoice Cancellation] = 1"
            'End If

            'satu kondisi
            If ChkIRevision.Checked = True And ChkNotaRetur.Checked = False And ChkICancellation.Checked = False Then
                filter &= " AND [Tipe SR] = 1"
            ElseIf ChkIRevision.Checked = False And ChkNotaRetur.Checked = True And ChkICancellation.Checked = False Then
                filter &= " AND [Tipe SR] = 2 AND [Invoice Cancellation] = 0"
            ElseIf ChkIRevision.Checked = False And ChkNotaRetur.Checked = False And ChkICancellation.Checked = True Then
                filter &= " AND [Tipe SR] = 2 AND [Invoice Cancellation] = 1"
                'dua kondisi
            ElseIf ChkIRevision.Checked = True And ChkNotaRetur.Checked = True And ChkICancellation.Checked = False Then
                filter &= " AND [Invoice Cancellation] = 0"
            ElseIf ChkIRevision.Checked = False And ChkNotaRetur.Checked = True And ChkICancellation.Checked = True Then
                filter &= " AND [Tipe SR] = 2"
            ElseIf ChkIRevision.Checked = True And ChkNotaRetur.Checked = False And ChkICancellation.Checked = True Then
                filter &= " AND [Invoice Cancellation] <> 0 OR [Tipe SR] <> 2 "
            ElseIf ChkIRevision.Checked = False And ChkNotaRetur.Checked = False And ChkICancellation.Checked = False Then
                MsgBox("Please select at least 1 type return", MsgBoxStyle.Information, "Information")
                ChkIRevision.Focus()
                Exit Sub
            End If


            Dim FilterUser As String = ""
            'Filter Created By 
            If Me.DgvUser.Rows.Count > 0 Then
                For Each rowDGV1 As DataGridViewRow In Me.DgvUser.Rows
                    If Not IsDBNull(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) AndAlso rowDGV1.Cells(Me.ColCreatedBy.Index).Value <> "" AndAlso Not String.IsNullOrEmpty(rowDGV1.Cells(Me.ColCreatedBy.Index).Value) Then
                        If FilterUser = "" Then
                            FilterUser &= " [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        Else
                            FilterUser &= " Or [Created By] = '" & rowDGV1.Cells(Me.ColCreatedBy.Index).Value.ToString & "'"
                        End If
                    End If
                Next
            End If
            If FilterUser <> "" Then
                filter &= " AND (" & FilterUser & ") "
            End If

            If Me.ChkAllProductSRTN.Checked = True Then
                '======================================================================================================================================
                'Show All Product --> hanya filter primary PRTN yang punya Product ybs
                '======================================================================================================================================
                Dim ListPrimarySRTN As New List(Of Double)
                Dim FilterProduct As String = ""
                FilterProduct = "[Primary] = [Primary]"
                If PrimaryProduct <> -1 Then
                    FilterProduct &= " And [Primary Product] = " & PrimaryProduct
                End If

                'product group, dkk
                Dim FilterGroup As String = ""
                For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
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
                                FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
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
                                FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        End If
                    End If
                Next

                If FilterGroup <> "" Then
                    FilterProduct &= " AND (" & FilterGroup & ") "
                End If

                TBLFilter.DefaultView.RowFilter = Nothing
                TBLFilter.DefaultView.RowFilter = FilterProduct
                If TBLFilter.DefaultView.Count > 0 Then
                    For Each dvFilter As Data.DataRowView In TBLFilter.DefaultView
                        RowFilter = dvFilter.Row
                        With RowFilter
                            If ListPrimarySRTN.Contains(.Primary) = False Then
                                ListPrimarySRTN.Add(.Primary)
                            End If
                        End With
                    Next

                    Dim FilterPrimaryPRTN As String = ""
                    If ListPrimarySRTN.Count > 0 Then
                        For intA As Integer = 0 To ListPrimarySRTN.Count - 1
                            If FilterPrimaryPRTN = "" Then
                                FilterPrimaryPRTN &= " [Primary] = " & ListPrimarySRTN.Item(intA)
                            Else
                                FilterPrimaryPRTN &= " OR [Primary] = " & ListPrimarySRTN.Item(intA)
                            End If
                        Next
                    End If
                    filter &= " AND (" & FilterPrimaryPRTN & ") "
                Else
                    'jika filter product tidak sesuai maka kosongkan data
                    filter &= " AND [Primary] = 0 "
                End If
            Else
                '======================================================================================================================================
                'Show All filter primary Product yang sesuai
                '======================================================================================================================================
                If PrimaryProduct <> -1 Then
                    filter &= " And [Primary Product] = " & PrimaryProduct
                End If

                'product group, dkk
                Dim FilterGroup As String = ""
                For Each rowDGV As DataGridViewRow In Me.dgvItemProductGroup.Rows
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
                                FilterGroup &= " [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
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
                                FilterGroup &= " OR [primary Product] = " & rowDGV.Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value.ToString
                            End If
                        End If
                    End If
                Next

                If FilterGroup <> "" Then
                    filter &= " AND (" & FilterGroup & ") "
                End If
            End If
            If PrimaryCurrency <> -1 Then
                filter &= " AND [Primary Currency] = " & PrimaryCurrency & ""
            End If
            TBLCetakDummy.DefaultView.RowFilter = Nothing
            TBLCetakDummy.DefaultView.RowFilter = filter
            If Me.rboSRTNNumber.Checked = False Then
                TBLCetakDummy.DefaultView.Sort = "[Product Code] "
            End If
            If TBLCetakDummy.DefaultView.Count > 0 Then
                For Each DV In TBLCetakDummy.DefaultView
                    RowCetakDummy = DV.Row
                    TBLCetak.DefaultView.RowFilter = Nothing
                    TBLCetak.DefaultView.RowFilter = "[Primary] = " & RowCetakDummy.Primary.ToString & " AND [Primary Product] = " & RowCetakDummy.Primary_Product
                    If TBLCetak.DefaultView.Count > 0 Then
                        RowCetak = TBLCetak.DefaultView(0).Row
                        RowCetak.Total_Qty_Multi += RowCetakDummy.Total_Qty_Multi
                        RowCetak.Amount += RowCetakDummy.Amount
                    Else
                        'isi tblCetak dengan tblCetakDummy hasil filteran disini --------------------------------------------------------------------------
                        RowCetak = TBLCetak.NewTbl_CetakRow
                        RowCetak.Primary = RowCetakDummy.Primary
                        RowCetak.SR_No = RowCetakDummy.SR_No
                        RowCetak.SR_Date = RowCetakDummy.SR_Date
                        RowCetak.Primary_Currency = RowCetakDummy.Primary_Currency
                        RowCetak.Currency_Code = RowCetakDummy.Currency_Code
                        RowCetak.Primary_Customer = RowCetakDummy.Primary_Customer
                        RowCetak.Primary_Customer_Quick = RowCetakDummy.Primary_Customer_Quick
                        RowCetak.Customer_Code = RowCetakDummy.Customer_Code
                        RowCetak.Customer_Name = RowCetakDummy.Customer_Name
                        RowCetak.Customer_Code_Quick = RowCetakDummy.Customer_Code_Quick
                        RowCetak.Customer_Name_Quick = RowCetakDummy.Customer_Name_Quick
                        If IsNumeric(RowCetakDummy.Payment_Term) AndAlso RowCetakDummy.Payment_Term = 0 Then
                            RowCetak.Payment_Term_Days = "-"
                        ElseIf IsNumeric(RowCetakDummy.Payment_Term) AndAlso RowCetakDummy.Payment_Term = 1 Then
                            RowCetak.Payment_Term_Days = RowCetakDummy.Payment_Term & " DAY"
                        Else
                            RowCetak.Payment_Term_Days = RowCetakDummy.Payment_Term & " DAYS"
                        End If
                        RowCetak.Payment_Term = RowCetakDummy.Payment_Term
                        RowCetak.Tipe_SR = RowCetakDummy.Tipe_SR
                        RowCetak.Primary_SO = RowCetakDummy.Primary_SO
                        RowCetak.SO_No = RowCetakDummy.SO_No
                        RowCetak.Primary_Product = RowCetakDummy.Primary_Product
                        RowCetak.Product_Code = RowCetakDummy.Product_Code
                        RowCetak.Product_Desc = RowCetakDummy.Product_Desc
                        RowCetak.Primary_Units = RowCetakDummy.Primary_Units
                        RowCetak.Primary_Units_Multi = RowCetakDummy.Primary_Units_Multi
                        RowCetak.Unit_Code_Multi = RowCetakDummy.Unit_Code_Multi
                        RowCetak._Decimal = RowCetakDummy._Decimal
                        RowCetak.Nett_Price = RowCetakDummy.Nett_Price
                        RowCetak.Sub_Total = 0
                        RowCetak.Tax_Amount = RowCetakDummy.Tax_Amount
                        RowCetak.Total = 0
                        RowCetak.Default_Precision = RowCetakDummy.Default_Precision
                        RowCetak.Created_By = RowCetakDummy.Created_By
                        RowCetak.Ratio_Desc = RowCetakDummy.Ratio_Desc
                        RowCetak.Ratio_Reference = RowCetakDummy.Ratio_Reference
                        RowCetak.Ratio_Default = RowCetakDummy.Ratio_Default
                        RowCetak.Total_Qty_Multi = RowCetakDummy.Total_Qty_Multi
                        RowCetak.Unit_Code_Multi = RowCetakDummy.Unit_Code_Multi
                        RowCetak.Unit_Code_Default = RowCetakDummy.Unit_Code_Default
                        RowCetak.Primary_Category = RowCetakDummy.Primary_Category
                        RowCetak.Primary_Sub_Brand = RowCetakDummy.Primary_Sub_Brand
                        RowCetak.Primary_Brand = RowCetakDummy.Primary_Brand
                        RowCetak.Primary_Group = RowCetakDummy.Primary_Group
                        RowCetak.Amount = RowCetakDummy.Amount
                        RowCetak.Primary_Curr_Tax = RowCetakDummy.Primary_Curr_Tax
                        RowCetak.Currency_Code_Tax = RowCetakDummy.Currency_Code_Tax
                        RowCetak._Default__Precision__Tax = RowCetakDummy._Default__Precision__Tax
                        RowCetak.Tax_Percent = RowCetakDummy.Tax_Percent
                        RowCetak.Tax_Exchange_Rate = RowCetakDummy.Tax_Exchange_Rate
                        RowCetak.Duty_stamp = RowCetakDummy.Duty_stamp
                        If RowCetakDummy.Tipe_SR = 1 Then
                            'tipe SR = invoice revisi amount di nol kan berdasarkan request Pa Iwan tertanggal 10/10/2016. 
                            'Untuk Error no 70. ammont di 0 kan jika tipe Invoice Revision
                            'Format laporan untuk perhitungan exchange rate mengikuti laporan sales by invoice. TTD Taufik
                            RowCetak.Return_Type = "Invoice Revision"
                            RowCetak.Nett_Price = 0
                            RowCetak.Amount = 0
                        ElseIf RowCetakDummy.Tipe_SR = 2 AndAlso RowCetakDummy.Invoice_Cancellation = 0 Then
                            RowCetak.Return_Type = "Tax Return Notes"
                        ElseIf RowCetakDummy.Tipe_SR = 2 AndAlso RowCetakDummy.Invoice_Cancellation = 1 Then
                            RowCetak.Return_Type = "Invoice Cancellation"
                        End If
                        RowCetak.Invoice_Cancellation = RowCetakDummy.Invoice_Cancellation
                        RowCetak.VAT = 0

                        RowCetak.Decimal_multi = RowCetakDummy.Dec_Multi
                        RowCetak.Dec_Multi = RowCetakDummy.Dec_Multi
                        RowCetak.Primary_Units_Multi = RowCetakDummy.Primary_Units_Multi
                        RowCetak.price_by_multi_unit = RowCetakDummy.price_by_multi_unit
                        RowCetak.Total_Qty = RowCetakDummy.Total_Qty

                        If RowCetakDummy.Primary_Units <> RowCetakDummy.Primary_Units_Multi And RowCetakDummy.price_by_multi_unit = 0 Then
                            RowCetak.Total_Qty_Multi = RowCetakDummy.Total_Qty
                            RowCetak.Unit_Code_Multi = RowCetakDummy.Unit_Code_Default
                            RowCetak.Dec_Multi = RowCetakDummy._Decimal
                        End If

                        If RowCetakDummy.Primary_Units_Multi <> RowCetakDummy.Primary_Units Then
                            RowCetak.Ratio_Desc = "(" & Me.ObjCd.DigitGrouping(RowCetakDummy.Ratio_Reference, RowCetakDummy.Dec_Multi) & " " & _
                            RowCetakDummy.Unit_Code_Multi & " = " & Me.ObjCd.DigitGrouping(RowCetakDummy.Ratio_Default, RowCetakDummy._Decimal) & " " & _
                            RowCetakDummy.Unit_Code_Default & ")"

                            RowCetak.Ratio_Desc = "(" & Me.ObjCd.DigitGrouping(RowCetakDummy.Total_Qty_Multi, RowCetakDummy.Dec_Multi) & " " & RowCetakDummy.Unit_Code_Multi & ")"
                        Else
                            RowCetak.Ratio_Desc = ""
                        End If

                        TBLCetak.Rows.Add(RowCetak)
                    End If
                Next
            End If

            For Each RowCetak In TBLCetak.Rows
                'Dim DecSubTotal As Decimal = 0
                TBLCetakDummy.DefaultView.RowFilter = Nothing
                If PrimaryProduct <> -1 Then
                    If Me.ChkAllProductSRTN.Checked = False Then
                        TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary] = " & RowCetak.Primary & _
                                                              " AND [Primary Product] = " & PrimaryProduct & ""
                    Else
                        TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary] = " & RowCetak.Primary
                    End If
                Else
                    TBLCetakDummy.DefaultView.RowFilter = filter & " And [Primary] = " & RowCetak.Primary
                End If

                For Each DV1 In TBLCetakDummy.DefaultView
                    RowCetakDummy = DV1.Row
                    If RowCetakDummy.Tipe_SR <> 1 Then
                        If IsNumeric(RowCetakDummy.Amount) Then
                            RowCetak.Sub_Total += CDec(RowCetakDummy.Amount)
                            RowCetak.VAT = RowCetakDummy.Sub_Total * (CDec(RowCetakDummy.Tax_Percent) / 100) 'CDec(RowCetakDummy.Tax_Amount) 
                        End If
                    End If
                Next

                If RowCetak.Return_Type <> "Invoice Revision" Then
                    RowCetak.Total = RowCetak.Sub_Total + RowCetak.VAT
                Else
                    RowCetak.Total = 0
                End If
            Next

            For Each RowCetak In TBLCetak.Rows
                If RowCetak.Primary_Currency <> -1 Then
                    If RowCetak.Primary_Currency <> -1 Then
                        If ListCombinasi.Contains(RowCetak.Primary_Currency) = False Then
                            ListCombinasi.Add(RowCetak.Primary_Currency)
                        End If
                    End If
                End If
            Next

            For intX As Integer = 0 To ListCombinasi.Count - 1
                RowTotalCurr = TBLTotalCurr.NewTotal_CurrRow
                RowTotalCurr.Sub_Total = 0
                RowTotalCurr.VAT = 0
                RowTotalCurr.Primary_Currency = ListCombinasi(intX)
                RowTotalCurr.Currency_Code = ""
                RowTotalCurr.Default_Precision = 0
                TBLTotalCurr.Rows.Add(RowTotalCurr)

                TBLCetak.DefaultView.RowFilter = Nothing
                TBLCetak.DefaultView.RowFilter = " [Primary Currency] = " & ListCombinasi(intX)
                If TBLCetak.DefaultView.Count > 0 Then
                    Dim ListPrimary As New List(Of Double)
                    For Each dv3 As Data.DataRowView In TBLCetak.DefaultView
                        RowCetak = dv3.Row
                        If ListPrimary.Contains(RowCetak.Primary) = False Then
                            ListPrimary.Add(RowCetak.Primary)
                            RowTotalCurr.Sub_Total += CDec(RowCetak.Sub_Total)
                            RowTotalCurr.VAT += CDec(RowCetak.VAT)
                            RowTotalCurr.Total += CDec(RowCetak.Total)
                            RowTotalCurr.Currency_Code = RowCetak.Currency_Code
                            RowTotalCurr.Default_Precision = RowCetak.Default_Precision
                        End If
                    Next
                End If
            Next

            If Me.rboSRTNNumber.Checked = True Then
                'generate report
                Report.SetDataSource(DS)
                Report.SetParameterValue(Report.Parameter_StartDate.ParameterFieldName, paramStartDate)
                Report.SetParameterValue(Report.Parameter_EndDate.ParameterFieldName, paramEndDate)
                crv.ReportSource = Report
                crv.Show()
            Else
                ''generate report
                Report2.SetDataSource(DS)
                Report2.SetParameterValue(Report2.Parameter_StartDate.ParameterFieldName, paramStartDate)
                Report2.SetParameterValue(Report2.Parameter_EndDate.ParameterFieldName, paramEndDate)
                crv.ReportSource = Report2
                crv.Show()
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

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
                    If Me.objCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                        Me.txtPrimaryCustomer.Text = CStr(Me.objCustomer.GetPrimary(Me.txtCustomerCode.Text))
                        Me.txtCustomerCode.Text = Me.objCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCustomerName.Text = Me.objCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCurrencyCode.Focus()
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
                If Me.objCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = CStr(Me.objCustomer.GetPrimary(Me.txtCustomerCode.Text))
                    Me.txtCustomerCode.Text = Me.objCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                    Me.txtCustomerName.Text = Me.objCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        hj.txtCari.Text = Me.txtCustomerCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = CStr(hj.PrimaryCustomer)
                        Me.txtCustomerCode.Text = Me.objCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCustomerName.Text = Me.objCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                        Me.txtCurrencyCode.Focus()
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
                Me.txtCustomerCode.Text = Me.objCustomer.GetCustomerCode(CDbl(Me.txtPrimaryCustomer.Text))
                Me.txtCustomerName.Text = Me.objCustomer.GetCustName(CDbl(Me.txtPrimaryCustomer.Text))
                Me.txtCurrencyCode.Focus()
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


    Private Sub txtProductCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.GotFocus, txtCurrencyCode.GotFocus, txtCustomerCode.GotFocus
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

    Private Sub txtProductCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProductCode.LostFocus, txtCustomerCode.LostFocus, txtCurrencyCode.LostFocus
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
                Me.ChkAllProductSRTN.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProductPO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProductSRTN.Click
        Try
            If Me.txtPrimaryProduct.Text = -1 AndAlso (Me.dgvItemProductGroup.RowCount > 0 AndAlso _
                                                      CDec(Me.dgvItemProductGroup.Rows(0).Cells(Me.ColProductGroup_PrimaryAcuan.Index).Value) <= 0) Then
                Me.ChkAllProductSRTN.Checked = False
                MsgBox("Product Or " & Me.dgvItemProductGroup.Rows(0).Cells(Me.ColProductGroup_Level.Index).Value & " cannot be an empty string ", MsgBoxStyle.Information, "Information")
                Me.txtProductCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseOrderReportByDateDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItemProductGroup.Rows.Clear()
            Me.DgvUser.Rows.Clear()
            Me.ColProductGroup_Remarks.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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
                            Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                If Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = -1 Then
                    '2. jika bukan baris baru dan primary =-1
                    'yang bisa diedit hanya code product group dan level
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or e.ColumnIndex = Me.ColProductGroup_Level.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    '3. jika bukan baris baru dan primary <> -1
                    'yang bisa diedit: code, remarks
                    If e.ColumnIndex = Me.ColProductGroup_CodeNo.Index Or _
                    e.ColumnIndex = Me.ColProductGroup_Remarks.Index Then
                        Me.buffer = Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value
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
                        Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProductGroup.GetPrimary(Me.buffer, vProductGroup_Level)
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjCategory.GetPrimary(Me.buffer)
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
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
                    Me.dgvItemProductGroup.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.buffer)
                    Me.dgvItemProductGroup.Item(Me.ColProductGroup_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(Me.dgvItemProductGroup.Item(Me.ColProductGroup_PrimaryAcuan.Index, e.RowIndex).Value)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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

    Private Sub dgvItemProductGroup_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemProductGroup.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemProductGroup)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDGV_ItemGroup()
        Dim ObjSettingPG As New Core_Module.ClsTransaksiSetting_ProdGroup

        Try
            ObjSettingPG.IsikanDGV(Me.UserCode, Core_Module.ClsTransaksiSetting_ProdGroup.EnumTipeTransaksi.LaporanPurchaseRequest, Me.dgvItemProductGroup, _
            Me.ColProductGroup_No.Index, Me.ColProductGroup_Level.Index, Me.ColProductGroup_CodeNo.Index, Me.ColProductGroup_Desc.Index, _
            Me.ColProductGroup_Remarks.Index, Me.ColProductGroup_PrimaryAcuan.Index)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSettingPG = Nothing
        End Try
    End Sub

    Private Sub dgvItemProductGroup_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItemProductGroup.DataError
        'do nothing -> dibuat untuk handle data error event, jangan dihapuskan
    End Sub


    Private Sub DgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvUser.CellBeginEdit
        Try
            If Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value <> -1 Then
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                    Me.buffer = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vKodeUser As String, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvUser.Rows.Count - 2
                RowDGV = Me.DgvUser.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColCreatedBy.Index).Value = vKodeUser Then
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

    Private Sub DgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvUser.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvUser.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim KodeUser As String = ""
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColCreatedBy.Index Then
                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        'kemungkinan
                        If Me.ObjUser.ApakahSudahAda(Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            '1. kode akun yang diketik ada --> load data deskripsi + primary akun
                            KodeUser = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            index = Me.GetIndexItem(KodeUser, e.RowIndex)
                            If index < 0 Then
                                Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                                Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value
                                Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BgwUser.RunWorkerAsync()
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmCariMasterUser
                            ij.TxtCari.Text = Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.backColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Kodeuser <> "" Then
                                index = Me.GetIndexItem(ij.Kodeuser, e.RowIndex)
                                If index < 0 Then
                                    Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                    Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                    Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer) = True Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        Else
                                            Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        End If
                                        Exit Sub
                                    Else
                                        Me.DgvUser.Item(Me.ColCreatedBy.Index, e.RowIndex).Value = ij.Kodeuser
                                        Me.DgvUser.Item(Me.ColCreatedBy_Name.Index, e.RowIndex).Value = Me.ObjUser.GetNama(ij.Kodeuser)
                                        Me.DgvUser.Item(Me.ColCreatedBy_Primary.Index, e.RowIndex).Value = Me.ObjUser.GetPrimary(ij.Kodeuser)
                                    End If
                                End If
                            ElseIf ij.Kodeuser = "" Then
                                If e.ColumnIndex = Me.ColCreatedBy.Index Then
                                    If String.IsNullOrEmpty(Me.buffer) = True Then
                                        If Me.DgvUser.Rows.Count - 1 <> e.RowIndex Then
                                            Me.EraseRowIndex = e.RowIndex
                                            Me.BgwUser.RunWorkerAsync()
                                        End If
                                    Else
                                        Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    End If
                                Else
                                    Me.DgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                End If
                            End If
                            ij = Nothing
                        End If
                        'Me.ObjCD.SetNumber(Me.DgvUser)
                    End If
                Else
                    Me.EraseRowIndex = e.RowIndex
                    Me.BgwUser.RunWorkerAsync()
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvUser.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            ' If Me.TxtNoPO.Enabled = True Then
            rowDGV = Me.DgvUser.Rows(Me.DgvUser.RowCount - 1)
            With rowDGV
                ' .Cells(Me.ColN.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColCreatedBy.Index).Value = ""
                .Cells(Me.ColCreatedBy_Name.Index).Value = ""
                .Cells(Me.ColCreatedBy_Primary.Index).Value = "-1"
            End With
            ' End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwUser.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwUser_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwUser.RunWorkerCompleted
        Try
            Me.DgvUser.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
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

    Private Sub rboProduct_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboProduct.CheckedChanged

        Try
            If Me.rboProduct.Checked = True Then
                Me.ChkAllProductSRTN.Enabled = False
            ElseIf Me.rboSRTNNumber.Checked = True Then
                Me.ChkAllProductSRTN.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProductPRTN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAllProductSRTN.CheckedChanged

    End Sub

    Private Sub crv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ChkAOPCancellation_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkICancellation.CheckedChanged

    End Sub
End Class