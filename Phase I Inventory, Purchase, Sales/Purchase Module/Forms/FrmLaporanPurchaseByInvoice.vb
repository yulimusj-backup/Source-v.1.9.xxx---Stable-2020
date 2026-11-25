Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseByInvoice
    Public BAckColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjAccount As New Core_Module.ClsMasterAkun1
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjSettingProd As New Core_Module.ClsTransaksiSetting_ProdGroup
    Dim objGroup As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim objItem As New Inventory_Module.ClsTransaksiMaterialPlanningGroupItem
    Dim TipeProductGroup As Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel = Core_Module.ClsTransaksiSetting_ProdGroup.EnumGroupLevel.ProductGroup
    Dim ProductGroup_Level As Integer = 3
    Dim Brand_Level As Integer = 2
    Dim SubBrand_Level As Integer = 1 '
    Dim ObjProductGroup As New Core_Module.ClsMasterProductGroup
    Dim DummyIndex As Integer = -1
    Dim Buffer As String = ""

    Private Sub FrmLaporanPurchaseReportRekap_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboFormat.SelectedIndex = 0
            Me.ShowIcon = False
            Me.dgvItemProductGroup.AlternatingRowsDefaultCellStyle.BackColor = Me.BAckColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub BtnPrint_Click(ByVal sSender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.DtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.DtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.DtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("Starting period must be smaller than ending period", MsgBoxStyle.Information, "Information")
                Me.DtpFrom.Focus()
                Exit Sub
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.CboFormat.SelectedIndex = 0 Then
                Me.PrintRekap(Me.DtpTo.Value, Me.DtpFrom.Value, CDbl(Me.txtPrimaryCodeNo.Text), CDbl(Me.TxtPrimarySupplier.Text), CDbl(Me.TxtPrimaryAcc.Text))
            Else
                Me.PrintDetail(Me.DtpTo.Value, Me.DtpFrom.Value, CDbl(Me.txtPrimaryCodeNo.Text), CDbl(Me.TxtPrimarySupplier.Text), CDbl(Me.TxtPrimaryAcc.Text))
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini dipanggil jika menggunakan filter product code dan check all item po
    Private Sub GetAllProductPO(ByRef vTBLCetak1 As DseLaporanPurchaseReport_Rekap.TblDetailDataTable, ByVal vPrimaryProduct As Double)
        Dim TBLCetak1_New As New DseLaporanPurchaseReport_Rekap.TblDetailDataTable
        Dim RowCetak1_New As DseLaporanPurchaseReport_Rekap.TblDetailRow
        Dim ListPrimaryPO As New List(Of Double)

        Try
            TBLCetak1_New = vTBLCetak1.Copy
            vTBLCetak1.Clear()
            TBLCetak1_New.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct & ""
            If TBLCetak1_New.DefaultView.Count > 0 Then
                For Each dv As DataRowView In TBLCetak1_New.DefaultView
                    RowCetak1_New = dv.Row
                    If ListPrimaryPO.Contains(RowCetak1_New.Primary) = False Then
                        ListPrimaryPO.Add(RowCetak1_New.Primary)
                    End If
                Next

                If ListPrimaryPO.Count > 0 Then
                    For Each PrimaryPO As Double In ListPrimaryPO
                        TBLCetak1_New.DefaultView.RowFilter = "[Primary] = " & PrimaryPO
                        If TBLCetak1_New.DefaultView.Count > 0 Then
                            vTBLCetak1.Merge(TBLCetak1_New.DefaultView.ToTable)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLCetak1_New.Dispose()
        End Try
    End Sub

    Private Sub PrintRekap(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimarySupplier As Double, ByVal PrimaryAccount As Double)
        Dim DS As New DseLaporanPurchaseReport_Rekap
        Dim ReportRekap As New CrpLaporanPurchaseReport_Rekap
        Dim DTACetak1 As New DseLaporanPurchaseReport_RekapTableAdapters.TblDetailTableAdapter
        Dim DTACetak3 As New DseLaporanPurchaseReport_RekapTableAdapters.tbltrans_aop_itemcostTableAdapter
        Dim TBLCetak1 As New DseLaporanPurchaseReport_Rekap.TblDetailDataTable
        Dim TBLCetak3 As New DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostDataTable
        Dim TBLCetakRekap As DseLaporanPurchaseReport_Rekap.TblCetak_RekapDataTable = DS.TblCetak_Rekap
        Dim TBLTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalDataTable = DS.TBLDetail_Total
        Dim RowCetak1, RowCetak1_Aft As DseLaporanPurchaseReport_Rekap.TblDetailRow
        Dim RowCetak3 As DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostRow
        Dim RowCetakRekap As DseLaporanPurchaseReport_Rekap.TblCetak_RekapRow
        Dim RowTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalRow

        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim primaryMax As Double = -1
        Dim primaryMin As Double = -1
        Dim filter As String = ""
        Dim HitungSubTotal As Decimal = 0
        Dim FilterGroup As String = ""
        Try
            DTACetak1.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak3.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            paramStartDate = Me.DtpFrom.Value
            paramEndDate = Me.DtpTo.Value

            'Isi tblCetak ----------------------------------------------------------------------------------------------------------------
            DTACetak1.FillBy_StartEndDate(TBLCetak1, paramStartDate, paramEndDate)
            DTACetak3.FillBy_StartEndDate(TBLCetak3, paramEndDate, paramStartDate)

            'Filter tblCetak1 ------------------------------------------------------------------------------------------------------
            filter = "[Primary] = [Primary]"
            If PrimaryProduct <> -1 Then
                'If Me.ChkAllProductPO.Checked = False Then
                '    filter &= "AND [Primary Product] = " & PrimaryProduct & ""
                'Else
                '    Me.GetAllProductPO(TBLCetak1, PrimaryProduct)
                'End If
            End If
            If PrimarySupplier <> -1 Then
                filter &= " AND [Primary Supplier] = " & PrimarySupplier & ""
            End If
            If PrimaryAccount <> -1 Then
                filter &= " AND [primary akun ap] = " & PrimaryAccount & ""
            End If
            FilterGroup = Me.GetFilterProduct(TBLCetak1)
            If FilterGroup <> "" Then
                filter &= " AND " & FilterGroup
            End If

            TBLCetak1.DefaultView.RowFilter = Nothing
            TBLCetak1.DefaultView.Sort = Nothing
            TBLCetak1.DefaultView.RowFilter = filter
            If Me.ChkUsingSupplierNumber.Checked = False Then
                TBLCetak1.DefaultView.Sort = "[Currency Code], [Tanggal Transaksi], [AOP Number], [Product Code]"
            End If
            If TBLCetak1.DefaultView.Count > 0 Then
                For idx As Integer = 0 To TBLCetak1.DefaultView.Count - 1
                    RowCetak1 = TBLCetak1.DefaultView(idx).Row
                    With RowCetak1
                        RowCetakRekap = TBLCetakRekap.NewTblCetak_RekapRow
                        RowCetakRekap.Primary_Transaksi = .Primary
                        RowCetakRekap.Curr_Code = .Currency_Code
                        RowCetakRekap.Curr_Description = .Currency_Description
                        RowCetakRekap.AOP_No = .AOP_Number
                        If Me.ChkUsingSupplierNumber.Checked = True Then
                            If .Invoice_No <> "" Then
                                RowCetakRekap.Invoice_No = .Invoice_No
                            Else
                                RowCetakRekap.Invoice_No = .AOP_Number
                            End If
                        Else
                            RowCetakRekap.Invoice_No = .AOP_Number
                        End If
                        RowCetakRekap._Date = .Tanggal_Transaksi
                        RowCetakRekap.Supplier_Desc = .Supplier_Name
                        If IsNumeric(.Payment_Term) Then
                            RowCetakRekap.Payment_Term = .Payment_Term & " DAYS"
                        Else
                            RowCetakRekap.Payment_Term = .Payment_Term
                        End If
                        RowCetakRekap.Default_Precision = .Default_Precision
                        RowCetakRekap.Amount = .Amount

                        HitungSubTotal += .Amount

                        Dim ApakahHitung As Boolean = True
                        If idx < TBLCetak1.DefaultView.Count - 1 Then
                            RowCetak1_Aft = TBLCetak1.DefaultView(idx + 1).Row
                            If RowCetakRekap.Primary_Transaksi = RowCetak1_Aft.Primary Then
                                ApakahHitung = False
                            End If
                        End If

                        If ApakahHitung = True Then
                            RowCetakRekap.SubTotal = HitungSubTotal
                            RowCetakRekap.Down_Payment = .Down_Payment * -1
                            If .Tax_Exclude = True Then
                                RowCetakRekap.VAT = ((HitungSubTotal - .Down_Payment) * .Persentase) / 100
                            Else
                                RowCetakRekap.VAT = 0
                            End If
                            RowCetakRekap.Total_Amount = HitungSubTotal - .Down_Payment + RowCetakRekap.VAT

                            'Hitungan Subtotal Other Cost 
                            TBLCetak3.DefaultView.RowFilter = Nothing
                            TBLCetak3.DefaultView.RowFilter = "[Primary AOP] = " & RowCetakRekap.Primary_Transaksi & " AND [Primary Currency] = " & .Primary_Currency
                            If TBLCetak3.DefaultView.Count > 0 Then
                                RowCetak3 = TBLCetak3.DefaultView(0).Row
                                RowCetakRekap.Other_Cost = RowCetak3.Other_Cost
                            End If

                            'Hitung Total Currency
                            TBLTotalCurr.DefaultView.RowFilter = "[Primary Currency] = " & .Primary_Currency
                            If TBLTotalCurr.DefaultView.Count > 0 Then
                                RowTotalCurr = TBLTotalCurr.DefaultView(0).Row
                                RowTotalCurr.Total_VAT += RowCetakRekap.VAT
                                RowTotalCurr.Total_Amount += RowCetakRekap.Total_Amount
                                RowTotalCurr.Total_SubTotal += RowCetakRekap.SubTotal
                                RowTotalCurr.Total_Down_Payment += RowCetakRekap.Down_Payment
                                RowTotalCurr.Other_Cost_Amount += RowCetakRekap.Other_Cost
                            Else
                                RowTotalCurr = TBLTotalCurr.NewTBLDetail_TotalRow
                                RowTotalCurr.Primary_AOP = RowCetakRekap.Primary_Transaksi
                                RowTotalCurr.Primary_Currency = .Primary_Currency
                                RowTotalCurr.Total_Description = ObjCurr.GetCurrencyCode(.Primary_Currency)
                                RowTotalCurr.Other_Cost_Desc = ObjCurr.GetCurrencyCode(.Primary_Currency)
                                RowTotalCurr.Total_VAT = RowCetakRekap.VAT
                                RowTotalCurr.Total_Amount = RowCetakRekap.Total_Amount
                                RowTotalCurr.Total_SubTotal = RowCetakRekap.SubTotal
                                RowTotalCurr.Total_Down_Payment = RowCetakRekap.Down_Payment
                                RowTotalCurr.Other_Cost_Amount = RowCetakRekap.Other_Cost
                                RowTotalCurr.Default_Precision = ObjCurr.GetDefaultPrecission(.Primary_Currency)
                                TBLTotalCurr.Rows.Add(RowTotalCurr)
                            End If

                            HitungSubTotal = 0
                        End If
                        TBLCetakRekap.Rows.Add(RowCetakRekap)
                    End With
                Next
            End If

            ReportRekap.SetDataSource(DS)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_startDate.ParameterFieldName, paramStartDate)
            ReportRekap.SetParameterValue(ReportRekap.Parameter_endDate.ParameterFieldName, paramEndDate)
            Crv.ReportSource = ReportRekap
            Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak1.Dispose()
            DTACetak3.Dispose()
            TBLCetak1.Dispose()
            TBLCetak3.Dispose()
            TBLCetakRekap.Dispose()
            TBLTotalCurr.Dispose()
        End Try
    End Sub

    Private Sub PrintRekap_Obs(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimarySupplier As Double, ByVal PrimaryAccount As Double)
        'Dim DTACetak1 As New DseLaporanPurchaseReport_RekapTableAdapters.TblDetail_Header1TableAdapter
        'Dim DTACetak2 As New DseLaporanPurchaseReport_RekapTableAdapters.TBLDetail_ItemTableAdapter
        'Dim DTACetak3 As New DseLaporanPurchaseReport_RekapTableAdapters.tbltrans_aop_itemcostTableAdapter
        'Dim TBLCetakRekap As DseLaporanPurchaseReport_Rekap.TblCetak_RekapDataTable = DS.TblCetak_Rekap
        'Dim TBLCetak1 As DseLaporanPurchaseReport_Rekap.TblDetail_Header1DataTable = DS.TblDetail_Header1
        'Dim TBLCetak2 As DseLaporanPurchaseReport_Rekap.TBLDetail_ItemDataTable = DS.TBLDetail_Item
        'Dim TBLCetak3 As DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostDataTable = DS.tbltrans_aop_itemcost
        'Dim TBLSubtotal As DseLaporanPurchaseReport_Rekap.TBLDetail_SubtotalDataTable = DS.TBLDetail_Subtotal
        'Dim TBLTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalDataTable = DS.TBLDetail_Total
        'Dim RowCetakRekap As DseLaporanPurchaseReport_Rekap.TblCetak_RekapRow
        'Dim RowCetak1 As DseLaporanPurchaseReport_Rekap.TblDetail_Header1Row
        'Dim RowCetak2 As DseLaporanPurchaseReport_Rekap.TBLDetail_ItemRow
        'Dim RowCetak3 As DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostRow
        'Dim RowSubtotal As DseLaporanPurchaseReport_Rekap.TBLDetail_SubtotalRow
        'Dim RowTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalRow
        'Dim DV1 As DataRowView
        'Dim DV2 As DataRowView
        'Dim DV3 As DataRowView
        'Dim DV4 As DataRowView

        'Dim paramStartDate As Date = Now
        'Dim paramEndDate As Date = Now
        'Dim primaryMax As Double = -1
        'Dim primaryMin As Double = -1
        'Dim filter As String = ""
        'Dim filter2 As String = ""
        'Try
        '    DTACetak1.Connection.ConnectionString = Core_Module.GetConstring
        '    DTACetak2.Connection.ConnectionString = Core_Module.GetConstring
        '    DTACetak3.Connection.ConnectionString = Core_Module.GetConstring
        '    Windows.Forms.Cursor.Current = Cursors.WaitCursor

        '    paramStartDate = Me.DtpFrom.Value
        '    paramEndDate = Me.DtpTo.Value
        '    TBLCetakRekap.Rows.Clear()
        '    TBLSubtotal.Rows.Clear()
        '    TBLTotalCurr.Rows.Clear()

        '    'Isi tblCetakDummy ----------------------------------------------------------------------------------------------------------------
        '    If Me.ChkAllProductPO.Checked = True Then
        '        DTACetak1.FillBy_AllProductTransaksi(TBLCetak1, paramStartDate, paramEndDate, PrimaryProduct)
        '    Else
        '        DTACetak1.FillBy_StartEndDate(TBLCetak1, paramStartDate, paramEndDate)
        '    End If
        '    'TBLCetak1.DefaultView.Sort = Nothing
        '    'TBLCetak1.DefaultView.Sort = "[Primary]"
        '    'If TBLCetak1.DefaultView.Count > 0 Then
        '    '    RowCetak1 = TBLCetak1.DefaultView(0).Row
        '    '    primaryMin = RowCetak1.Primary
        '    '    RowCetak1 = TBLCetak1.DefaultView(TBLCetak1.DefaultView.Count - 1).Row
        '    '    primaryMax = RowCetak1.Primary
        '    'End If
        '    DTACetak2.Fill(TBLCetak2)
        '    DTACetak3.Fill(TBLCetak3)

        '    'Filter tblCetakDummy disini ------------------------------------------------------------------------------------------------------
        '    filter = "[Primary] = [Primary]"
        '    If PrimarySupplier <> -1 Then
        '        filter &= " AND [Primary Supplier] = " & PrimarySupplier & ""
        '    End If
        '    If PrimaryAccount <> -1 Then
        '        filter &= " AND [primary akun ap] = " & PrimaryAccount & ""
        '    End If

        '    TBLCetak1.DefaultView.RowFilter = Nothing
        '    TBLCetak1.DefaultView.RowFilter = filter
        '    If TBLCetak1.DefaultView.Count > 0 Then
        '        For Each DV1 In TBLCetak1.DefaultView
        '            RowCetak1 = DV1.Row
        '            If PrimaryProduct <> -1 Then
        '                If Me.ChkAllProductPO.Checked = False Then
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString & " AND [Primary Product] = " & PrimaryProduct & ""
        '                Else
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '                End If
        '            Else
        '                filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '            End If

        '            TBLCetak2.DefaultView.RowFilter = Nothing
        '            TBLCetak2.DefaultView.RowFilter = filter2
        '            If TBLCetak2.DefaultView.Count > 0 Then
        '                For Each DV2 In TBLCetak2.DefaultView
        '                    RowCetak2 = DV2.Row
        '                    'Hitung Subtotal
        '                    TBLSubtotal.DefaultView.RowFilter = Nothing
        '                    If PrimaryProduct <> -1 Then
        '                        If Me.ChkAllProductPO.Checked = False Then
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary & " AND [Primary Product] = " & PrimaryProduct & ""
        '                        Else
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '                        End If
        '                    Else
        '                        TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary
        '                    End If

        '                    If TBLSubtotal.DefaultView.Count > 0 Then
        '                        For Each DV3 In TBLSubtotal.DefaultView
        '                            RowSubtotal = DV3.Row
        '                            If IsNumeric(RowCetak2.Amount) Then
        '                                RowSubtotal.Subtotal += CDec(RowCetak2.Amount)
        '                                RowSubtotal.Primary_Product = RowCetak2.Primary_Product
        '                                RowSubtotal.DP = -RowCetak1.Down_Payment
        '                                If RowCetak1.Tax_Exclude = True Then
        '                                    RowSubtotal.VAT = ((RowSubtotal.Subtotal - RowCetak1.Down_Payment) * CDec(RowCetak1.Persentase)) / 100
        '                                Else
        '                                    RowSubtotal.VAT = 0
        '                                End If
        '                                RowSubtotal.Total = RowSubtotal.Subtotal - (-RowSubtotal.DP) + RowSubtotal.VAT
        '                            End If
        '                        Next
        '                    Else
        '                        RowSubtotal = TBLSubtotal.NewTBLDetail_SubtotalRow
        '                        RowSubtotal.Primary_Transaksi = RowCetak1.Primary
        '                        RowSubtotal.Primary_Product = RowCetak2.Primary_Product
        '                        RowSubtotal.Primary_Currency = RowCetak1.Primary_Currency
        '                        RowSubtotal.Subtotal += CDec(RowCetak2.Amount)
        '                        RowSubtotal.DP = -RowCetak1.Down_Payment
        '                        If RowCetak1.Tax_Exclude = True Then
        '                            RowSubtotal.VAT = ((RowSubtotal.Subtotal - RowCetak1.Down_Payment) * CDec(RowCetak1.Persentase)) / 100
        '                        Else
        '                            RowSubtotal.VAT = 0
        '                        End If
        '                        RowSubtotal.Total = RowSubtotal.Subtotal - (-RowSubtotal.DP) + RowSubtotal.VAT
        '                        RowSubtotal.Default_Precision = RowCetak1.Default_Precision
        '                        TBLSubtotal.Rows.Add(RowSubtotal)
        '                    End If
        '                    '------
        '                Next
        '            End If

        '            'Hitungan Subtotal Other Cost 
        '            TBLCetak3.DefaultView.RowFilter = Nothing
        '            TBLCetak3.DefaultView.RowFilter = "[Primary AOP] = " & RowCetak1.Primary & " AND [Primary Currency] = " & RowCetak1.Primary_Currency
        '            If TBLCetak3.DefaultView.Count > 0 Then
        '                For Each DV4 In TBLCetak3.DefaultView
        '                    RowCetak3 = DV4.Row
        '                    TBLSubtotal.DefaultView.RowFilter = Nothing
        '                    If PrimaryProduct <> -1 Then
        '                        If Me.ChkAllProductPO.Checked = False Then
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary & " AND [Primary Product] = " & PrimaryProduct & ""
        '                        Else
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '                        End If
        '                    Else
        '                        TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary
        '                    End If

        '                    If TBLSubtotal.DefaultView.Count > 0 Then
        '                        For Each DV3 In TBLSubtotal.DefaultView
        '                            RowSubtotal = DV3.Row
        '                            RowSubtotal.Other_Cost = RowCetak3.Other_Cost
        '                        Next
        '                    End If
        '                Next
        '            End If
        '            '------
        '        Next
        '    End If

        '    'Hitung Total Currency
        '    If TBLSubtotal.Rows.Count > 0 Then
        '        For Each RowSubtotal In TBLSubtotal.Rows
        '            With RowSubtotal
        '                TBLTotalCurr.DefaultView.RowFilter = "[Primary Currency] = " & .Primary_Currency
        '                If TBLTotalCurr.DefaultView.Count > 0 Then
        '                    RowTotalCurr = TBLTotalCurr.DefaultView(0).Row
        '                    RowTotalCurr.Total_SubTotal += .Subtotal
        '                    RowTotalCurr.Total_Down_Payment += .DP
        '                    RowTotalCurr.Total_VAT += .VAT
        '                    RowTotalCurr.Total_Amount += .Total
        '                    RowTotalCurr.Other_Cost_Amount += .Other_Cost
        '                Else
        '                    RowTotalCurr = TBLTotalCurr.NewTBLDetail_TotalRow
        '                    RowTotalCurr.Primary_AOP = .Primary_Transaksi
        '                    RowTotalCurr.Primary_Currency = .Primary_Currency
        '                    RowTotalCurr.Primary_Product = .Primary_Product
        '                    RowTotalCurr.Total_Description = ObjCurr.GetCurrencyCode(.Primary_Currency)
        '                    RowTotalCurr.Other_Cost_Desc = ObjCurr.GetCurrencyCode(.Primary_Currency)
        '                    RowTotalCurr.Total_SubTotal = .Subtotal
        '                    RowTotalCurr.Total_Down_Payment = .DP
        '                    RowTotalCurr.Total_VAT = .VAT
        '                    RowTotalCurr.Total_Amount = .Total
        '                    RowTotalCurr.Other_Cost_Amount = .Other_Cost
        '                    RowTotalCurr.Default_Precision = .Default_Precision
        '                    TBLTotalCurr.Rows.Add(RowTotalCurr)
        '                End If
        '            End With
        '        Next
        '    End If
        '    '------

        '    TBLCetak1.DefaultView.RowFilter = Nothing
        '    TBLCetak1.DefaultView.RowFilter = filter
        '    If TBLCetak1.DefaultView.Count > 0 Then
        '        For Each DV1 In TBLCetak1.DefaultView
        '            RowCetak1 = DV1.Row
        '            If PrimaryProduct <> -1 Then
        '                If Me.ChkAllProductPO.Checked = True Then
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString & " AND [Primary Product] = " & PrimaryProduct & ""
        '                Else
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString & " AND [Primary Product] = " & PrimaryProduct & ""
        '                End If
        '            Else
        '                filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '            End If
        '            TBLCetak2.DefaultView.RowFilter = Nothing
        '            TBLCetak2.DefaultView.RowFilter = filter2
        '            If TBLCetak2.DefaultView.Count > 0 Then
        '                For Each DV2 In TBLCetak2.DefaultView
        '                    RowCetak2 = DV2.Row
        '                    'isi tblCetakDetail dengan tblCetak1 dan tblCetak2 hasil filteran disini --------------------------------------------------------------------------
        '                    RowCetakRekap = TBLCetakRekap.NewTblCetak_RekapRow
        '                    RowCetakRekap.Primary_Transaksi = RowCetak1.Primary
        '                    RowCetakRekap.Curr_Code = RowCetak1.Currency_Code
        '                    RowCetakRekap.Curr_Description = RowCetak1.Currency_Description
        '                    If Me.ChkUsingSupplierNumber.Checked = True Then
        '                        RowCetakRekap.Invoice_No = RowCetak1.Invoice_No
        '                        If RowCetak1.Invoice_No = "" Then
        '                            RowCetakRekap.Invoice_No = RowCetak1.AOP_Number
        '                        End If
        '                    Else
        '                        RowCetakRekap.Invoice_No = RowCetak1.AOP_Number
        '                    End If
        '                    RowCetakRekap._Date = RowCetak1.Tanggal_Transaksi
        '                    RowCetakRekap.Supplier_Desc = RowCetak1.Supplier_Name
        '                    If IsNumeric(RowCetak1.Payment_Term) Then
        '                        RowCetakRekap.Payment_Term = RowCetak1.Payment_Term & " DAYS"
        '                    Else
        '                        RowCetakRekap.Payment_Term = RowCetak1.Payment_Term
        '                    End If
        '                    RowCetakRekap.Default_Precision = RowCetak1.Default_Precision
        '                    TBLCetakRekap.Rows.Add(RowCetakRekap)
        '                Next
        '            End If
        '        Next
        '    End If

        '    ReportRekap.SetDataSource(DS)
        '    ReportRekap.SetParameterValue(ReportRekap.Parameter_startDate.ParameterFieldName, paramStartDate)
        '    ReportRekap.SetParameterValue(ReportRekap.Parameter_endDate.ParameterFieldName, paramEndDate)
        '    Crv.ReportSource = ReportRekap
        '    Crv.Show()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    DS.Dispose()
        '    DTACetak1.Dispose()
        '    DTACetak2.Dispose()
        '    DTACetak3.Dispose()
        '    TBLCetakRekap.Dispose()
        '    TBLCetak1.Dispose()
        '    TBLCetak2.Dispose()
        '    TBLCetak3.Dispose()
        '    TBLSubtotal.Dispose()
        '    TBLTotalCurr.Dispose()
        'End Try
    End Sub

    Private Sub PrintDetail(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimarySupplier As Double, ByVal PrimaryAccount As Double)
        Dim DS As New DseLaporanPurchaseReport_Rekap
        Dim ReportDetail As New CrpLaporanPurcahseReport_Detail
        Dim DTACetak1 As New DseLaporanPurchaseReport_RekapTableAdapters.TblDetailTableAdapter
        Dim DTACetak3 As New DseLaporanPurchaseReport_RekapTableAdapters.tbltrans_aop_itemcostTableAdapter
        Dim TBLCetak1 As New DseLaporanPurchaseReport_Rekap.TblDetailDataTable
        Dim TBLCetak3 As New DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostDataTable
        Dim TBLCetakDetail As DseLaporanPurchaseReport_Rekap.TblCetak_DetailDataTable = DS.TblCetak_Detail
        Dim TBLTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalDataTable = DS.TBLDetail_Total
        Dim RowCetak1, RowCetak1_Aft As DseLaporanPurchaseReport_Rekap.TblDetailRow
        Dim RowCetak3 As DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostRow
        Dim RowCetakDetail As DseLaporanPurchaseReport_Rekap.TblCetak_DetailRow
        Dim RowTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalRow

        Dim paramStartDate As Date = Now
        Dim paramEndDate As Date = Now
        Dim primaryMax As Double = -1
        Dim primaryMin As Double = -1
        Dim filter As String = ""
        Dim HitungSubTotal As Decimal = 0
        Dim FilterGroup As String = ""
        Dim DecMulti As Decimal = 0
        Try
            DTACetak1.Connection.ConnectionString = Core_Module.GetConstring
            DTACetak3.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            paramStartDate = Me.DtpFrom.Value
            paramEndDate = Me.DtpTo.Value

            'Isi tblCetak ----------------------------------------------------------------------------------------------------------------
            DTACetak1.FillBy_StartEndDate(TBLCetak1, paramStartDate, paramEndDate)
            DTACetak3.FillBy_StartEndDate(TBLCetak3, paramEndDate, paramStartDate)

            'Filter tblCetak1 ------------------------------------------------------------------------------------------------------
            filter = "[Primary] = [Primary]"
            'If PrimaryProduct <> -1 Then
            '    If Me.ChkAllProductPO.Checked = False Then
            '        filter &= "AND [Primary Product] = " & PrimaryProduct & ""
            '    Else
            '        Me.GetAllProductPO(TBLCetak1, PrimaryProduct)
            '    End If
            'End If
            If PrimarySupplier <> -1 Then
                filter &= " AND [Primary Supplier] = " & PrimarySupplier & ""
            End If
            If PrimaryAccount <> -1 Then
                filter &= " AND [primary akun ap] = " & PrimaryAccount & ""
            End If

            FilterGroup = Me.GetFilterProduct(TBLCetak1)
            If FilterGroup <> "" Then
                filter &= " AND " & FilterGroup
            End If

            TBLCetak1.DefaultView.RowFilter = Nothing
            TBLCetak1.DefaultView.Sort = Nothing
            TBLCetak1.DefaultView.RowFilter = filter
            If Me.ChkUsingSupplierNumber.Checked = False Then
                TBLCetak1.DefaultView.Sort = "[Currency Code], [Tanggal Transaksi], [AOP Number], [Product Code]"
            End If
            If TBLCetak1.DefaultView.Count > 0 Then
                For idx As Integer = 0 To TBLCetak1.DefaultView.Count - 1
                    RowCetak1 = TBLCetak1.DefaultView(idx).Row
                    With RowCetak1
                        RowCetakDetail = TBLCetakDetail.NewTblCetak_DetailRow
                        RowCetakDetail.Primary_Transaksi = .Primary
                        RowCetakDetail.Curr_Code = .Currency_Code
                        RowCetakDetail.Curr_Description = .Currency_Description
                        RowCetakDetail.AOP_No = .AOP_Number
                        If Me.ChkUsingSupplierNumber.Checked = True Then
                            If .Invoice_No <> "" Then
                                RowCetakDetail.Invoice_No = .Invoice_No
                            Else
                                RowCetakDetail.Invoice_No = .AOP_Number
                            End If
                        Else
                            RowCetakDetail.Invoice_No = .AOP_Number
                        End If
                        RowCetakDetail._Date = .Tanggal_Transaksi
                        RowCetakDetail.Supplier_Desc = .Supplier_Name
                        If IsNumeric(.Payment_Term) Then
                            RowCetakDetail.Payment_term = .Payment_Term & " DAYS"
                        Else
                            RowCetakDetail.Payment_term = .Payment_Term
                        End If
                        RowCetakDetail.Default_Precision = .Default_Precision

                        '----Data Item 
                        RowCetakDetail.DO_Number = .DO_Number
                        RowCetakDetail.PO_Number = .PO_Number
                        RowCetakDetail.Product_Code = .Product_Code
                        RowCetakDetail.Product_Description = .Product_Description

                        If .Primary_Units <> .Primary_Units_Multi And .price_by_multi_unit = 0 Then
                            RowCetakDetail.Qty = .Total_Qty_Default
                            RowCetakDetail.Unit = .Unit_Code_Default
                            DecMulti = .Decimal_Multi
                            RowCetakDetail.Decimal_Multi = ._Decimal
                        Else
                            RowCetakDetail.Qty = .Total_Qty
                            RowCetakDetail.Unit = .Unit_Code
                            RowCetakDetail.Decimal_Multi = .Decimal_Multi
                        End If

                        RowCetakDetail.Unit_Price = .Nett_Price
                        RowCetakDetail.Amount = .Amount

                        HitungSubTotal += .Amount

                        Dim ApakahHitung As Boolean = True
                        If idx < TBLCetak1.DefaultView.Count - 1 Then
                            RowCetak1_Aft = TBLCetak1.DefaultView(idx + 1).Row
                            If RowCetakDetail.Primary_Transaksi = RowCetak1_Aft.Primary Then
                                ApakahHitung = False
                            End If
                        End If

                        If ApakahHitung = True Then
                            RowCetakDetail.SubTotal = HitungSubTotal
                            RowCetakDetail.Down_Payment = .Down_Payment
                            If .Tax_Exclude = True Then
                                RowCetakDetail.VAT = ((HitungSubTotal - .Down_Payment) * .Persentase) / 100
                            Else
                                RowCetakDetail.VAT = 0
                            End If
                            RowCetakDetail.Total_Amount = HitungSubTotal - .Down_Payment + RowCetakDetail.VAT

                            'Hitungan Subtotal Other Cost 
                            TBLCetak3.DefaultView.RowFilter = Nothing
                            TBLCetak3.DefaultView.RowFilter = "[Primary AOP] = " & RowCetakDetail.Primary_Transaksi & " AND [Primary Currency] = " & .Primary_Currency
                            If TBLCetak3.DefaultView.Count > 0 Then
                                RowCetak3 = TBLCetak3.DefaultView(0).Row
                                RowCetakDetail.Other_Cost = RowCetak3.Other_Cost
                            End If

                            'Hitung Total Currency
                            TBLTotalCurr.DefaultView.RowFilter = "[Primary Currency] = " & .Primary_Currency
                            If TBLTotalCurr.DefaultView.Count > 0 Then
                                RowTotalCurr = TBLTotalCurr.DefaultView(0).Row
                                RowTotalCurr.Total_VAT += RowCetakDetail.VAT
                                RowTotalCurr.Total_Amount += RowCetakDetail.Total_Amount
                                RowTotalCurr.Total_SubTotal += RowCetakDetail.SubTotal
                                RowTotalCurr.Total_Down_Payment += RowCetakDetail.Down_Payment
                                RowTotalCurr.Other_Cost_Amount += RowCetakDetail.Other_Cost
                            Else
                                RowTotalCurr = TBLTotalCurr.NewTBLDetail_TotalRow
                                RowTotalCurr.Primary_AOP = RowCetakDetail.Primary_Transaksi
                                RowTotalCurr.Primary_Currency = .Primary_Currency
                                RowTotalCurr.Total_Description = ObjCurr.GetCurrencyCode(.Primary_Currency)
                                RowTotalCurr.Other_Cost_Desc = ObjCurr.GetCurrencyCode(.Primary_Currency)
                                RowTotalCurr.Total_VAT = RowCetakDetail.VAT
                                RowTotalCurr.Total_Amount = RowCetakDetail.Total_Amount
                                RowTotalCurr.Total_SubTotal = RowCetakDetail.SubTotal
                                RowTotalCurr.Total_Down_Payment = RowCetakDetail.Down_Payment
                                RowTotalCurr.Other_Cost_Amount = RowCetakDetail.Other_Cost
                                RowTotalCurr.Default_Precision = ObjCurr.GetDefaultPrecission(.Primary_Currency)
                                TBLTotalCurr.Rows.Add(RowTotalCurr)
                            End If

                            HitungSubTotal = 0
                        End If

                        RowCetakDetail._Decimal = ._Decimal

                        If Me.ChkShowRatio.Checked = True And .Primary_Units_Multi <> .Primary_Units And .price_by_multi_unit = 1 Then
                            RowCetakDetail.Ratio_Desc = "(" & ObjCd.DigitGrouping(.Ratio_Reference, RowCetakDetail.Decimal_Multi) & " " & RowCetakDetail.Unit & "=" & _
                                                          ObjCd.DigitGrouping(.Ratio_Default, RowCetakDetail._Decimal) & " " & .Unit_Description_Default & ")"
                        Else
                            RowCetakDetail.Ratio_Desc = ""
                        End If

                        If .Primary_Units_Multi <> .Primary_Units And .price_by_multi_unit = 0 Then
                            RowCetakDetail.Ratio_Desc = "(" & ObjCd.DigitGrouping(.Total_Qty, DecMulti) & " " & .Unit_Code & ")"
                        End If

                        TBLCetakDetail.Rows.Add(RowCetakDetail)
                    End With
                Next
            End If

            ReportDetail.SetDataSource(DS)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
            ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
            Crv.ReportSource = ReportDetail
            Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTACetak1.Dispose()
            DTACetak3.Dispose()
            TBLCetak1.Dispose()
            TBLCetak3.Dispose()
            TBLCetakDetail.Dispose()
            TBLTotalCurr.Dispose()
        End Try
    End Sub

    Private Sub PrintDetail_Obs(ByVal EndDate As Date, ByVal StartDate As Date, ByVal PrimaryProduct As Double, ByVal PrimarySupplier As Double, ByVal PrimaryAccount As Double)
        'Dim DS As New DseLaporanPurchaseReport_Rekap
        'Dim Report As New CrpLaporanPurcahseReport_Detail

        'Dim DTACetak1 As New DseLaporanPurchaseReport_RekapTableAdapters.TblDetail_Header1TableAdapter
        'Dim DTACetak2 As New DseLaporanPurchaseReport_RekapTableAdapters.TBLDetail_ItemTableAdapter
        'Dim DTACetak3 As New DseLaporanPurchaseReport_RekapTableAdapters.tbltrans_aop_itemcostTableAdapter
        'Dim TBLCetakDetail As DseLaporanPurchaseReport_Rekap.TblCetak_DetailDataTable = DS.TblCetak_Detail
        'Dim TBLCetak1 As DseLaporanPurchaseReport_Rekap.TblDetail_Header1DataTable = DS.TblDetail_Header1
        'Dim TBLCetak2 As DseLaporanPurchaseReport_Rekap.TBLDetail_ItemDataTable = DS.TBLDetail_Item
        'Dim TBLCetak3 As DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostDataTable = DS.tbltrans_aop_itemcost
        'Dim TBLSubtotal As DseLaporanPurchaseReport_Rekap.TBLDetail_SubtotalDataTable = DS.TBLDetail_Subtotal
        'Dim TBLTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalDataTable = DS.TBLDetail_Total
        'Dim RowCetakDetail As DseLaporanPurchaseReport_Rekap.TblCetak_DetailRow
        'Dim RowCetak1 As DseLaporanPurchaseReport_Rekap.TblDetail_Header1Row
        'Dim RowCetak2 As DseLaporanPurchaseReport_Rekap.TBLDetail_ItemRow
        'Dim RowCetak3 As DseLaporanPurchaseReport_Rekap.tbltrans_aop_itemcostRow
        'Dim RowSubtotal As DseLaporanPurchaseReport_Rekap.TBLDetail_SubtotalRow
        'Dim RowTotalCurr As DseLaporanPurchaseReport_Rekap.TBLDetail_TotalRow
        'Dim DV1 As DataRowView
        'Dim DV2 As DataRowView
        'Dim DV3 As DataRowView
        'Dim DV4 As DataRowView

        'Dim paramStartDate As Date = Now
        'Dim paramEndDate As Date = Now
        'Dim primaryMax As Double = -1
        'Dim primaryMin As Double = -1
        'Dim filter As String = ""
        'Dim filter2 As String = ""
        'Try
        '    DTACetak1.Connection.ConnectionString = Core_Module.GetConstring
        '    DTACetak2.Connection.ConnectionString = Core_Module.GetConstring
        '    DTACetak3.Connection.ConnectionString = Core_Module.GetConstring
        '    Windows.Forms.Cursor.Current = Cursors.WaitCursor

        '    paramStartDate = Me.DtpFrom.Value
        '    paramEndDate = Me.DtpTo.Value

        '    'Isi tblCetak ----------------------------------------------------------------------------------------------------------------
        '    If Me.ChkAllProductPO.Checked = True Then
        '        DTACetak1.FillBy_AllProductTransaksi(TBLCetak1, paramStartDate, paramEndDate, PrimaryProduct)
        '    Else
        '        DTACetak1.FillBy_StartEndDate(TBLCetak1, paramStartDate, paramEndDate)
        '    End If

        '    'DTAProduct.Fill(TblProduct, PrimaryProduct)
        '    'TBLCetak1.DefaultView.Sort = Nothing
        '    'TBLCetak1.DefaultView.Sort = "[Primary]"
        '    'If TBLCetak1.DefaultView.Count > 0 Then
        '    '    RowCetak1 = TBLCetak1.DefaultView(0).Row
        '    '    primaryMin = RowCetak1.Primary
        '    '    RowCetak1 = TBLCetak1.DefaultView(TBLCetak1.DefaultView.Count - 1).Row
        '    '    primaryMax = RowCetak1.Primary
        '    'End If
        '    DTACetak2.Fill(TBLCetak2)
        '    DTACetak3.Fill(TBLCetak3)

        '    'Filter tblCetak1 ------------------------------------------------------------------------------------------------------
        '    filter = "[Primary] = [Primary]"
        '    If PrimarySupplier <> -1 Then
        '        filter &= " AND [Primary Supplier] = " & PrimarySupplier & ""
        '    End If
        '    If PrimaryAccount <> -1 Then
        '        filter &= " AND [primary akun ap] = " & PrimaryAccount & ""
        '    End If

        '    TBLCetak1.DefaultView.RowFilter = Nothing
        '    TBLCetak1.DefaultView.RowFilter = filter
        '    If TBLCetak1.DefaultView.Count > 0 Then
        '        For Each DV1 In TBLCetak1.DefaultView
        '            RowCetak1 = DV1.Row
        '            If PrimaryProduct <> -1 Then
        '                If Me.ChkAllProductPO.Checked = False Then
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString & " AND [Primary Product] = " & PrimaryProduct & ""
        '                Else
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '                End If
        '            Else
        '                filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '            End If

        '            TBLCetak2.DefaultView.RowFilter = Nothing
        '            TBLCetak2.DefaultView.RowFilter = filter2
        '            If TBLCetak2.DefaultView.Count > 0 Then
        '                For Each DV2 In TBLCetak2.DefaultView
        '                    RowCetak2 = DV2.Row
        '                    'Hitung Subtotal
        '                    TBLSubtotal.DefaultView.RowFilter = Nothing
        '                    If PrimaryProduct <> -1 Then
        '                        If Me.ChkAllProductPO.Checked = False Then
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary & " AND [Primary Product] = " & PrimaryProduct & ""
        '                        Else
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '                        End If
        '                    Else
        '                        TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary
        '                    End If

        '                    If TBLSubtotal.DefaultView.Count > 0 Then
        '                        For Each DV3 In TBLSubtotal.DefaultView
        '                            RowSubtotal = DV3.Row
        '                            If IsNumeric(RowCetak2.Amount) Then
        '                                RowSubtotal.Subtotal += CDec(RowCetak2.Amount)
        '                                RowSubtotal.Primary_Product = RowCetak2.Primary_Product
        '                                RowSubtotal.DP = -RowCetak1.Down_Payment
        '                                If RowCetak1.Tax_Exclude = True Then
        '                                    RowSubtotal.VAT = ((RowSubtotal.Subtotal - RowCetak1.Down_Payment) * CDec(RowCetak1.Persentase)) / 100
        '                                Else
        '                                    RowSubtotal.VAT = 0
        '                                End If
        '                                RowSubtotal.Total = RowSubtotal.Subtotal - (-RowSubtotal.DP) + RowSubtotal.VAT
        '                            End If
        '                        Next
        '                    Else
        '                        RowSubtotal = TBLSubtotal.NewTBLDetail_SubtotalRow
        '                        RowSubtotal.Primary_Transaksi = RowCetak1.Primary
        '                        RowSubtotal.Primary_Product = RowCetak2.Primary_Product
        '                        RowSubtotal.Primary_Currency = RowCetak1.Primary_Currency
        '                        RowSubtotal.Subtotal += CDec(RowCetak2.Amount)
        '                        RowSubtotal.DP = -RowCetak1.Down_Payment
        '                        If RowCetak1.Tax_Exclude = True Then
        '                            RowSubtotal.VAT = ((RowSubtotal.Subtotal - RowCetak1.Down_Payment) * CDec(RowCetak1.Persentase)) / 100
        '                        Else
        '                            RowSubtotal.VAT = 0
        '                        End If
        '                        RowSubtotal.Total = RowSubtotal.Subtotal - (-RowSubtotal.DP) + RowSubtotal.VAT
        '                        RowSubtotal.Default_Precision = RowCetak1.Default_Precision
        '                        TBLSubtotal.Rows.Add(RowSubtotal)
        '                    End If
        '                    '------
        '                Next
        '            End If

        '            'Hitungan Subtotal Other Cost 
        '            TBLCetak3.DefaultView.RowFilter = Nothing
        '            TBLCetak3.DefaultView.RowFilter = "[Primary AOP] = " & RowCetak1.Primary & " AND [Primary Currency] = " & RowCetak1.Primary_Currency
        '            If TBLCetak3.DefaultView.Count > 0 Then
        '                For Each DV4 In TBLCetak3.DefaultView
        '                    RowCetak3 = DV4.Row
        '                    TBLSubtotal.DefaultView.RowFilter = Nothing
        '                    If PrimaryProduct <> -1 Then
        '                        If Me.ChkAllProductPO.Checked = False Then
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary & " AND [Primary Product] = " & PrimaryProduct & ""
        '                        Else
        '                            TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary

        '                        End If
        '                    Else
        '                        TBLSubtotal.DefaultView.RowFilter = "[Primary Transaksi] = " & RowCetak1.Primary
        '                    End If
        '                    If TBLSubtotal.DefaultView.Count > 0 Then
        '                        For Each DV3 In TBLSubtotal.DefaultView
        '                            RowSubtotal = DV3.Row
        '                            RowSubtotal.Other_Cost = RowCetak3.Other_Cost
        '                        Next
        '                    End If
        '                Next
        '            End If
        '            '------
        '        Next
        '    End If

        '    'Hitung Total Currency
        '    If TBLSubtotal.Rows.Count > 0 Then
        '        For Each RowSubtotal In TBLSubtotal.Rows
        '            With RowSubtotal
        '                TBLTotalCurr.DefaultView.RowFilter = "[Primary Currency] = " & .Primary_Currency
        '                If TBLTotalCurr.DefaultView.Count > 0 Then
        '                    RowTotalCurr = TBLTotalCurr.DefaultView(0).Row
        '                    RowTotalCurr.Total_VAT += .VAT
        '                    RowTotalCurr.Total_Amount += .Total
        '                    RowTotalCurr.Total_SubTotal += .Subtotal
        '                    RowTotalCurr.Total_Down_Payment += .DP
        '                    RowTotalCurr.Other_Cost_Amount += .Other_Cost
        '                Else
        '                    RowTotalCurr = TBLTotalCurr.NewTBLDetail_TotalRow
        '                    RowTotalCurr.Primary_AOP = .Primary_Transaksi
        '                    RowTotalCurr.Primary_Currency = .Primary_Currency
        '                    RowTotalCurr.Primary_Product = .Primary_Product
        '                    RowTotalCurr.Total_Description = ObjCurr.GetCurrencyCode(.Primary_Currency)
        '                    RowTotalCurr.Other_Cost_Desc = ObjCurr.GetCurrencyCode(.Primary_Currency)
        '                    RowTotalCurr.Total_VAT = .VAT
        '                    RowTotalCurr.Total_Amount = .Total
        '                    RowTotalCurr.Total_SubTotal = .Subtotal
        '                    RowTotalCurr.Total_Down_Payment = .DP
        '                    RowTotalCurr.Other_Cost_Amount = .Other_Cost
        '                    RowTotalCurr.Default_Precision = .Default_Precision
        '                    TBLTotalCurr.Rows.Add(RowTotalCurr)
        '                End If
        '            End With
        '        Next
        '    End If
        '    '------

        '    TBLCetak1.DefaultView.RowFilter = Nothing
        '    TBLCetak1.DefaultView.RowFilter = filter
        '    If TBLCetak1.DefaultView.Count > 0 Then
        '        For Each DV1 In TBLCetak1.DefaultView
        '            RowCetak1 = DV1.Row
        '            If PrimaryProduct <> -1 Then
        '                If Me.ChkAllProductPO.Checked = False Then
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString & " AND [Primary Product] = " & PrimaryProduct & ""

        '                Else
        '                    filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '                End If
        '            Else
        '                filter2 = "[Primary Transaksi] = " & RowCetak1.Primary.ToString
        '            End If
        '            TBLCetak2.DefaultView.RowFilter = Nothing
        '            TBLCetak2.DefaultView.RowFilter = filter2
        '            If TBLCetak2.DefaultView.Count > 0 Then
        '                For Each DV2 In TBLCetak2.DefaultView
        '                    RowCetak2 = DV2.Row
        '                    'isi tblCetakDetail dengan tblCetak1 dan tblCetak2 hasil filteran disini --------------------------------------------------------------------------
        '                    RowCetakDetail = TBLCetakDetail.NewTblCetak_DetailRow
        '                    RowCetakDetail.Primary_Transaksi = RowCetak1.Primary
        '                    RowCetakDetail.Curr_Code = RowCetak1.Currency_Code
        '                    RowCetakDetail.Curr_Description = RowCetak1.Currency_Description
        '                    If Me.ChkUsingSupplierNumber.Checked = True Then
        '                        RowCetakDetail.Invoice_No = RowCetak1.Invoice_No
        '                        If RowCetak1.Invoice_No = "" Then
        '                            RowCetakDetail.Invoice_No = RowCetak1.AOP_Number
        '                        End If
        '                    Else
        '                        RowCetakDetail.Invoice_No = RowCetak1.AOP_Number
        '                    End If
        '                    RowCetakDetail._Date = RowCetak1.Tanggal_Transaksi
        '                    RowCetakDetail.Supplier_Desc = RowCetak1.Supplier_Name
        '                    If IsNumeric(RowCetak1.Payment_Term) Then
        '                        RowCetakDetail.Payment_term = RowCetak1.Payment_Term & " DAYS"
        '                    Else
        '                        RowCetakDetail.Payment_term = RowCetak1.Payment_Term
        '                    End If
        '                    RowCetakDetail.Default_Precision = RowCetak1.Default_Precision
        '                    '----Data Item 
        '                    RowCetakDetail.DO_Number = RowCetak2.DO_Number
        '                    RowCetakDetail.PO_Number = RowCetak2.PO_Number
        '                    RowCetakDetail.Primary_Product = RowCetak2.Primary_Product

        '                    RowCetakDetail.Product_Code = RowCetak2.Product_Code
        '                    RowCetakDetail.Product_Description = RowCetak2.Product_Description
        '                    RowCetakDetail.Qty = RowCetak2.Total_Qty
        '                    RowCetakDetail.Unit = RowCetak2.Unit_Code
        '                    RowCetakDetail.Unit_Price = RowCetak2.Nett_Price
        '                    RowCetakDetail.Amount = RowCetak2.Amount
        '                    RowCetakDetail.SubTotal = 0
        '                    RowCetakDetail.Down_Payment = 0
        '                    RowCetakDetail.VAT = 0
        '                    RowCetakDetail.Total_Price = 0
        '                    RowCetakDetail.Other_Cost = 0

        '                    'aul
        '                    RowCetakDetail.Primary_Units_Multi = RowCetak2.Primary_Units_Multi
        '                    RowCetakDetail.Primary_Unit_Default = RowCetak2.Primary_Units_Default
        '                    RowCetakDetail.Unit_Default = RowCetak2.Unit_Code_Default
        '                    RowCetakDetail.Ratio_Reference = RowCetak2.Ratio_Reference
        '                    RowCetakDetail.Ratio_Default = RowCetak2.Ratio_Default
        '                    RowCetakDetail._Decimal = RowCetak2._Decimal
        '                    If Me.ChkShowRatio.Checked = True And RowCetakDetail.Primary_Units_Multi <> RowCetakDetail.Primary_Unit_Default Then
        '                        RowCetakDetail.Decimal_Multi = ObjProduct.GetAngkaDecimal(RowCetakDetail.Primary_Product, RowCetakDetail.Primary_Units_Multi)
        '                        RowCetakDetail.Ratio_Desc = "(" & ObjCd.DigitGrouping(RowCetakDetail.Ratio_Reference, RowCetakDetail.Decimal_Multi) & " " & RowCetakDetail.Unit & "=" & _
        '                        ObjCd.DigitGrouping(RowCetakDetail.Ratio_Default, RowCetakDetail._Decimal) & " " & RowCetakDetail.Unit_Default & ")"
        '                    Else
        '                        RowCetakDetail.Ratio_Desc = ""
        '                    End If
        '                    If RowCetakDetail.Primary_Unit_Default <> RowCetakDetail.Primary_Units_Multi Then
        '                        If Me.ChkShowRatio.Checked = False Then
        '                            RowCetakDetail.Decimal_Multi = ObjProduct.GetAngkaDecimal(RowCetakDetail.Primary_Product, RowCetakDetail.Primary_Units_Multi)
        '                        End If
        '                        RowCetakDetail.Decimal_Multi = RowCetakDetail.Decimal_Multi
        '                    Else
        '                        RowCetakDetail.Decimal_Multi = RowCetakDetail._Decimal
        '                    End If
        '                    '---
        '                    TBLCetakDetail.Rows.Add(RowCetakDetail)
        '                Next
        '            End If
        '        Next
        '    End If

        '    ReportDetail.SetDataSource(DS)
        '    ReportDetail.SetParameterValue(ReportDetail.Parameter_StartDate.ParameterFieldName, paramStartDate)
        '    ReportDetail.SetParameterValue(ReportDetail.Parameter_EndDate.ParameterFieldName, paramEndDate)
        '    Crv.ReportSource = ReportDetail
        '    Crv.Show()
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    DS.Dispose()
        '    DTACetak1.Dispose()
        '    DTACetak2.Dispose()
        '    DTACetak3.Dispose()
        '    TBLCetakDetail.Dispose()
        '    TBLCetak1.Dispose()
        '    TBLCetak2.Dispose()
        '    TBLCetak3.Dispose()
        '    TBLSubtotal.Dispose()
        '    TBLTotalCurr.Dispose()
        'End Try
    End Sub

    Private Sub txtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeNo.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.txtCodeNo.Text <> "" Then
                    If Me.ObjProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                        Me.txtPrimaryCodeNo.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text))
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                        Me.TxtCodeNoDesc.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryCodeNo.Text))
                        Me.TxtKodeSupplier.Focus()
                    Else
                        Me.BtnFindCodeNo.PerformClick()
                    End If
                Else
                    Me.BtnFindCodeNo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Leave
        Dim hj As New Core_Module.frmCariMasterProduct
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.txtCodeNo.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                    Me.txtPrimaryCodeNo.Text = CStr(Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text))
                    Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                    Me.TxtCodeNoDesc.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryCodeNo.Text))
                Else
                    If Me.txtCodeNo.Text <> "" Then
                        hj.txtCari.Text = Me.txtCodeNo.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = CStr(hj.PrimaryProduct)
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                        Me.TxtCodeNoDesc.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryCodeNo.Text))
                        Me.TxtKodeSupplier.Focus()
                    Else
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.TxtCodeNoDesc.Clear()
                        Me.txtCodeNo.Clear()
                        Me.txtCodeNo.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.TxtCodeNoDesc.Clear()
                Me.txtCodeNo.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim hj As New Core_Module.frmCariMasterProduct
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.txtCodeNo.Text <> "" Then
                hj.txtCari.Text = Me.txtCodeNo.Text
            End If

            hj.ShowDialog()
            If hj.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = CStr(hj.PrimaryProduct)
                Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(CDbl(Me.txtPrimaryCodeNo.Text))
                Me.TxtCodeNoDesc.Text = Me.ObjProduct.GetDescription(CDbl(Me.txtPrimaryCodeNo.Text))
                Me.TxtKodeSupplier.Focus()
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.TxtCodeNoDesc.Clear()
                Me.txtCodeNo.Clear()
                Me.txtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtKodeSupplier_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtKodeSupplier.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtKodeSupplier.Text <> "" Then
                    If Me.ObjSupplier.ApakahSudahAda(Me.TxtKodeSupplier.Text) = True Then
                        Me.TxtPrimarySupplier.Text = CStr(Me.ObjSupplier.GetPrimary(Me.TxtKodeSupplier.Text))
                        Me.TxtKodeSupplier.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text))
                        Me.TxtSuppliserDesc.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.TxtPrimarySupplier.Text))
                        Me.TxtAccCode.Focus()
                    Else
                        Me.BtnFindSupp.PerformClick()
                    End If
                Else
                    Me.BtnFindSupp.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtKodeSupplier_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtKodeSupplier.Leave
        Dim hj As New Core_Module.FrmCariMasterSupplier
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtKodeSupplier.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.TxtKodeSupplier.Text) = True Then
                    Me.TxtPrimarySupplier.Text = CStr(Me.ObjSupplier.GetPrimary(Me.TxtKodeSupplier.Text))
                    Me.TxtKodeSupplier.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text))
                    Me.TxtSuppliserDesc.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.TxtPrimarySupplier.Text))
                Else
                    If Me.TxtKodeSupplier.Text <> "" Then
                        hj.txtCari.Text = Me.TxtKodeSupplier.Text
                    End If
                    hj.ShowDialog()

                    If hj.PrimarySupplier > 0 Then
                        Me.TxtPrimarySupplier.Text = CStr(hj.PrimarySupplier)
                        Me.TxtKodeSupplier.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text))
                        Me.TxtSuppliserDesc.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.TxtPrimarySupplier.Text))
                        Me.TxtAccCode.Focus()
                    Else
                        Me.TxtPrimarySupplier.Text = "-1"
                        Me.TxtSuppliserDesc.Clear()
                        Me.TxtKodeSupplier.Clear()
                        Me.TxtKodeSupplier.Focus()
                    End If
                End If
            Else
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSuppliserDesc.Clear()
                Me.TxtKodeSupplier.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindSupp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindSupp.Click
        Dim hj As New Core_Module.FrmCariMasterSupplier
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.TxtKodeSupplier.Text <> "" Then
                hj.txtCari.Text = Me.TxtKodeSupplier.Text
            End If

            hj.ShowDialog()
            If hj.PrimarySupplier > 0 Then
                Me.TxtPrimarySupplier.Text = CStr(hj.PrimarySupplier)
                Me.TxtKodeSupplier.Text = Me.ObjSupplier.GetSupplierCode(CDbl(Me.TxtPrimarySupplier.Text))
                Me.TxtSuppliserDesc.Text = Me.ObjSupplier.GetSupplierName(CDbl(Me.TxtPrimarySupplier.Text))
                Me.TxtAccCode.Focus()
            Else
                Me.TxtPrimarySupplier.Text = "-1"
                Me.TxtSuppliserDesc.Clear()
                Me.TxtKodeSupplier.Clear()
                Me.TxtKodeSupplier.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub TxtAccCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAccCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.TxtAccCode.Text <> "" Then
                    If Me.ObjAccount.ApakahSudahAda(Me.TxtAccCode.Text) = True Then
                        Me.TxtPrimaryAcc.Text = CStr(Me.ObjAccount.GetPrimary(Me.TxtAccCode.Text))
                        Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                        Me.TxtAccountDesc.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
                        Me.TxtAccCode.Focus()
                    Else
                        Me.BtnFindSupp.PerformClick()
                    End If
                Else
                    Me.BtnFindSupp.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtAccCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAccCode.Leave
        Dim hj As New Core_Module.FrmMasterAkun_Cari1
        'hj.BackColorAll = Me.BAckColorAll
        Try
            If Me.TxtAccCode.Text <> "" Then
                If Me.ObjAccount.ApakahSudahAda(Me.TxtAccCode.Text) = True Then
                    Me.TxtPrimaryAcc.Text = CStr(Me.ObjAccount.GetPrimary(Me.TxtAccCode.Text))
                    Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                    Me.TxtAccountDesc.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
                Else
                    If Me.TxtAccCode.Text <> "" Then
                        hj.txtCari.Text = Me.TxtAccCode.Text
                    End If
                    hj.ShowDialog()

                    If hj.mPrimaryAkun > 0 Then
                        Me.TxtPrimaryAcc.Text = CStr(hj.mPrimaryAkun)
                        Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                        Me.TxtAccountDesc.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
                        Me.TxtAccCode.Focus()
                    Else
                        Me.TxtPrimaryAcc.Text = "-1"
                        Me.TxtAccountDesc.Clear()
                        Me.TxtAccCode.Clear()
                        Me.TxtAccCode.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryAcc.Text = "-1"
                Me.TxtAccountDesc.Clear()
                Me.TxtAccCode.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub BtnFindAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindAcc.Click
        Dim hj As New Core_Module.FrmMasterAkun_Cari1
        Try
            'hj.BackColorAll = Me.BAckColorAll
            If Me.TxtAccCode.Text <> "" Then
                hj.txtCari.Text = Me.TxtAccCode.Text
            End If

            hj.ShowDialog()
            If hj.mPrimaryAkun > 0 Then
                Me.TxtPrimaryAcc.Text = CStr(hj.mPrimaryAkun)
                Me.TxtAccCode.Text = Me.ObjAccount.GetAccountCode(CDbl(Me.TxtPrimaryAcc.Text))
                Me.TxtAccountDesc.Text = Me.ObjAccount.GetDescription_ByPrimaryAcc(CDbl(Me.TxtPrimaryAcc.Text))
                Me.TxtAccCode.Focus()
            Else
                Me.TxtPrimaryAcc.Text = "-1"
                Me.TxtAccountDesc.Clear()
                Me.TxtAccCode.Clear()
                Me.TxtAccCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.GotFocus, TxtKodeSupplier.GotFocus, TxtAccCode.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnPrint.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtCodeNo.BackColor
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

    Private Sub txtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.LostFocus, TxtKodeSupplier.LostFocus, TxtAccCode.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.BtnPrint.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtCodeNo.BackColor
                        Else
                            Txt.BackColor = Me.txtPrimaryCodeNo.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCodeNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPrimaryCodeNo.TextChanged
        Try
            If Me.txtPrimaryCodeNo.Text = -1 Then
                Me.ChkAllProductPO.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkAllProductPO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAllProductPO.Click
        'If Me.txtPrimaryCodeNo.Text <> -1 Then

        'Else
        '    Me.ChkAllProductPO.Checked = False
        '    MsgBox("Product cannot be an empty string", MsgBoxStyle.Information, "Information")
        '    Me.txtCodeNo.Focus()
        'End If
        If Me.ApakahPakeFilterProduct = False Then
            Me.ChkAllProductPO.Checked = False
            MsgBox("Product cannot be an empty string", MsgBoxStyle.Information, "Information")
            Me.txtCodeNo.Focus()
        End If
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
                Me.ChkAllProductPO.Checked = False
            End If
            Me.ObjCd.SetNumber(Me.dgvItemProductGroup)
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

    Private Function GetFilterProduct(ByVal TBLCetak As DseLaporanPurchaseReport_Rekap.TblDetailDataTable) As String
        Dim rowDGV As New DataGridViewRow
        Dim FilterGroup As String = ""
        Dim Row As DseLaporanPurchaseReport_Rekap.TblDetailRow
        Dim ListPrimaryPO As New List(Of Double)
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
                If Me.ChkAllProductPO.Checked = True Then
                    TBLCetak.DefaultView.RowFilter = FilterGroup
                    For i As Integer = 0 To TBLCetak.DefaultView.Count - 1
                        Row = TBLCetak.DefaultView(i).Row
                        If ListPrimaryPO.Contains(Row.Primary) = False Then
                            ListPrimaryPO.Add(Row.Primary)
                        End If
                    Next

                    For Each LS As Double In ListPrimaryPO
                        If FilterSO = "" Then
                            FilterSO = "[Primary]=" & LS
                        Else
                            FilterSO &= " OR [Primary]= " & LS
                        End If
                    Next

                    If FilterSO <> "" Then
                        FilterGroup = "(" & FilterSO & ")"
                    End If
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
                    Me.dgvItemProductGroup.AllowUserToAddRows = True
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