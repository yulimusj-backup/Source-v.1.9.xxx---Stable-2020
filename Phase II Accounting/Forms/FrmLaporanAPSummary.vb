Option Explicit On
Option Strict Off

Public Class FrmLaporanAPSummary
    Public Usercode As String = ""
    Public BackColorAll As Color = Color.LightSeaGreen
    Dim DS As New DseLaporanAPSummary
    Dim DTA1 As New DseLaporanAPSummaryTableAdapters.balance_purchase_viewTableAdapter
    Dim DTA3 As New DseLaporanAPSummaryTableAdapters.tblmaster_supplierTableAdapter
    Dim ObjCD As New Core_Module.ClsCekData
    Dim report As New CrpLaporanAPSummary
    Dim report_Detail_IDR As New CrpLaporanAPSummary_DetailIDR
    Dim report_Detail_NonIDR As New CrpLaporanAPSummary_DetailNonIDR
    Dim reportRecapOriginal As New CrpLaporanAPSummary2
    Dim reportChart As New CrpLaporanChart_APSummary

    Dim TipeTrans1, TipeTrans2, TipeTrans3, TipeTrans4, TipeTrans5, TipeTrans6, TipeTrans7, TipeTrans8, TipeTrans9 As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
    Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP

    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim TipeLaporan As EnumFormat = EnumFormat.Detail_In_IDR
    Dim ObjLicense As New Core_Module.ClsLicense

    Enum EnumFormat As Integer
        Recap
        Recap_In_Original_Currency
        Detail_In_Original_Currency
        Detail_In_IDR
        RecapChart
    End Enum

    Private Sub cboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFormat.SelectedIndexChanged
        Try
            If Me.cboFormat.SelectedIndex = 0 Then
                Me.TipeLaporan = EnumFormat.Recap
            ElseIf Me.cboFormat.SelectedIndex = 1 Then
                Me.TipeLaporan = EnumFormat.Recap_In_Original_Currency
            ElseIf Me.cboFormat.SelectedIndex = 2 Then
                Me.TipeLaporan = EnumFormat.Detail_In_Original_Currency
            ElseIf Me.cboFormat.SelectedIndex = 3 Then
                Me.TipeLaporan = EnumFormat.Detail_In_IDR
            Else
                Me.TipeLaporan = EnumFormat.RecapChart
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmLaporanAPSummary_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.cboFormat.SelectedIndex = 0
            Me.TipeKartu = CInt(ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP).ToString

            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis
            Me.TipeTrans7 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            Me.TipeTrans8 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Me.TipeTrans9 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual

            Me.MinimizeBox = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateDetail_InIDR()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row_Supplier As DseLaporanAPSummary.tblmaster_supplierRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Dim TBLSaldoAwal As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TBLSales As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TBLSaldoAkhir As New DseLaporanAPSummary.TblPurchaseDataTable

        Try
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Incorrect period", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Me.crv1.Refresh()
                Exit Sub
            End If

            'isikan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_supplier)
            'isikan data beginning balance (invoice + SR)
            Me.IsiSaldo(Me.dtpFrom.Value.Date - TimeSpan.FromDays(1), TBLSaldoAwal)
            'isikan data new invoice
            Me.IsiSales(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, TBLSales)
            'isikan data ending balance
            Me.IsiSaldo(Me.dtpTo.Value.Date, TBLSaldoAkhir)
            'semua data yang ada masukkan ke dalam tabel customer
            Me.IsikanSaldoAwal_DetailIDR(Me.DS.tblmaster_supplier, TBLSaldoAwal)
            Me.IsikanSaldoAkhir_DetailIDR(Me.DS.tblmaster_supplier, TBLSaldoAkhir)
            Me.IsikanNewInvoice_DetailIDR(Me.DS.tblmaster_supplier, TBLSales)

            'hitung data payment
            For Each row_Supplier In Me.DS.tblmaster_supplier.Rows
                row_Supplier.Payments = Me.ObjCD.DigitGrouping(CDbl(row_Supplier.Beginning_Balance) + CDbl(row_Supplier.New_Purchase) - _
                CDbl(row_Supplier.Ending_Balance))
            Next

            'hilangkan data yang semua datanya nol
            SubTotal = 0
            While SubTotal < Me.DS.tblmaster_supplier.Rows.Count
                row_Supplier = Me.DS.tblmaster_supplier.Rows(SubTotal)
                If row_Supplier.Beginning_Balance = 0 And row_Supplier.New_Purchase = 0 And row_Supplier.Payments = 0 And _
                row_Supplier.Ending_Balance = 0 Then
                    Me.DS.tblmaster_supplier.Rows.Remove(row_Supplier)
                Else
                    SubTotal += 1
                End If
            End While

            'isikan no transaksi
            Me.IsikanNoTransaksi()

            'tampilkan laporan
            Me.report_Detail_IDR.SetDataSource(DS)
            Me.report_Detail_IDR.SetParameterValue(report_Detail_IDR.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.report_Detail_IDR.SetParameterValue(report_Detail_IDR.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
            Me.report_Detail_IDR.SetParameterValue(report_Detail_IDR.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_Detail_IDR.SetParameterValue(report_Detail_IDR.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.report_Detail_IDR
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBLSaldoAwal.Dispose()
            TBLSales.Dispose()
            TBLSaldoAkhir.Dispose()
        End Try
    End Sub

    Private Sub CreateDetail_OriginalCurrency()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row_Supplier As DseLaporanAPSummary.tblmaster_supplierRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Dim TBLSaldoAwal As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TBLSales As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TBLSaldoAkhir As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TblTotCurr As DseLaporanAPSummary.Total_CurrDataTable = DS.Total_Curr
        Dim rowTotCurr As DseLaporanAPSummary.Total_CurrRow

        Try
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Incorrect period", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Me.crv1.Refresh()
                Exit Sub
            End If

            'isikan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_supplier)
            'isikan data beginning balance (invoice + SR)
            Me.IsiSaldo(Me.dtpFrom.Value.Date - TimeSpan.FromDays(1), TBLSaldoAwal)
            'isikan data new invoice
            Me.IsiSales(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, TBLSales)
            'isikan data ending balance
            Me.IsiSaldo(Me.dtpTo.Value.Date, TBLSaldoAkhir)
            'semua data yang ada masukkan ke dalam tabel customer
            Me.IsikanSaldoAwal_DetailOriginal(Me.DS.tblmaster_supplier, TBLSaldoAwal)
            Me.IsikanSaldoAkhir_DetailOriginal(Me.DS.tblmaster_supplier, TBLSaldoAkhir)
            Me.IsikanNewInvoice_DetailOriginal(Me.DS.tblmaster_supplier, TBLSales)


            'Dim i As New Form4
            'i.DataGridView1.DataSource = Nothing
            'DS.tblmaster_supplier.DefaultView.RowFilter = Nothing
            'i.DataGridView1.DataSource = DS.tblmaster_supplier
            'i.ShowDialog()


            'hitung data payment
            TblTotCurr.Rows.Clear()
            For Each row_Supplier In Me.DS.tblmaster_supplier.Rows
                With row_Supplier
                    .Payments = Me.ObjCD.DigitGrouping(CDbl(.Beginning_Balance) + CDbl(.New_Purchase) - _
                                CDbl(.Ending_Balance))

                    'aul tambahan /isikan total currency
                    TblTotCurr.DefaultView.RowFilter = Nothing
                    TblTotCurr.DefaultView.RowFilter = "[primary currency] = " & .Primary_Currency
                    If TblTotCurr.DefaultView.Count > 0 Then
                        rowTotCurr = TblTotCurr.DefaultView(0).Row
                        rowTotCurr.Beginning_Balance += .Beginning_Balance
                        rowTotCurr.Credit += .New_Purchase
                        rowTotCurr.Debit += .Payments
                        rowTotCurr.End_Balance_Original += .Ending_Balance
                        rowTotCurr.End_Balance += .Ending_Balance_IDR
                    Else
                        rowTotCurr = TblTotCurr.NewTotal_CurrRow
                        rowTotCurr.primary_currency = .Primary_Currency
                        rowTotCurr.Currency_Code = .Currency_Code
                        rowTotCurr.Default_Prec = .Default_Prec
                        rowTotCurr.Beginning_Balance = .Beginning_Balance
                        rowTotCurr.Credit = .New_Purchase
                        rowTotCurr.Debit = .Payments
                        rowTotCurr.End_Balance_Original = .Ending_Balance
                        rowTotCurr.End_Balance = .Ending_Balance_IDR
                        TblTotCurr.Rows.Add(rowTotCurr)
                    End If
                    '---
                End With
            Next

            'hilangkan data yang semua datanya nol
            SubTotal = 0
            While SubTotal < Me.DS.tblmaster_supplier.Rows.Count
                row_Supplier = Me.DS.tblmaster_supplier.Rows(SubTotal)
                If row_Supplier.Beginning_Balance = 0 And row_Supplier.New_Purchase = 0 And row_Supplier.Payments = 0 And _
                row_Supplier.Ending_Balance = 0 Then
                    Me.DS.tblmaster_supplier.Rows.Remove(row_Supplier)
                Else
                    SubTotal += 1
                End If
            End While

            'isikan no transaksi
            Me.IsikanNoTransaksi()

            'tampilkan laporan
            Me.report_Detail_NonIDR.SetDataSource(DS)
            Me.report_Detail_NonIDR.SetParameterValue(report_Detail_NonIDR.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.report_Detail_NonIDR.SetParameterValue(report_Detail_NonIDR.Parameter_DefPrice.ParameterFieldName, Me.Def_Prec)
            Me.report_Detail_NonIDR.SetParameterValue(report_Detail_NonIDR.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
            Me.report_Detail_NonIDR.SetParameterValue(report_Detail_NonIDR.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report_Detail_NonIDR.SetParameterValue(report_Detail_NonIDR.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            Me.report_Detail_NonIDR.SetParameterValue(report_Detail_NonIDR.Parameter_TotalCurr_DefPrecIdr.ParameterFieldName, Me.Def_Prec)
            Me.report_Detail_NonIDR.SetParameterValue(report_Detail_NonIDR.Parameter_TotalCurr_Negatif2.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.report_Detail_NonIDR
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBLSaldoAwal.Dispose()
            TBLSales.Dispose()
            TBLSaldoAkhir.Dispose()
            TblTotCurr.Dispose()
        End Try
    End Sub

    Private Sub IsikanNoTransaksi()
        Dim row As DseLaporanAPSummary.tblmaster_supplierRow
        Dim DTA As New DseLaporanAPSummaryTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New DseLaporanAPSummary.tbltrans_aopDataTable
        Dim row1 As DseLaporanAPSummary.tbltrans_aopRow
        Dim PrimaryMin As Double = Double.MaxValue
        Dim TipeTransaksi1 As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
        Dim Filter As String = ""

        Dim DTAPR As New DseLaporanAPSummaryTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBLPR As New DseLaporanAPSummary.tbltrans_purchasereturnDataTable
        Dim rowPR As DseLaporanAPSummary.tbltrans_purchasereturnRow

        Dim DTABalanceInvoice As New DseLaporanAPSummaryTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim TBLBalanceInvoice As New DseLaporanAPSummary.tbltrans_historicalpurchasebalancemanualDataTable
        Dim rowBalanceInvoice As DseLaporanAPSummary.tbltrans_historicalpurchasebalancemanualRow
        Dim DTABalancePR As New DseLaporanAPSummaryTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim TBLBalancePR As New DseLaporanAPSummary.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim rowBalancePR As DseLaporanAPSummary.tbltrans_historicalpurchasereturnbalancemanualRow
        Dim DTADNote As New DseLaporanAPSummaryTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBLDNote As New DseLaporanAPSummary.tbltrans_credit_noteDataTable
        Dim rowDNote As DseLaporanAPSummary.tbltrans_credit_noteRow

        'aul add
        Dim DTAManualDP As New DseLaporanARSummaryTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TBLManualDP As New DseLaporanARSummary.tbltrans_historicalbalancemanual_dpDataTable
        Dim rowManualDP As DseLaporanARSummary.tbltrans_historicalbalancemanual_dpRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAPR.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTABalanceInvoice.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTABalancePR.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADNote.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAManualDP.Connection.ConnectionString = DTA.Connection.ConnectionString 'aul add

            'isikan aop
            PrimaryMin = Double.MaxValue
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis
            For Each row In Me.DS.tblmaster_supplier.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    If PrimaryMin > row.Primary_Transaksi_Asal Then
                        PrimaryMin = row.Primary_Transaksi_Asal
                    End If
                End If
            Next
            If PrimaryMin < Double.MaxValue Then
                DTA.Fill(TBL, PrimaryMin)
                For Each row In Me.DS.tblmaster_supplier.Rows
                    If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                        Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                        TBL.DefaultView.RowFilter = Filter

                        If TBL.DefaultView.Count > 0 Then
                            row1 = TBL.DefaultView(0).Row
                            If Me.ChkUseSupplierInvoice.Checked = True Then
                                row.No_Transaksi_Asal = row1.Supplier_Invoice_No
                            Else
                                row.No_Transaksi_Asal = row1.No_Transaksi
                            End If
                        End If
                    End If
                Next
            End If

            'isikan purchase return 
            PrimaryMin = Double.MaxValue
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis
            For Each row In Me.DS.tblmaster_supplier.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    If PrimaryMin > row.Primary_Transaksi_Asal Then
                        PrimaryMin = row.Primary_Transaksi_Asal
                    End If
                End If
            Next
            If PrimaryMin < Double.MaxValue Then
                DTAPR.Fill(TBLPR, PrimaryMin)
                For Each row In Me.DS.tblmaster_supplier.Rows
                    If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                        Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                        TBLPR.DefaultView.RowFilter = Filter

                        If TBLPR.DefaultView.Count > 0 Then
                            rowPR = TBLPR.DefaultView(0).Row
                            row.No_Transaksi_Asal = rowPR.No_Transaksi
                        End If
                    End If
                Next
            End If

            'isikan historical purchase balance
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
            DTABalanceInvoice.Fill(TBLBalanceInvoice)
            For Each row In Me.DS.tblmaster_supplier.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                    TBLBalanceInvoice.DefaultView.RowFilter = Filter

                    If TBLBalanceInvoice.DefaultView.Count > 0 Then
                        rowBalanceInvoice = TBLBalanceInvoice.DefaultView(0).Row
                        row.No_Transaksi_Asal = rowBalanceInvoice.AOP_Number
                    End If
                End If
            Next

            'aul add
            'isikan historical sales balance
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual
            DTAManualDP.Fill(TBLManualDP, ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP)
            For Each row In Me.DS.tblmaster_supplier.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                    TBLManualDP.DefaultView.RowFilter = Filter

                    If TBLManualDP.DefaultView.Count > 0 Then
                        rowManualDP = TBLManualDP.DefaultView(0).Row
                        If Me.ChkUseSupplierInvoice.Checked = True Then
                            row.No_Transaksi_Asal = rowManualDP.Invoice_No
                        Else
                            row.No_Transaksi_Asal = rowManualDP.DP_No
                        End If
                    End If
                End If
            Next

            'isikan historical purchase return balance
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual
            DTABalancePR.Fill(TBLBalancePR)
            For Each row In Me.DS.tblmaster_supplier.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                    TBLBalancePR.DefaultView.RowFilter = Filter

                    If TBLBalancePR.DefaultView.Count > 0 Then
                        rowBalancePR = TBLBalancePR.DefaultView(0).Row
                        row.No_Transaksi_Asal = rowBalancePR.AOPR_Number
                    End If
                End If
            Next

            'isikan Debet Note & Credit Note
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Dim Tipe As Integer = ClsTransaksiCreditNote.EnumTipeKartu.Supplier

            DTADNote.Fill(TBLDNote)
            For Each row In Me.DS.tblmaster_supplier.Rows
                If row.Tipe_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                    Filter = "Primary = " & row.Primary_Transaksi_Asal.ToString & _
                             " And [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                             " And [tipe kartu] = " & Tipe
                    TBLDNote.DefaultView.RowFilter = Nothing
                    TBLDNote.DefaultView.RowFilter = Filter
                    If TBLDNote.DefaultView.Count > 0 Then
                        rowDNote = TBLDNote.DefaultView(0).Row
                        If Me.ChkUseSupplierInvoice.Checked = True Then
                            row.No_Transaksi_Asal = rowDNote.customer_debet_note
                        Else
                            row.No_Transaksi_Asal = rowDNote.no_transksi
                        End If
                    End If
                ElseIf row.Tipe_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                    Filter = "Primary = " & row.Primary_Transaksi_Asal.ToString & _
                                  " And [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE & _
                                  " And [tipe kartu] = " & Tipe
                    TBLDNote.DefaultView.RowFilter = Nothing
                    TBLDNote.DefaultView.RowFilter = Filter
                    If TBLDNote.DefaultView.Count > 0 Then
                        rowDNote = TBLDNote.DefaultView(0).Row
                        If Me.ChkUseSupplierInvoice.Checked = True Then
                            row.No_Transaksi_Asal = rowDNote.customer_debet_note
                        Else
                            row.No_Transaksi_Asal = rowDNote.no_transksi
                        End If
                    End If
                End If
            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAPR.Dispose()
            TBLPR.Dispose()
            DTADNote.Dispose()
            TBLDNote.Dispose()
        End Try
    End Sub

    Private Sub IsiSaldo(ByVal vTanggal As Date, ByVal vTBLHasil As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim DTASaldo As New DseLaporanAPSummaryTableAdapters.TBLSaldoTableAdapter
        Dim rowSaldo As DseLaporanAPSummary.TBLSaldoRow
        Dim TBLSaldo As New DseLaporanAPSummary.TBLSaldoDataTable
        Dim row As DseLaporanAPSummary.TblPurchaseRow
        Dim i As Double = 0
        Dim Filter As String = ""

        Try
            DTASaldo.Connection.ConnectionString = Core_Module.GetConstring

            'kosongkan tabel saldo
            vTBLHasil.Rows.Clear()

            'isikan data sales
            DTASaldo.Fill(TBLSaldo, Me.TipeKartu, vTanggal)

            'masukkan data sales ke dalam tabel saldo
            For Each rowSaldo In TBLSaldo.Rows
                row = vTBLHasil.NewTblPurchaseRow

                For j As Integer = 0 To vTBLHasil.Columns.Count - 1
                    row.Item(j) = rowSaldo.Item(j)
                Next

                vTBLHasil.Rows.Add(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASaldo.Dispose()
            TBLSaldo.Dispose()
        End Try
    End Sub

    Private Sub IsiSales(ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date, ByVal vTBLSales As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim DTASales As New DseLaporanAPSummaryTableAdapters.TblPurchaseTableAdapter

        Try
            DTASales.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan tabel sales
            vTBLSales.Rows.Clear()

            'isikan data sales
            DTASales.Fill(vTBLSales, vTanggalDari, Me.TipeKartu, Me.TipeTrans1, Me.TipeTrans2, _
            Me.TipeTrans3, Me.TipeTrans4, Me.TipeTrans5, Me.TipeTrans6, Me.TipeTrans7, Me.TipeTrans8, Me.TipeTrans9, vTanggalSampai) 'add aul
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASales.Dispose()
        End Try
    End Sub

    Private Sub IsikanSaldoAwal_DetailIDR(ByVal vTBLCust As DseLaporanAPSummary.tblmaster_supplierDataTable, _
   ByVal vTBLSaldoAwal As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim rowSaldoAwal As DseLaporanAPSummary.TblPurchaseRow
        Dim rowSupp As DseLaporanAPSummary.tblmaster_supplierRow
        Dim DefPrec As Integer = 0

        Try
            vTBLCust.Rows.Clear()

            DefPrec = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            For Each rowSaldoAwal In vTBLSaldoAwal.Rows
                rowSupp = vTBLCust.Newtblmaster_supplierRow

                With rowSaldoAwal
                    rowSupp.Primary_Supplier = .Primary_Supplier
                    rowSupp.Beginning_Balance = .Balance
                    rowSupp.Currency_Code = .Currency_Code
                    rowSupp.Supplier_Code = .Supplier_Code
                    rowSupp.Supplier_Name = .Supplier_Name
                    rowSupp.Default_Prec = DefPrec
                    rowSupp.Ending_Balance = 0
                    rowSupp.Ending_Balance_IDR = 0
                    rowSupp.New_Purchase = 0
                    rowSupp.No_Transaksi_Asal = ""
                    rowSupp.Primary_Currency = .primary_currency
                    rowSupp.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                    rowSupp.Payments = 0
                    rowSupp.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                End With

                vTBLCust.Rows.Add(rowSupp)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanSaldoAwal_DetailOriginal(ByVal vTBLCust As DseLaporanAPSummary.tblmaster_supplierDataTable, _
    ByVal vTBLSaldoAwal As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim rowSaldoAwal As DseLaporanAPSummary.TblPurchaseRow
        Dim rowSupp As DseLaporanAPSummary.tblmaster_supplierRow

        Try
            vTBLCust.Rows.Clear()

            For Each rowSaldoAwal In vTBLSaldoAwal.Rows
                rowSupp = vTBLCust.Newtblmaster_supplierRow

                With rowSaldoAwal
                    rowSupp.Primary_Supplier = .Primary_Supplier
                    rowSupp.Beginning_Balance = .BalanceOriginal
                    rowSupp.Currency_Code = .Currency_Code
                    rowSupp.Supplier_Code = .Supplier_Code
                    rowSupp.Supplier_Name = .Supplier_Name
                    rowSupp.Default_Prec = .Default_Precision
                    rowSupp.Ending_Balance = 0
                    rowSupp.Ending_Balance_IDR = 0
                    rowSupp.New_Purchase = 0
                    rowSupp.No_Transaksi_Asal = ""
                    rowSupp.Primary_Currency = .primary_currency
                    rowSupp.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                    rowSupp.Payments = 0
                    rowSupp.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                End With

                vTBLCust.Rows.Add(rowSupp)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanSaldoAkhir_DetailIDR(ByVal vTBLCust As DseLaporanAPSummary.tblmaster_supplierDataTable, _
    ByVal vTBLSaldoAkhir As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim rowSaldoAkhir As DseLaporanAPSummary.TblPurchaseRow
        Dim rowSupp As DseLaporanAPSummary.tblmaster_supplierRow
        Dim DefPrec As Integer = 0
        Dim Filter As String = ""

        Try
            DefPrec = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            For Each rowSaldoAkhir In vTBLSaldoAkhir.Rows
                Filter = "[Primary Supplier] = " + rowSaldoAkhir.Primary_Supplier.ToString + " AND [Tipe Transaksi Asal] = " + rowSaldoAkhir.Type_Transaksi_Asal.ToString + _
                " AND [Primary Currency] = " + rowSaldoAkhir.primary_currency.ToString + " AND [Primary Transaksi Asal] = " + rowSaldoAkhir.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowSupp = vTBLCust.Newtblmaster_supplierRow

                    With rowSaldoAkhir
                        rowSupp.Primary_Supplier = .Primary_Supplier
                        rowSupp.Beginning_Balance = 0
                        rowSupp.Currency_Code = .Currency_Code
                        rowSupp.Supplier_Code = .Supplier_Code
                        rowSupp.Supplier_Name = .Supplier_Name
                        rowSupp.Default_Prec = DefPrec
                        rowSupp.Ending_Balance = .Balance
                        rowSupp.Ending_Balance_IDR = .Balance
                        rowSupp.New_Purchase = 0
                        rowSupp.No_Transaksi_Asal = ""
                        rowSupp.Primary_Currency = .primary_currency
                        rowSupp.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowSupp.Payments = 0
                        rowSupp.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowSupp)
                Else
                    rowSupp = vTBLCust.DefaultView(0).Row
                    rowSupp.Ending_Balance = rowSaldoAkhir.Balance
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanSaldoAkhir_DetailOriginal(ByVal vTBLCust As DseLaporanAPSummary.tblmaster_supplierDataTable, _
  ByVal vTBLSaldoAkhir As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim rowSaldoAkhir As DseLaporanAPSummary.TblPurchaseRow
        Dim rowSupp As DseLaporanAPSummary.tblmaster_supplierRow
        Dim Filter As String = ""

        Try
            For Each rowSaldoAkhir In vTBLSaldoAkhir.Rows
                Filter = "[Primary Supplier] = " + rowSaldoAkhir.Primary_Supplier.ToString + " AND [Tipe Transaksi Asal] = " + rowSaldoAkhir.Type_Transaksi_Asal.ToString + _
                " AND [Primary Currency] = " + rowSaldoAkhir.primary_currency.ToString + " AND [Primary Transaksi Asal] = " + rowSaldoAkhir.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowSupp = vTBLCust.Newtblmaster_supplierRow

                    With rowSaldoAkhir
                        rowSupp.Primary_Supplier = .Primary_Supplier
                        rowSupp.Beginning_Balance = 0
                        rowSupp.Currency_Code = .Currency_Code
                        rowSupp.Supplier_Code = .Supplier_Code
                        rowSupp.Supplier_Name = .Supplier_Name
                        rowSupp.Default_Prec = .Default_Precision
                        rowSupp.Ending_Balance = .BalanceOriginal
                        rowSupp.Ending_Balance_IDR = .Balance
                        rowSupp.New_Purchase = 0
                        rowSupp.No_Transaksi_Asal = ""
                        rowSupp.Primary_Currency = .primary_currency
                        rowSupp.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowSupp.Payments = 0
                        rowSupp.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowSupp)
                Else
                    rowSupp = vTBLCust.DefaultView(0).Row
                    rowSupp.Ending_Balance = rowSaldoAkhir.BalanceOriginal
                    rowSupp.Ending_Balance_IDR = rowSaldoAkhir.Balance
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanNewInvoice_DetailIDR(ByVal vTBLCust As DseLaporanAPSummary.tblmaster_supplierDataTable, _
    ByVal vTBLNewInvoice As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim rowSales As DseLaporanAPSummary.TblPurchaseRow
        Dim rowSupp As DseLaporanAPSummary.tblmaster_supplierRow
        Dim DefPrec As Integer = 0
        Dim Filter As String = ""

        Try
            DefPrec = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            For Each rowSales In vTBLNewInvoice.Rows
                Filter = "[Primary Supplier] = " + rowSales.Primary_Supplier.ToString + " AND [Tipe Transaksi Asal] = " + rowSales.Type_Transaksi_Asal.ToString + _
                " AND [Primary Currency] = " + rowSales.primary_currency.ToString + " AND [Primary Transaksi Asal] = " + rowSales.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowSupp = vTBLCust.Newtblmaster_supplierRow

                    With rowSales
                        rowSupp.Primary_Supplier = .Primary_Supplier
                        rowSupp.Beginning_Balance = 0
                        rowSupp.Currency_Code = .Currency_Code
                        rowSupp.Supplier_Code = .Supplier_Code
                        rowSupp.Supplier_Name = .Supplier_Name
                        rowSupp.Default_Prec = DefPrec
                        rowSupp.Ending_Balance = 0
                        rowSupp.Ending_Balance_IDR = 0
                        rowSupp.New_Purchase = .Balance
                        rowSupp.No_Transaksi_Asal = ""
                        rowSupp.Primary_Currency = .primary_currency
                        rowSupp.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowSupp.Payments = 0
                        rowSupp.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowSupp)
                Else
                    rowSupp = vTBLCust.DefaultView(0).Row
                    rowSupp.New_Purchase = rowSales.Balance
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanNewInvoice_DetailOriginal(ByVal vTBLCust As DseLaporanAPSummary.tblmaster_supplierDataTable, _
   ByVal vTBLNewInvoice As DseLaporanAPSummary.TblPurchaseDataTable)
        Dim rowSales As DseLaporanAPSummary.TblPurchaseRow
        Dim rowSupp As DseLaporanAPSummary.tblmaster_supplierRow
        Dim Filter As String = ""

        Try
            For Each rowSales In vTBLNewInvoice.Rows
                Filter = "[Primary Supplier] = " + rowSales.Primary_Supplier.ToString + " AND [Tipe Transaksi Asal] = " + rowSales.Type_Transaksi_Asal.ToString + _
                " AND [Primary Currency] = " + rowSales.primary_currency.ToString + " AND [Primary Transaksi Asal] = " + rowSales.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowSupp = vTBLCust.Newtblmaster_supplierRow

                    With rowSales
                        rowSupp.Primary_Supplier = .Primary_Supplier
                        rowSupp.Beginning_Balance = 0
                        rowSupp.Currency_Code = .Currency_Code
                        rowSupp.Supplier_Code = .Supplier_Code
                        rowSupp.Supplier_Name = .Supplier_Name
                        rowSupp.Default_Prec = .Default_Precision
                        rowSupp.Ending_Balance = 0
                        rowSupp.Ending_Balance_IDR = 0
                        rowSupp.New_Purchase = .BalanceOriginal
                        rowSupp.No_Transaksi_Asal = ""
                        rowSupp.Primary_Currency = .primary_currency
                        rowSupp.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowSupp.Payments = 0
                        rowSupp.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowSupp)
                Else
                    rowSupp = vTBLCust.DefaultView(0).Row
                    rowSupp.New_Purchase = rowSales.BalanceOriginal
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul tambahan 
    Private Sub CreateRecap_OriginalCurrency()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row_Supplier As DseLaporanAPSummary.tblmaster_supplierRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Dim TBLSaldoAwal As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TBLSales As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TBLSaldoAkhir As New DseLaporanAPSummary.TblPurchaseDataTable
        Dim TblTotCurr As DseLaporanAPSummary.Total_CurrDataTable = DS.Total_Curr
        Dim rowTotCurr As DseLaporanAPSummary.Total_CurrRow

        Dim tblDummy As New DseLaporanAPSummary.tblmaster_supplierDataTable
        Dim rowDummy As DseLaporanAPSummary.tblmaster_supplierRow

        Try
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Incorrect period", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Me.crv1.Refresh()
                Exit Sub
            End If

            'isikan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_supplier)
            'isikan data beginning balance (invoice + SR)
            Me.IsiSaldo(Me.dtpFrom.Value.Date - TimeSpan.FromDays(1), TBLSaldoAwal)
            'isikan data new invoice
            Me.IsiSales(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, TBLSales)
            'isikan data ending balance
            Me.IsiSaldo(Me.dtpTo.Value.Date, TBLSaldoAkhir)
            'semua data yang ada masukkan ke dalam tabel customer
            Me.IsikanSaldoAwal_DetailOriginal(Me.DS.tblmaster_supplier, TBLSaldoAwal)
            Me.IsikanSaldoAkhir_DetailOriginal(Me.DS.tblmaster_supplier, TBLSaldoAkhir)
            Me.IsikanNewInvoice_DetailOriginal(Me.DS.tblmaster_supplier, TBLSales)

            'Dim F As New Form4
            'Me.DS.tblmaster_supplier.DefaultView.RowFilter = Nothing
            'F.DataGridView1.DataSource = Nothing
            'F.DataGridView1.DataSource = Me.DS.tblmaster_supplier
            'F.ShowDialog()

            For Each row_Supplier In Me.DS.tblmaster_supplier
                rowDummy = tblDummy.Newtblmaster_supplierRow
                For ij As Integer = 0 To tblDummy.Columns.Count - 1
                    rowDummy.Item(ij) = row_Supplier.Item(ij)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next


            DS.tblmaster_supplier.Rows.Clear()
            Dim ListX As New List(Of String)
            For Each rowDummy In tblDummy.Rows
                With rowDummy
                    DS.tblmaster_supplier.DefaultView.RowFilter = Nothing
                    DS.tblmaster_supplier.DefaultView.RowFilter = "[primary currency] = " & .Primary_Currency & _
                                                                  " And [primary Supplier] = " & .Primary_Supplier
                    If DS.tblmaster_supplier.DefaultView.Count > 0 Then
                        Dim dvX As Data.DataRowView
                        For Each dvX In DS.tblmaster_supplier.DefaultView
                            row_Supplier = dvX.Row
                            row_Supplier.Beginning_Balance += .Beginning_Balance
                            row_Supplier.Ending_Balance += .Ending_Balance
                            row_Supplier.Ending_Balance_IDR += .Ending_Balance_IDR
                            row_Supplier.New_Purchase += .New_Purchase
                            row_Supplier.Payments += .Payments
                        Next
                    Else
                        row_Supplier = DS.tblmaster_supplier.Newtblmaster_supplierRow
                        row_Supplier.Primary_Supplier = .Primary_Supplier
                        row_Supplier.Supplier_Code = .Supplier_Code
                        row_Supplier.Supplier_Name = .Supplier_Name
                        row_Supplier.Beginning_Balance = .Beginning_Balance
                        row_Supplier.New_Purchase = .New_Purchase
                        row_Supplier.Payments = .Payments
                        row_Supplier.Ending_Balance = .Ending_Balance
                        row_Supplier.Ending_Balance_IDR = .Ending_Balance_IDR
                        row_Supplier.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        row_Supplier.No_Transaksi_Asal = .No_Transaksi_Asal
                        row_Supplier.Primary_Currency = .Primary_Currency
                        row_Supplier.Currency_Code = .Currency_Code
                        row_Supplier.Default_Prec = .Default_Prec
                        row_Supplier.Tipe_Transaksi_Asal = .Tipe_Transaksi_Asal
                        DS.tblmaster_supplier.Rows.Add(row_Supplier)
                    End If
                End With
            Next


            'hitung data payment
            TblTotCurr.Clear()
            For Each row_Supplier In Me.DS.tblmaster_supplier.Rows
                With row_Supplier
                    .Payments = Me.ObjCD.DigitGrouping(CDbl(.Beginning_Balance) + CDbl(.New_Purchase) - _
                                CDbl(.Ending_Balance))

                    'aul tambahan /isikan total currency
                    TblTotCurr.DefaultView.RowFilter = Nothing
                    TblTotCurr.DefaultView.RowFilter = "[primary currency] = " & .Primary_Currency

                    'Dim dv As Data.DataRowView
                    'For Each dv In TblTotCurr.DefaultView
                    '    rowTotCurr = dv.Row
                    '    wi()
                    'Next

                    If TblTotCurr.DefaultView.Count > 0 Then
                        rowTotCurr = TblTotCurr.DefaultView(0).Row
                        rowTotCurr.Beginning_Balance += .Beginning_Balance
                        rowTotCurr.Credit += .New_Purchase
                        rowTotCurr.Debit += .Payments
                        rowTotCurr.End_Balance_Original += .Ending_Balance
                        rowTotCurr.End_Balance += .Ending_Balance_IDR
                    Else
                        rowTotCurr = TblTotCurr.NewTotal_CurrRow
                        rowTotCurr.primary_currency = .Primary_Currency
                        rowTotCurr.Currency_Code = .Currency_Code
                        rowTotCurr.Default_Prec = .Default_Prec
                        rowTotCurr.Beginning_Balance = .Beginning_Balance
                        rowTotCurr.Credit = .New_Purchase
                        rowTotCurr.Debit = .Payments
                        rowTotCurr.End_Balance_Original = .Ending_Balance
                        rowTotCurr.End_Balance = .Ending_Balance_IDR
                        TblTotCurr.Rows.Add(rowTotCurr)
                    End If
                    '---
                End With
            Next


            'hilangkan data yang semua datanya nol
            SubTotal = 0
            While SubTotal < Me.DS.tblmaster_supplier.Rows.Count
                row_Supplier = Me.DS.tblmaster_supplier.Rows(SubTotal)
                If row_Supplier.Beginning_Balance = 0 And row_Supplier.New_Purchase = 0 And row_Supplier.Payments = 0 And _
                row_Supplier.Ending_Balance = 0 Then
                    Me.DS.tblmaster_supplier.Rows.Remove(row_Supplier)
                Else
                    SubTotal += 1
                End If
            End While

            'isikan no transaksi
            Me.IsikanNoTransaksi()

            'tampilkan laporan
            Me.reportRecapOriginal.SetDataSource(DS)
            Me.reportRecapOriginal.SetParameterValue(reportRecapOriginal.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.reportRecapOriginal.SetParameterValue(reportRecapOriginal.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
            Me.reportRecapOriginal.SetParameterValue(reportRecapOriginal.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.reportRecapOriginal.SetParameterValue(reportRecapOriginal.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            Me.reportRecapOriginal.SetParameterValue(reportRecapOriginal.Parameter_Total_Curr_DefPrecIDR.ParameterFieldName, Me.Def_Prec)
            Me.reportRecapOriginal.SetParameterValue(reportRecapOriginal.Parameter_Total_Curr_Negatif2.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.reportRecapOriginal
            Me.crv1.Refresh()
            Me.crv1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            TBLSaldoAwal.Dispose()
            TBLSales.Dispose()
            TBLSaldoAkhir.Dispose()
            TblTotCurr.Dispose()
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DataView As DataRowView
        Dim row As DseLaporanAPSummary.balance_purchase_viewRow
        Dim row1 As DseLaporanAPSummary.balance_purchase_viewRow
        Dim row_Supplier As DseLaporanAPSummary.tblmaster_supplierRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0
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
                MsgBox("Incorrect period", MsgBoxStyle.Information, "Information")
                Me.crv1.ReportSource = Nothing
                Me.crv1.Refresh()
                Exit Sub
            End If

            If Me.TipeLaporan = EnumFormat.Detail_In_IDR Then
                Me.CreateDetail_InIDR()
                Exit Sub
            ElseIf Me.TipeLaporan = EnumFormat.Detail_In_Original_Currency Then
                Me.CreateDetail_OriginalCurrency()
                Exit Sub
            ElseIf Me.TipeLaporan = EnumFormat.RecapChart Then
                Me.CreateRecapChart()
                Exit Sub
            ElseIf Me.TipeLaporan = EnumFormat.Recap_In_Original_Currency Then
                Me.CreateRecap_OriginalCurrency()
                Exit Sub
            End If

            'isikan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_supplier)

            'start dari data aop + purchase return
            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis
            Me.TipeTrans7 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            Me.TipeTrans8 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Me.TipeTrans9 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual

            'isikan data beginning balance
            Me.DTA1.Fill(DS.balance_purchase_view, Me.TipeKartu, Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
            For Each row_Supplier In Me.DS.tblmaster_supplier.Rows
                SubTotal = 0
                Filter = "[Primary Supplier] =" + row_Supplier.Primary_Supplier.ToString
                Me.DS.balance_purchase_view.DefaultView.RowFilter = Filter
                If Me.DS.balance_purchase_view.DefaultView.Count > 0 Then
                    For Each DataView In Me.DS.balance_purchase_view.DefaultView
                        row = DataView.Row
                        SubTotal += row.Balance
                    Next
                End If
                row_Supplier.Beginning_Balance = Me.ObjCD.DigitGrouping(SubTotal)
            Next

            'isikan data new purchase
            Me.DTA1.FillBy_TanggalSampaiDari(DS.balance_purchase_view, Me.TipeTrans1, Me.TipeTrans2, Me.TipeTrans3, _
            Me.TipeTrans4, Me.TipeTrans5, Me.TipeTrans6, Me.TipeTrans7, Me.TipeTrans8, Me.TipeTrans9, Me.TipeKartu, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
            For Each row_Supplier In Me.DS.tblmaster_supplier.Rows
                SubTotal = 0
                Filter = "[Primary Supplier] =" + row_Supplier.Primary_Supplier.ToString
                Me.DS.balance_purchase_view.DefaultView.RowFilter = Filter
                If Me.DS.balance_purchase_view.DefaultView.Count > 0 Then
                    For Each DataView In Me.DS.balance_purchase_view.DefaultView
                        row1 = DataView.Row
                        SubTotal += row1.Balance
                    Next
                End If
                row_Supplier.New_Purchase = Me.ObjCD.DigitGrouping(SubTotal)
            Next

            'isikan data ending balance
            Me.DTA1.Fill(DS.balance_purchase_view, Me.TipeKartu, Me.dtpTo.Value.Date)
            For Each row_Supplier In Me.DS.tblmaster_supplier.Rows
                SubTotal = 0
                Filter = "[Primary Supplier] =" + row_Supplier.Primary_Supplier.ToString
                Me.DS.balance_purchase_view.DefaultView.RowFilter = Filter
                If Me.DS.balance_purchase_view.DefaultView.Count > 0 Then
                    For Each DataView In Me.DS.balance_purchase_view.DefaultView
                        row = DataView.Row
                        SubTotal += row.Balance
                    Next
                End If
                row_Supplier.Ending_Balance = Me.ObjCD.DigitGrouping(SubTotal)
            Next

            'hitung data payment
            For Each row_Supplier In Me.DS.tblmaster_supplier.Rows
                row_Supplier.Payments = Me.ObjCD.DigitGrouping(CDbl(row_Supplier.Beginning_Balance) + CDbl(row_Supplier.New_Purchase) - _
                CDbl(row_Supplier.Ending_Balance))
            Next

            'hilangkan data yang semua datanya nol
            SubTotal = 0
            While SubTotal < Me.DS.tblmaster_supplier.Rows.Count
                row_Supplier = Me.DS.tblmaster_supplier.Rows(SubTotal)
                If row_Supplier.Beginning_Balance = 0 And row_Supplier.New_Purchase = 0 And row_Supplier.Payments = 0 And _
                row_Supplier.Ending_Balance = 0 Then
                    Me.DS.tblmaster_supplier.Rows.Remove(row_Supplier)
                Else
                    SubTotal += 1
                End If
            End While

            'tampilkan laporan
            Me.report.SetDataSource(DS)
            Me.report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.report.SetParameterValue(report.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
            Me.report.SetParameterValue(report.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.report.SetParameterValue(report.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.report
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Public Sub FromCutOff_APSummary(ByRef vTbl_Supplier As DseLaporanAPSummary.tblmaster_supplierDataTable, _
    ByVal vDtpCutOffPeriod_From As Date, ByVal vDtpCutOffPeriod_To As Date)

        Try
            Me.DS.tblmaster_supplier.Clear()

            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.cboFormat.SelectedIndex = 0
            Me.TipeKartu = CInt(ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP).ToString
            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis
            Me.TipeTrans7 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            Me.TipeTrans8 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Me.TipeTrans9 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual

            Me.dtpFrom.Value = vDtpCutOffPeriod_From.Date
            Me.dtpTo.Value = vDtpCutOffPeriod_To.Date
            Me.CreateDetail_InIDR()

            vTbl_Supplier = Me.DS.tblmaster_supplier.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class