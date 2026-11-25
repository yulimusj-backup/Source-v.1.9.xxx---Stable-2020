Option Explicit On
Option Strict Off

Partial Public Class FrmLaporanAPSummary
    Public Sub CreateRecapChart()
        Dim StartDate, EndDate As Date
        Dim TBL As DseLaporanAPSummary.AP_ChartDataTable = Me.DS.AP_Chart
        Dim row As DseLaporanAPSummary.AP_ChartRow
        Dim TBLAsal As DseLaporanAPSummary.tblmaster_supplierDataTable = Me.DS.tblmaster_supplier
        Dim rowAsal As DseLaporanAPSummary.tblmaster_supplierRow
        Dim BeginingBalance, NewInvoice, Payment, EndingBalance As Decimal
        Dim DummyEndDate As Date

        Try
            TBL.Rows.Clear()

            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value.Date)

            While StartDate < EndDate
                'buatkan bulanan
                DummyEndDate = Me.ObjCD.GetLasDateOfTheMonth(StartDate)
                Me.CreateRecapChart(StartDate, DummyEndDate)

                'get data untuk bulan itu
                BeginingBalance = 0
                NewInvoice = 0
                Payment = 0
                EndingBalance = 0
                For Each rowAsal In TBLAsal.Rows
                    BeginingBalance += rowAsal.Beginning_Balance
                    NewInvoice += rowAsal.New_Purchase
                    Payment += rowAsal.Payments 
                    EndingBalance += rowAsal.Ending_Balance
                Next

                'pindahkan data ke dalam cross tab
                '1. Begining Balance
                'row = TBL.NewAP_ChartRow
                'row._Date = StartDate
                'row.Tipe_Data = "Begining Balance"
                'row.Value = BeginingBalance
                'TBL.Rows.Add(row)
                '2. New Invoice                 
                row = TBL.NewAP_ChartRow
                row._Date = DummyEndDate
                row.Tipe_Data = "New Invoice"
                row.Value = NewInvoice
                TBL.Rows.Add(row)
                '3. Payment
                row = TBL.NewAP_ChartRow
                row._Date = DummyEndDate
                row.Tipe_Data = "Payment"
                row.Value = Payment
                TBL.Rows.Add(row)
                '4. Ending Balance
                row = TBL.NewAP_ChartRow
                row._Date = DummyEndDate
                row.Tipe_Data = "Ending Balance"
                row.Value = EndingBalance
                TBL.Rows.Add(row)

                'majukan bulannya
                StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))
            End While

            'tampilkan laporan
            Me.reportChart.SetDataSource(DS)
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.reportChart.SetParameterValue(reportChart.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.reportChart.SetParameterValue(reportChart.Parameter_TanggalDari.ParameterFieldName, Me.dtpFrom.Value.Date)
            Me.reportChart.SetParameterValue(reportChart.Parameter_TanggalSampai.ParameterFieldName, Me.dtpTo.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.reportChart.SetParameterValue(reportChart.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.reportChart
            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CreateRecapChart(ByVal vStartDate As Date, ByVal vEndDate As Date)
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DataView As DataRowView
        Dim row As DseLaporanAPSummary.balance_purchase_viewRow
        Dim row1 As DseLaporanAPSummary.balance_purchase_viewRow
        Dim row_Supplier As DseLaporanAPSummary.tblmaster_supplierRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Try
            'isikan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_supplier)

            'start dari data aop + purchase return
            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis

            'isikan data beginning balance
            Me.DTA1.Fill(DS.balance_purchase_view, Me.TipeKartu, vStartDate - TimeSpan.FromDays(1))
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
            Me.TipeTrans4, Me.TipeTrans5, Me.TipeTrans6, Me.TipeTrans7, Me.TipeTrans8, Me.TipeTrans9, Me.TipeKartu, vEndDate, vStartDate)
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
            Me.DTA1.Fill(DS.balance_purchase_view, Me.TipeKartu, vEndDate)
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub
End Class
