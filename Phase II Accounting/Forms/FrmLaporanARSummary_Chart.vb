Option Explicit On
Option Strict Off

Partial Public Class FrmLaporanARSummary
    Public Sub CreateMonthlyChart()
        Dim StartDate, EndDate As Date
        Dim TBL As DseLaporanARSummary.AR_ChartDataTable = Me.DS.AR_Chart
        Dim row As DseLaporanARSummary.AR_ChartRow
        Dim TBLAsal As DseLaporanARSummary.tblmaster_customerDataTable = Me.DS.tblmaster_customer
        Dim rowAsal As DseLaporanARSummary.tblmaster_customerRow
        Dim BeginingBalance, NewInvoice, Payment, EndingBalance As Decimal
        Dim DummyEndDate As Date

        Try
            TBL.Rows.Clear()

            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpTo.Value.Date)

            While StartDate < EndDate
                'buatkan bulanan
                DummyEndDate = Me.ObjCD.GetLasDateOfTheMonth(StartDate)
                Me.CreateMonthlyChart(StartDate, DummyEndDate)

                'get data untuk bulan itu
                BeginingBalance = 0
                NewInvoice = 0
                Payment = 0
                EndingBalance = 0
                For Each rowAsal In TBLAsal.Rows
                    BeginingBalance += rowAsal.Beginning_Balance
                    NewInvoice += rowAsal.New_Invoice
                    Payment += rowAsal.Receivement
                    EndingBalance += rowAsal.Ending_Balance
                Next

                'pindahkan data ke dalam cross tab
                '1. Begining Balance
                'row = TBL.NewAR_ChartRow
                'row._Date = StartDate
                'row.Tipe_Data = "Begining Balance"
                'row.Value = BeginingBalance
                'TBL.Rows.Add(row)
                '2. New Invoice                 
                row = TBL.NewAR_ChartRow
                row._Date = DummyEndDate
                row.Tipe_Data = "New Invoice"
                row.Value = NewInvoice
                TBL.Rows.Add(row)
                '3. Payment
                row = TBL.NewAR_ChartRow
                row._Date = DummyEndDate
                row.Tipe_Data = "Payment"
                row.Value = Payment
                TBL.Rows.Add(row)
                '4. Ending Balance
                row = TBL.NewAR_ChartRow
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

    Public Sub CreateMonthlyChart(ByVal vStartDate As Date, ByVal vEndDate As Date)
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim DataView As DataRowView
        Dim row As DseLaporanARSummary.balance_invoice_viewRow
        Dim row1 As DseLaporanARSummary.balance_invoice_viewRow
        Dim row_Customer As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Try
            Me.DS.tblmaster_customer.Rows.Clear()
            Me.DS.balance_invoice_view.Rows.Clear()

            'isikan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_customer)

            'start dari data invoice
            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis

            'isikan data beginning balance (invoice + SR)
            Me.DTA1.Fill(DS.balance_invoice_view, _
            Me.TipeKartu, vStartDate.Date - TimeSpan.FromDays(1))
            For Each row_Customer In Me.DS.tblmaster_customer.Rows
                SubTotal = 0
                Filter = "[Primary Customer] =" + row_Customer.Primary.ToString
                Me.DS.balance_invoice_view.DefaultView.RowFilter = Filter
                If Me.DS.balance_invoice_view.DefaultView.Count > 0 Then
                    For Each DataView In Me.DS.balance_invoice_view.DefaultView
                        row = DataView.Row
                        SubTotal += row.Balance
                    Next
                End If
                row_Customer.Beginning_Balance = Me.ObjCD.DigitGrouping(SubTotal)
            Next

            'isikan data new invoice
            Me.DTA1.FillBy_TanggalSampaiDari(DS.balance_invoice_view, Me.TipeTrans1, Me.TipeTrans2, _
            Me.TipeTrans3, Me.TipeTrans4, Me.TipeTrans5, Me.TipeTrans6, Me.TipeTrans7, Me.TipeTrans8, Me.TipeTrans9, Me.TipeKartu, vEndDate.Date, vStartDate.Date)
            For Each row_Customer In Me.DS.tblmaster_customer.Rows
                SubTotal = 0
                Filter = "[Primary Customer] =" + row_Customer.Primary.ToString
                Me.DS.balance_invoice_view.DefaultView.RowFilter = Filter
                If Me.DS.balance_invoice_view.DefaultView.Count > 0 Then
                    For Each DataView In Me.DS.balance_invoice_view.DefaultView
                        row1 = DataView.Row
                        SubTotal += row1.Balance
                    Next
                End If
                row_Customer.New_Invoice = Me.ObjCD.DigitGrouping(SubTotal)
            Next

            'isikan data ending balance
            Me.DTA1.Fill(DS.balance_invoice_view, Me.TipeKartu, vEndDate.Date)
            For Each row_Customer In Me.DS.tblmaster_customer.Rows
                SubTotal = 0
                Filter = "[Primary Customer] =" + row_Customer.Primary.ToString
                Me.DS.balance_invoice_view.DefaultView.RowFilter = Filter
                If Me.DS.balance_invoice_view.DefaultView.Count > 0 Then
                    For Each DataView In Me.DS.balance_invoice_view.DefaultView
                        row = DataView.Row
                        SubTotal += row.Balance
                    Next
                End If
                row_Customer.Ending_Balance = Me.ObjCD.DigitGrouping(SubTotal)
            Next

            'hitung data payment
            For Each row_Customer In Me.DS.tblmaster_customer.Rows
                row_Customer.Receivement = Me.ObjCD.DigitGrouping(CDbl(row_Customer.Beginning_Balance) + CDbl(row_Customer.New_Invoice) - _
                CDbl(row_Customer.Ending_Balance))
            Next

            'hilangkan data yang semua datanya nol
            SubTotal = 0
            While SubTotal < Me.DS.tblmaster_customer.Rows.Count
                row_Customer = Me.DS.tblmaster_customer.Rows(SubTotal)
                If row_Customer.Beginning_Balance = 0 And row_Customer.New_Invoice = 0 And row_Customer.Receivement = 0 And _
                row_Customer.Ending_Balance = 0 Then
                    Me.DS.tblmaster_customer.Rows.Remove(row_Customer)
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
