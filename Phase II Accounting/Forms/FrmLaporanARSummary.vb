Option Explicit On
Option Strict Off

Public Class FrmLaporanARSummary
    Public Usercode As String = ""
    Public BackcolorAll As Color = Color.LightSeaGreen
    Dim DS As New DseLaporanARSummary
    Dim DTA1 As New DseLaporanARSummaryTableAdapters.balance_invoice_viewTableAdapter
    Dim DTA3 As New DseLaporanARSummaryTableAdapters.tblmaster_customerTableAdapter
    Dim ObjCD As New Core_Module.ClsCekData
    Dim report As New CrpLaporanARSummary
    Dim reportChart As New CrpLaporanChart_ARSummary
    Dim reportDetailIDR As New CrpLaporanARSummary_Detail2
    Dim reportDetailOriginal As New CrpLaporanARSummary_Detail
    Dim reportRecapOriginal As New CrpLaporanARSummary2
    Dim TipeTrans1, TipeTrans2, TipeTrans3, TipeTrans4, TipeTrans5, TipeTrans6, TipeTrans7, TipeTrans8, TipeTrans9 As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
    Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmLaporanARSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.cboFormat.SelectedIndex = 0

            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis
            Me.TipeTrans7 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            Me.TipeTrans8 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Me.TipeTrans9 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual 'aul add
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiSaldo(ByVal vTanggal As Date, ByVal vTBLHasil As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim DTASaldo As New DseLaporanARSummaryTableAdapters.TBLSaldoTableAdapter
        Dim rowSaldo As DseLaporanARSummary.TBLSaldoRow
        Dim TBLSaldo As New DseLaporanARSummary.TBLSaldoDataTable
        Dim row As DseLaporanARSummary.tbltransaksi_kartu_salesRow
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
                row = vTBLHasil.Newtbltransaksi_kartu_salesRow

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

    Private Sub IsiSales(ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date, ByVal vTBLSales As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim DTASales As New DseLaporanARSummaryTableAdapters.tbltransaksi_kartu_salesTableAdapter

        Try
            DTASales.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan tabel sales
            vTBLSales.Rows.Clear()

            'isikan data sales
            DTASales.FillBy_TglDariSampai(vTBLSales, vTanggalDari, Me.TipeKartu, Me.TipeTrans1, Me.TipeTrans2, _
            Me.TipeTrans3, Me.TipeTrans4, Me.TipeTrans5, Me.TipeTrans6, Me.TipeTrans7, Me.TipeTrans8, Me.TipeTrans9, vTanggalSampai)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASales.Dispose()
        End Try
    End Sub

    Private Sub IsikanSaldoAwal_DetailIDR(ByVal vTBLCust As DseLaporanARSummary.tblmaster_customerDataTable, _
    ByVal vTBLSaldoAwal As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim rowSaldoAwal As DseLaporanARSummary.tbltransaksi_kartu_salesRow
        Dim rowCust As DseLaporanARSummary.tblmaster_customerRow
        Dim DefPrec As Integer = 0

        Try

            vTBLCust.Rows.Clear()

            DefPrec = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            For Each rowSaldoAwal In vTBLSaldoAwal.Rows
                rowCust = vTBLCust.Newtblmaster_customerRow

                With rowSaldoAwal
                    rowCust.Primary = .Primary_Customer
                    rowCust.Beginning_Balance = .Balance
                    rowCust.Currency_Code = .Currency_Code
                    rowCust.Customer_Code = .Customer_Code
                    rowCust.Customer_Name = .Customer_Name
                    rowCust.Default_Prec = DefPrec
                    rowCust.Ending_Balance = 0
                    rowCust.Ending_Balance_IDR = 0
                    rowCust.New_Invoice = 0
                    rowCust.No_Transaksi_Asal = ""
                    rowCust.Primary_Currency = .primary_currency
                    rowCust.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                    rowCust.Receivement = 0
                    rowCust.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                End With

                vTBLCust.Rows.Add(rowCust)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanSaldoAwal_DetailOriginal(ByVal vTBLCust As DseLaporanARSummary.tblmaster_customerDataTable, _
    ByVal vTBLSaldoAwal As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim rowSaldoAwal As DseLaporanARSummary.tbltransaksi_kartu_salesRow
        Dim rowCust As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""

        Try
            vTBLCust.Rows.Clear()

            For Each rowSaldoAwal In vTBLSaldoAwal.Rows
                rowCust = vTBLCust.Newtblmaster_customerRow

                With rowSaldoAwal
                    rowCust.Primary = .Primary_Customer
                    rowCust.Beginning_Balance = .BalanceOriginal
                    rowCust.Currency_Code = .Currency_Code
                    rowCust.Customer_Code = .Customer_Code
                    rowCust.Customer_Name = .Customer_Name
                    rowCust.Default_Prec = .Default_Precision
                    rowCust.Ending_Balance = 0
                    rowCust.Ending_Balance_IDR = 0
                    rowCust.New_Invoice = 0
                    rowCust.No_Transaksi_Asal = ""
                    rowCust.Primary_Currency = .primary_currency
                    rowCust.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                    rowCust.Receivement = 0
                    rowCust.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                End With

                vTBLCust.Rows.Add(rowCust)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanSaldoAkhir_DetailIDR(ByVal vTBLCust As DseLaporanARSummary.tblmaster_customerDataTable, _
    ByVal vTBLSaldoAkhir As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim rowSaldoAkhir As DseLaporanARSummary.tbltransaksi_kartu_salesRow
        Dim rowCust As DseLaporanARSummary.tblmaster_customerRow
        Dim DefPrec As Integer = 0
        Dim Filter As String = ""

        Try
            DefPrec = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            For Each rowSaldoAkhir In vTBLSaldoAkhir.Rows
                Filter = "Primary = " + rowSaldoAkhir.Primary_Customer.ToString + " AND [Tipe Transaksi Asal] = " + rowSaldoAkhir.Type_Transaksi_Asal.ToString + _
                " AND [Primary Currency] = " + rowSaldoAkhir.primary_currency.ToString + " AND [Primary Transaksi Asal] = " + rowSaldoAkhir.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowCust = vTBLCust.Newtblmaster_customerRow

                    With rowSaldoAkhir
                        rowCust.Primary = .Primary_Customer
                        rowCust.Beginning_Balance = 0
                        rowCust.Currency_Code = .Currency_Code
                        rowCust.Customer_Code = .Customer_Code
                        rowCust.Customer_Name = .Customer_Name
                        rowCust.Default_Prec = DefPrec
                        rowCust.Ending_Balance = .Balance
                        rowCust.Ending_Balance_IDR = .Balance
                        rowCust.New_Invoice = 0
                        rowCust.No_Transaksi_Asal = ""
                        rowCust.Primary_Currency = .primary_currency
                        rowCust.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowCust.Receivement = 0
                        rowCust.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowCust)
                Else
                    rowCust = vTBLCust.DefaultView(0).Row
                    rowCust.Ending_Balance = rowSaldoAkhir.Balance
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanSaldoAkhir_DetailOriginal(ByVal vTBLCust As DseLaporanARSummary.tblmaster_customerDataTable, _
  ByVal vTBLSaldoAkhir As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim rowSaldoAkhir As DseLaporanARSummary.tbltransaksi_kartu_salesRow
        Dim rowCust As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""

        Try
            For Each rowSaldoAkhir In vTBLSaldoAkhir.Rows
                Filter = "Primary = " + rowSaldoAkhir.Primary_Customer.ToString + " AND [Tipe Transaksi Asal] = " + rowSaldoAkhir.Type_Transaksi_Asal.ToString + _
                " AND [Primary Currency] = " + rowSaldoAkhir.primary_currency.ToString + " AND [Primary Transaksi Asal] = " + rowSaldoAkhir.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowCust = vTBLCust.Newtblmaster_customerRow

                    With rowSaldoAkhir
                        rowCust.Primary = .Primary_Customer
                        rowCust.Beginning_Balance = 0
                        rowCust.Currency_Code = .Currency_Code
                        rowCust.Customer_Code = .Customer_Code
                        rowCust.Customer_Name = .Customer_Name
                        rowCust.Default_Prec = .Default_Precision
                        rowCust.Ending_Balance = .BalanceOriginal
                        rowCust.Ending_Balance_IDR = .Balance
                        rowCust.New_Invoice = 0
                        rowCust.No_Transaksi_Asal = ""
                        rowCust.Primary_Currency = .primary_currency
                        rowCust.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowCust.Receivement = 0
                        rowCust.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowCust)
                Else
                    rowCust = vTBLCust.DefaultView(0).Row
                    rowCust.Ending_Balance = rowSaldoAkhir.BalanceOriginal
                    rowCust.Ending_Balance_IDR = rowSaldoAkhir.Balance
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanNewInvoice_DetailIDR(ByVal vTBLCust As DseLaporanARSummary.tblmaster_customerDataTable, _
    ByVal vTBLNewInvoice As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim rowSales As DseLaporanARSummary.tbltransaksi_kartu_salesRow
        Dim rowCust As DseLaporanARSummary.tblmaster_customerRow
        Dim DefPrec As Integer = 0
        Dim Filter As String = ""

        Try
            DefPrec = Me.ObjCur.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)

            For Each rowSales In vTBLNewInvoice.Rows
                Filter = "Primary = " + rowSales.Primary_Customer.ToString + " AND [Tipe Transaksi Asal] = " + rowSales.Type_Transaksi_Asal.ToString + _
                " AND [Primary Currency] = " + rowSales.primary_currency.ToString + " AND [Primary Transaksi Asal] = " + rowSales.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowCust = vTBLCust.Newtblmaster_customerRow

                    With rowSales
                        rowCust.Primary = .Primary_Customer
                        rowCust.Beginning_Balance = 0
                        rowCust.Currency_Code = .Currency_Code
                        rowCust.Customer_Code = .Customer_Code
                        rowCust.Customer_Name = .Customer_Name
                        rowCust.Default_Prec = DefPrec
                        rowCust.Ending_Balance = 0
                        rowCust.Ending_Balance_IDR = 0
                        rowCust.New_Invoice = .Balance
                        rowCust.No_Transaksi_Asal = ""
                        rowCust.Primary_Currency = .primary_currency
                        rowCust.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowCust.Receivement = 0
                        rowCust.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowCust)
                Else
                    rowCust = vTBLCust.DefaultView(0).Row
                    rowCust.New_Invoice = rowSales.Balance
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanNewInvoice_DetailOriginal(ByVal vTBLCust As DseLaporanARSummary.tblmaster_customerDataTable, _
   ByVal vTBLNewInvoice As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim rowSales As DseLaporanARSummary.tbltransaksi_kartu_salesRow
        Dim rowCust As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""

        Try
            For Each rowSales In vTBLNewInvoice.Rows

                Filter = "Primary = " & rowSales.Primary_Customer.ToString & _
                " AND [Tipe Transaksi Asal] = " & rowSales.Type_Transaksi_Asal.ToString & _
                " AND [Primary Currency] = " & rowSales.primary_currency.ToString & _
                " AND [Primary Transaksi Asal] = " & rowSales.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowCust = vTBLCust.Newtblmaster_customerRow
                    With rowSales
                        rowCust.Primary = .Primary_Customer
                        rowCust.Beginning_Balance = 0
                        rowCust.Currency_Code = .Currency_Code
                        rowCust.Customer_Code = .Customer_Code
                        rowCust.Customer_Name = .Customer_Name
                        rowCust.Default_Prec = .Default_Precision
                        rowCust.Ending_Balance = 0
                        rowCust.Ending_Balance_IDR = 0
                        rowCust.New_Invoice = .BalanceOriginal
                        rowCust.No_Transaksi_Asal = ""
                        rowCust.Primary_Currency = .primary_currency
                        rowCust.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowCust.Receivement = 0
                        rowCust.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowCust)
                Else
                    rowCust = vTBLCust.DefaultView(0).Row
                    rowCust.New_Invoice = rowSales.BalanceOriginal
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanNewCDNote_DetailOriginal(ByVal vTBLCust As DseLaporanARSummary.tblmaster_customerDataTable, _
 ByVal vTBLCDNote As DseLaporanARSummary.tbltransaksi_kartu_salesDataTable)
        Dim rowCDNote As DseLaporanARSummary.tbltransaksi_kartu_salesRow
        Dim rowCust As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""

        Try
            For Each rowCDNote In vTBLCDNote.Rows
                Filter = "Primary = " & rowCDNote.Primary_Customer.ToString & _
                         " AND [Tipe Transaksi Asal] = " & rowCDNote.Type_Transaksi_Asal.ToString & _
                         " AND [Primary Currency] = " & rowCDNote.primary_currency.ToString & _
                         " AND [Primary Transaksi Asal] = " + rowCDNote.Primary_Transaksi_Asal.ToString

                vTBLCust.DefaultView.RowFilter = Filter
                If vTBLCust.DefaultView.Count = 0 Then
                    rowCust = vTBLCust.Newtblmaster_customerRow

                    With rowCDNote
                        rowCust.Primary = .Primary_Customer
                        rowCust.Beginning_Balance = 0
                        rowCust.Currency_Code = .Currency_Code
                        rowCust.Customer_Code = .Customer_Code
                        rowCust.Customer_Name = .Customer_Name
                        rowCust.Default_Prec = .Default_Precision
                        rowCust.Ending_Balance = 0
                        rowCust.Ending_Balance_IDR = 0
                        rowCust.New_Invoice = .BalanceOriginal
                        rowCust.No_Transaksi_Asal = ""
                        rowCust.Primary_Currency = .primary_currency
                        rowCust.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        rowCust.Receivement = 0
                        rowCust.Tipe_Transaksi_Asal = .Type_Transaksi_Asal
                    End With

                    vTBLCust.Rows.Add(rowCust)
                Else
                    rowCust = vTBLCust.DefaultView(0).Row
                    rowCust.New_Invoice = rowCDNote.BalanceOriginal
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanNoTransaksi()
        Dim row As DseLaporanARSummary.tblmaster_customerRow
        Dim DTA As New DseLaporanARSummaryTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBL As New DseLaporanARSummary.tbltrans_salesinvoiceDataTable
        Dim row1 As DseLaporanARSummary.tbltrans_salesinvoiceRow
        Dim PrimaryMin As Double = Double.MaxValue
        Dim TipeTransaksi1 As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual
        Dim Filter As String = ""
        Dim ObjCDNote As New ClsTransaksiCreditNote

        Dim DTASR As New DseLaporanARSummaryTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBLSR As New DseLaporanARSummary.tbltrans_salesreturnDataTable
        Dim rowSR As DseLaporanARSummary.tbltrans_salesreturnRow

        Dim DTABalanceInvoice As New DseLaporanARSummaryTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim TBLBalanceInvoice As New DseLaporanARSummary.tbltrans_historicalsalesDataTable
        Dim rowBalanceInvoice As DseLaporanARSummary.tbltrans_historicalsalesRow

        Dim DTABalanceSR As New DseLaporanARSummaryTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TBLBalanceSR As New DseLaporanARSummary.tbltrans_historicalsalesreturn_manualDataTable
        Dim rowBalanceSR As DseLaporanARSummary.tbltrans_historicalsalesreturn_manualRow

        Dim DTADNote As New DseLaporanARSummaryTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBLDNote As New DseLaporanARSummary.tbltrans_credit_noteDataTable
        Dim rowDNote As DseLaporanARSummary.tbltrans_credit_noteRow

        'aul add
        Dim DTAManualDP As New DseLaporanARSummaryTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TBLManualDP As New DseLaporanARSummary.tbltrans_historicalbalancemanual_dpDataTable
        Dim rowManualDP As DseLaporanARSummary.tbltrans_historicalbalancemanual_dpRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTABalanceInvoice.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTABalanceSR.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADNote.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAManualDP.Connection.ConnectionString = DTA.Connection.ConnectionString 'aul add
          
            'isikan invoice
            PrimaryMin = Double.MaxValue
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
            For Each row In Me.DS.tblmaster_customer.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    If PrimaryMin > row.Primary_Transaksi_Asal Then
                        PrimaryMin = row.Primary_Transaksi_Asal
                    End If
                End If
            Next
            If PrimaryMin < Double.MaxValue Then
                DTA.Fill(TBL, PrimaryMin)
                For Each row In Me.DS.tblmaster_customer.Rows
                    If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                        Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                        TBL.DefaultView.RowFilter = Filter

                        If TBL.DefaultView.Count > 0 Then
                            row1 = TBL.DefaultView(0).Row
                            row.No_Transaksi_Asal = row1.No_Transaksi
                        End If
                    End If
                Next
            End If

            'isikan sales return 
            PrimaryMin = Double.MaxValue
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis
            For Each row In Me.DS.tblmaster_customer.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    If PrimaryMin > row.Primary_Transaksi_Asal Then
                        PrimaryMin = row.Primary_Transaksi_Asal
                    End If
                End If
            Next
            If PrimaryMin < Double.MaxValue Then
                DTASR.Fill(TBLSR, PrimaryMin)
                For Each row In Me.DS.tblmaster_customer.Rows
                    If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                        Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                        TBLSR.DefaultView.RowFilter = Filter

                        If TBLSR.DefaultView.Count > 0 Then
                            rowSR = TBLSR.DefaultView(0).Row
                            row.No_Transaksi_Asal = rowSR.No_Transaksi
                        End If
                    End If
                Next
            End If

            'isikan historical sales balance
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
            DTABalanceInvoice.Fill(TBLBalanceInvoice)
            For Each row In Me.DS.tblmaster_customer.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                    TBLBalanceInvoice.DefaultView.RowFilter = Filter

                    If TBLBalanceInvoice.DefaultView.Count > 0 Then
                        rowBalanceInvoice = TBLBalanceInvoice.DefaultView(0).Row
                        row.No_Transaksi_Asal = rowBalanceInvoice.Invoice_Number
                    End If
                End If
            Next

            'aul add
            'isikan historical sales balance
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual
            DTAManualDP.Fill(TBLManualDP, ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_Invoice)
            For Each row In Me.DS.tblmaster_customer.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                    TBLManualDP.DefaultView.RowFilter = Filter

                    If TBLManualDP.DefaultView.Count > 0 Then
                        rowManualDP = TBLManualDP.DefaultView(0).Row
                        row.No_Transaksi_Asal = rowManualDP.DP_No
                    End If
                End If
            Next

            'isikan historical sales return balance
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
            DTABalanceSR.Fill(TBLBalanceSR)
            For Each row In Me.DS.tblmaster_customer.Rows
                If row.Tipe_Transaksi_Asal = TipeTransaksi1 Then
                    Filter = "Primary =" + row.Primary_Transaksi_Asal.ToString
                    TBLBalanceSR.DefaultView.RowFilter = Filter

                    If TBLBalanceSR.DefaultView.Count > 0 Then
                        rowBalanceSR = TBLBalanceSR.DefaultView(0).Row
                        row.No_Transaksi_Asal = rowBalanceSR.AOSR_Number
                    End If
                End If
            Next


            'isikan Debet Note & Credit Note
            TipeTransaksi1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Dim Tipe As Integer = ClsTransaksiCreditNote.EnumTipeKartu.Customer

            DTADNote.Fill(TBLDNote)
            For Each row In Me.DS.tblmaster_customer.Rows
                If row.Tipe_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                    Filter = "Primary = " & row.Primary_Transaksi_Asal.ToString & _
                             " And [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                             " And [tipe kartu] = " & Tipe
                    TBLDNote.DefaultView.RowFilter = Nothing
                    TBLDNote.DefaultView.RowFilter = Filter
                    If TBLDNote.DefaultView.Count > 0 Then
                        rowDNote = TBLDNote.DefaultView(0).Row
                        row.No_Transaksi_Asal = rowDNote.no_transksi
                    End If
                ElseIf row.Tipe_Transaksi_Asal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                    Filter = "Primary = " & row.Primary_Transaksi_Asal.ToString & _
                                  " And [tipe transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE & _
                                  " And [tipe kartu] = " & Tipe
                    TBLDNote.DefaultView.RowFilter = Nothing
                    TBLDNote.DefaultView.RowFilter = Filter
                    If TBLDNote.DefaultView.Count > 0 Then
                        rowDNote = TBLDNote.DefaultView(0).Row
                        row.No_Transaksi_Asal = rowDNote.no_transksi
                    End If
                End If
            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Sub

    Private Sub CreateDetail_InIDR()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row_Customer As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Dim TBLSaldoAwal As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLSales As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLSaldoAkhir As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable

        Try
            'kosongkan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_customer)
            'isikan data beginning balance (invoice + SR)
            Me.IsiSaldo(Me.dtpFrom.Value.Date - TimeSpan.FromDays(1), TBLSaldoAwal)
            'isikan data new invoice
            Me.IsiSales(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, TBLSales)
            'isikan data ending balance
            Me.IsiSaldo(Me.dtpTo.Value.Date, TBLSaldoAkhir)
            'semua data yang ada masukkan ke dalam tabel customer
            Me.IsikanSaldoAwal_DetailIDR(Me.DS.tblmaster_customer, TBLSaldoAwal)
            Me.IsikanSaldoAkhir_DetailIDR(Me.DS.tblmaster_customer, TBLSaldoAkhir)
            Me.IsikanNewInvoice_DetailIDR(Me.DS.tblmaster_customer, TBLSales)

            'hitung data payment
            For Each row_Customer In Me.DS.tblmaster_customer.Rows
                row_Customer.Receivement = CDec(row_Customer.Beginning_Balance) + CDec(row_Customer.New_Invoice) - _
                CDec(row_Customer.Ending_Balance)
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

            'isikan no transaksi
            Me.IsikanNoTransaksi()

            'tampilkan laporan
            Me.reportDetailIDR.SetDataSource(DS)
            Me.reportDetailIDR.SetParameterValue(reportDetailIDR.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.reportDetailIDR.SetParameterValue(reportDetailIDR.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
            Me.reportDetailIDR.SetParameterValue(reportDetailIDR.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.reportDetailIDR.SetParameterValue(reportDetailIDR.Parameter_Negatif.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.reportDetailIDR
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
        Dim row_Customer As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Dim TBLSaldoAwal As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLSales As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLCDNote As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLSaldoAkhir As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TblTotCurr As DseLaporanARSummary.Total_CurrDataTable = DS.Total_Curr
        Dim rowTotCurr As DseLaporanARSummary.Total_CurrRow

        Try
            'MsgBox("belum original")
            'kosongkan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_customer)
            'isikan data beginning balance (invoice + SR)
            Me.IsiSaldo(Me.dtpFrom.Value.Date - TimeSpan.FromDays(1), TBLSaldoAwal)
            'isikan data new invoice
            Me.IsiSales(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, TBLSales)
            'isikan data ending balance
            Me.IsiSaldo(Me.dtpTo.Value.Date, TBLSaldoAkhir)
            'semua data yang ada masukkan ke dalam tabel customer
            Me.IsikanSaldoAwal_DetailOriginal(Me.DS.tblmaster_customer, TBLSaldoAwal)
            Me.IsikanSaldoAkhir_DetailOriginal(Me.DS.tblmaster_customer, TBLSaldoAkhir)
            Me.IsikanNewInvoice_DetailOriginal(Me.DS.tblmaster_customer, TBLSales)


            'hitung data payment
            TblTotCurr.Rows.Clear()
            For Each row_Customer In Me.DS.tblmaster_customer.Rows
                With row_Customer
                    .Receivement = CDec(.Beginning_Balance) + CDec(.New_Invoice) - _
                                   CDec(.Ending_Balance)

                    'aul tambahan /isikan total currency
                    TblTotCurr.DefaultView.RowFilter = Nothing
                    TblTotCurr.DefaultView.RowFilter = "[primary currency] = " & .Primary_Currency

                    If TblTotCurr.DefaultView.Count > 0 Then
                        rowTotCurr = TblTotCurr.DefaultView(0).Row
                        rowTotCurr.Beginning_Balance += .Beginning_Balance
                        rowTotCurr.Debit += .New_Invoice
                        rowTotCurr.Credit += .Receivement
                        rowTotCurr.End_Balance_Original += .Ending_Balance
                        rowTotCurr.End_Balance += .Ending_Balance_IDR
                    Else
                        rowTotCurr = TblTotCurr.NewTotal_CurrRow
                        rowTotCurr.primary_currency = .Primary_Currency
                        rowTotCurr.Currency_Code = .Currency_Code
                        rowTotCurr.Default_Prec = .Default_Prec
                        rowTotCurr.Beginning_Balance = .Beginning_Balance
                        rowTotCurr.Debit = .New_Invoice
                        rowTotCurr.Credit = .Receivement
                        rowTotCurr.End_Balance_Original = .Ending_Balance
                        rowTotCurr.End_Balance = .Ending_Balance_IDR
                        TblTotCurr.Rows.Add(rowTotCurr)
                    End If
                    '---
                End With
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

            'isikan no transaksi
            Me.IsikanNoTransaksi()

            'tampilkan laporan
            Me.reportDetailOriginal.SetDataSource(DS)
            Me.reportDetailOriginal.SetParameterValue(reportDetailOriginal.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
            Me.reportDetailOriginal.SetParameterValue(reportDetailOriginal.Parameter_TanggalDari_1.ParameterFieldName, Me.dtpFrom.Value.Date)
            Me.reportDetailOriginal.SetParameterValue(reportDetailOriginal.Parameter_TanggalSampai_1.ParameterFieldName, Me.dtpTo.Value.Date)
            If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Me.Negatif = 1
            Else
                Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.reportDetailOriginal.SetParameterValue(reportDetailOriginal.Parameter_Negatif.ParameterFieldName, Me.Negatif)
            Me.reportDetailOriginal.SetParameterValue(reportDetailOriginal.Parameter_TotalCurr_DefPrecIdr.ParameterFieldName, Me.Def_Prec)
            Me.reportDetailOriginal.SetParameterValue(reportDetailOriginal.Parameter_TotalCurr_Negatif2.ParameterFieldName, Me.Negatif)

            Me.crv1.ReportSource = Me.reportDetailOriginal
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

    'aul tambahan
    Private Sub CreateRecap_OriginalCurrency()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row_Customer As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0

        Dim TBLSaldoAwal As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLSales As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLCDNote As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TBLSaldoAkhir As New DseLaporanARSummary.tbltransaksi_kartu_salesDataTable
        Dim TblTotCurr As DseLaporanARSummary.Total_CurrDataTable = DS.Total_Curr
        Dim rowTotCurr As DseLaporanARSummary.Total_CurrRow
        Dim tblDummy As New DseLaporanARSummary.tblmaster_customerDataTable
        Dim rowDummy As DseLaporanARSummary.tblmaster_customerRow

        Try
            'MsgBox("belum original")
            'kosongkan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_customer)
            'isikan data beginning balance (invoice + SR)
            Me.IsiSaldo(Me.dtpFrom.Value.Date - TimeSpan.FromDays(1), TBLSaldoAwal)
            'isikan data new invoice
            Me.IsiSales(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, TBLSales)
            'isikan data ending balance
            Me.IsiSaldo(Me.dtpTo.Value.Date, TBLSaldoAkhir)
            'semua data yang ada masukkan ke dalam tabel customer
            Me.IsikanSaldoAwal_DetailOriginal(Me.DS.tblmaster_customer, TBLSaldoAwal)
            Me.IsikanSaldoAkhir_DetailOriginal(Me.DS.tblmaster_customer, TBLSaldoAkhir)
            Me.IsikanNewInvoice_DetailOriginal(Me.DS.tblmaster_customer, TBLSales)


            For Each row_Customer In Me.DS.tblmaster_customer
                rowDummy = tblDummy.Newtblmaster_customerRow
                For ij As Integer = 0 To tblDummy.Columns.Count - 1
                    rowDummy.Item(ij) = row_Customer.Item(ij)
                Next
                tblDummy.Rows.Add(rowDummy)
            Next


            DS.tblmaster_customer.Rows.Clear()
            Dim ListX As New List(Of String)
            For Each rowDummy In tblDummy.Rows
                With rowDummy
                    DS.tblmaster_customer.DefaultView.RowFilter = Nothing
                    DS.tblmaster_customer.DefaultView.RowFilter = "[primary currency] = " & .Primary_Currency & _
                                                                  " And [primary] = " & .Primary
                    If DS.tblmaster_customer.DefaultView.Count > 0 Then
                        Dim dvX As Data.DataRowView
                        For Each dvX In DS.tblmaster_customer.DefaultView
                            row_Customer = dvX.Row
                            row_Customer.Beginning_Balance += .Beginning_Balance
                            row_Customer.Ending_Balance += .Ending_Balance
                            row_Customer.Ending_Balance_IDR += .Ending_Balance_IDR
                            row_Customer.New_Invoice += .New_Invoice
                            row_Customer.Receivement += .Receivement
                        Next
                    Else
                        row_Customer = DS.tblmaster_customer.Newtblmaster_customerRow
                        row_Customer.Primary = .Primary
                        row_Customer.Customer_Code = .Customer_Code
                        row_Customer.Customer_Name = .Customer_Name
                        row_Customer.Beginning_Balance = .Beginning_Balance
                        row_Customer.New_Invoice = .New_Invoice
                        row_Customer.Receivement = .Receivement
                        row_Customer.Ending_Balance = .Ending_Balance
                        row_Customer.Ending_Balance_IDR = .Ending_Balance_IDR
                        row_Customer.Primary_Transaksi_Asal = .Primary_Transaksi_Asal
                        row_Customer.No_Transaksi_Asal = .No_Transaksi_Asal
                        row_Customer.Primary_Currency = .Primary_Currency
                        row_Customer.Currency_Code = .Currency_Code
                        row_Customer.Default_Prec = .Default_Prec
                        row_Customer.Tipe_Transaksi_Asal = .Tipe_Transaksi_Asal
                        DS.tblmaster_customer.Rows.Add(row_Customer)
                    End If
                End With
            Next


            'hitung data payment
            TblTotCurr.Clear()
            For Each row_Customer In Me.DS.tblmaster_customer.Rows
                With row_Customer
                    .Receivement = CDec(.Beginning_Balance) + CDec(.New_Invoice) - _
                                   CDec(.Ending_Balance)

                    'aul tambahan /isikan total currency
                    TblTotCurr.DefaultView.RowFilter = Nothing
                    TblTotCurr.DefaultView.RowFilter = "[primary currency] = " & .Primary_Currency

                    If TblTotCurr.DefaultView.Count > 0 Then
                        rowTotCurr = TblTotCurr.DefaultView(0).Row
                        rowTotCurr.Beginning_Balance += .Beginning_Balance
                        rowTotCurr.Debit += .New_Invoice
                        rowTotCurr.Credit += .Receivement
                        rowTotCurr.End_Balance_Original += .Ending_Balance
                        rowTotCurr.End_Balance += .Ending_Balance_IDR
                    Else
                        rowTotCurr = TblTotCurr.NewTotal_CurrRow
                        rowTotCurr.primary_currency = .Primary_Currency
                        rowTotCurr.Currency_Code = .Currency_Code
                        rowTotCurr.Default_Prec = .Default_Prec
                        rowTotCurr.Beginning_Balance = .Beginning_Balance
                        rowTotCurr.Debit = .New_Invoice
                        rowTotCurr.Credit = .Receivement
                        rowTotCurr.End_Balance_Original = .Ending_Balance
                        rowTotCurr.End_Balance = .Ending_Balance_IDR
                        TblTotCurr.Rows.Add(rowTotCurr)
                    End If
                    '---
                End With
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

            'isikan no transaksi
            Me.IsikanNoTransaksi()

            'aul tambahan /isikan total keseluruhan end balance IDR

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
        Dim row As DseLaporanARSummary.balance_invoice_viewRow
        Dim row1 As DseLaporanARSummary.balance_invoice_viewRow
        Dim row_Customer As DseLaporanARSummary.tblmaster_customerRow
        Dim Filter As String = ""
        Dim SubTotal As Decimal = 0
        Dim SubTotalIDR As Decimal = 0
        Dim EndLicenseDate As Date = #1/1/1950#

        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
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
                Me.CreateMonthlyChart()
                Exit Sub
            ElseIf Me.TipeLaporan = EnumFormat.Recap_In_Original_Currency Then
                Me.CreateRecap_OriginalCurrency()
                Exit Sub
            End If

            'isikan data master supplier
            Me.DTA3.Fill(Me.DS.tblmaster_customer)

            'start dari data invoice
            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis
            Me.TipeTrans7 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            Me.TipeTrans8 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Me.TipeTrans9 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual 'aul add

            'isikan data beginning balance (invoice + SR)
            Me.DTA1.Fill(DS.balance_invoice_view, _
            Me.TipeKartu, Me.dtpFrom.Value.Date - TimeSpan.FromDays(1))
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
            Me.TipeTrans3, Me.TipeTrans4, Me.TipeTrans5, Me.TipeTrans6, Me.TipeTrans7, Me.TipeTrans8, Me.TipeTrans9, Me.TipeKartu, Me.dtpTo.Value.Date, Me.dtpFrom.Value.Date)
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
            Me.DTA1.Fill(DS.balance_invoice_view, Me.TipeKartu, Me.dtpTo.Value.Date)
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

    Public Sub FromCutOff_ARSummary(ByRef vTbl_Customer As DseLaporanARSummary.tblmaster_customerDataTable, _
   ByVal vDtpCutOffPeriod_From As Date, ByVal vDtpCutOffPeriod_To As Date)

        Try
            Me.DS.tblmaster_customer.Rows.Clear()

            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.cboFormat.SelectedIndex = 0
            Me.TipeKartu = CInt(ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR).ToString
            Me.TipeTrans1 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual
            Me.TipeTrans2 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer
            Me.TipeTrans3 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis
            Me.TipeTrans4 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual
            Me.TipeTrans5 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual_Transfer
            Me.TipeTrans6 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis
            Me.TipeTrans7 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
            Me.TipeTrans8 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
            Me.TipeTrans9 = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual

            Me.dtpFrom.Value = vDtpCutOffPeriod_From.Date
            Me.dtpTo.Value = vDtpCutOffPeriod_To.Date
            Me.CreateDetail_InIDR()

            vTbl_Customer = Me.DS.tblmaster_customer.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class