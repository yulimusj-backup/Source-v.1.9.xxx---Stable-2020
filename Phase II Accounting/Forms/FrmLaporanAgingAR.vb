Option Strict Off
Option Explicit On

Public Class FrmLaporanAgingAR
    Dim DS As New DseLaporanAging
    Dim DTA As New DseLaporanAgingTableAdapters.balance_invoice_viewTableAdapter
    Dim DTA2 As New DseLaporanAgingTableAdapters.DataTable1TableAdapter
    Dim TBL As DseLaporanAging.balance_invoice_viewDataTable = DS.balance_invoice_view
    Dim TBL2 As DseLaporanAging.DataTable1DataTable = DS.DataTable1
    Dim row As DseLaporanAging.balance_invoice_viewRow
    Dim row2 As DseLaporanAging.DataTable1Row
    Dim TipeBalanceSR As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    Dim Periode5, Periode4, Periode3, Periode2, Periode1 As Date
    Dim ObjCD As New Core_Module.ClsCekData
    Dim report As New CrpLaporanAging
    Dim report1 As New CrpLaporanAging_Landscape
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim ObjCust As New Core_Module.clsMasterCustomer
    Dim TblCust As New Core_Module.dseCustomer.tblmaster_customer1DataTable
    Dim DV As DataRowView
    Dim rowCust As Core_Module.dseCustomer.tblmaster_customer1Row

    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim FontNumberSize As Integer = 9

    Private Function GetSRNumber(ByVal vPrimarySR As Double, ByVal vTBLSR As DseLaporanAging.tbltrans_salesreturnDataTable, _
   ByVal vTBLSR_Manual As DseLaporanAging.tbltrans_historicalsalesreturn_manualDataTable, ByVal vTipeTransaksi As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal) As String
        Dim rowSR As DseLaporanAging.tbltrans_salesreturnRow
        Dim rowSR_Manual As DseLaporanAging.tbltrans_historicalsalesreturn_manualRow
        Dim filter As String = ""

        Try
            If vTipeTransaksi = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Then
                'cari dalam tbl sr manual
                filter = "Primary = " + vPrimarySR.ToString
                vTBLSR_Manual.DefaultView.RowFilter = filter
                If vTBLSR_Manual.DefaultView.Count > 0 Then
                    rowSR_Manual = vTBLSR_Manual.DefaultView(0).Row
                    Return rowSR_Manual.AOSR_Number
                End If
            ElseIf vTipeTransaksi = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual_Transfer Or vTipeTransaksi = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
                'cari dalam tbl sr
                filter = "Primary = " + vPrimarySR.ToString
                vTBLSR.DefaultView.RowFilter = filter
                If vTBLSR.DefaultView.Count > 0 Then
                    rowSR = vTBLSR.DefaultView(0).Row
                    Return rowSR.No_Transaksi
                End If
            End If


            'kalo sampe sini kembalikan default value
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInvoiceNumber(ByVal vPrimaryInvoice As Double, ByVal vTBLInvoice As DseLaporanAging.tbltrans_salesinvoiceDataTable, _
    ByVal vTBL_InvManual As DseLaporanAging.tbltrans_historicalsalesDataTable, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal) As String
        Dim rowInv As DseLaporanAging.tbltrans_salesinvoiceRow
        Dim rowInv_Manual As DseLaporanAging.tbltrans_historicalsalesRow
        Dim Filter As String = ""

        Try
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer Or _
            vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
                'dari invoice biasa
                Filter = "[Primary Invoice] = " + vPrimaryInvoice.ToString
                vTBLInvoice.DefaultView.RowFilter = Filter
                If vTBLInvoice.DefaultView.Count > 0 Then
                    rowInv = vTBLInvoice.DefaultView(0).Row
                    Return rowInv.Invoice_Number
                End If
            Else
                'dari invoice manual
                Filter = "Primary = " + vPrimaryInvoice.ToString
                vTBL_InvManual.DefaultView.RowFilter = Filter
                If vTBL_InvManual.DefaultView.Count > 0 Then
                    rowInv_Manual = vTBL_InvManual.DefaultView(0).Row
                    Return rowInv_Manual.Invoice_Number
                End If
            End If

            'kalo mpe sini maka kemablikan nilai default
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetTanggalMinimum(ByVal vTBL As DseLaporanAging.tbltransaksi_kartu_ar_apDataTable) As Date
        Dim rowAR As DseLaporanAging.tbltransaksi_kartu_ar_apRow
        Dim D As Date = Date.MaxValue

        Try
            For Each rowAR In vTBL.Rows
                If D.Date < rowAR._Date.Date Then
                    D = rowAR._Date
                End If
            Next

            If D <> Date.MaxValue Then
                Return D
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsiTBL()
        Dim TBLAR As New DseLaporanAging.tbltransaksi_kartu_ar_apDataTable
        Dim DTAAR As New DseLaporanAgingTableAdapters.tbltransaksi_kartu_ar_apTableAdapter
        Dim rowAR As DseLaporanAging.tbltransaksi_kartu_ar_apRow
        Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR
        Dim TipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual

        Dim NoInvoice As String = ""

        Dim TBLInvoice As New DseLaporanAging.tbltrans_salesinvoiceDataTable
        Dim DTAInvoice As New DseLaporanAgingTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBLInvoice_Manual As New DseLaporanAging.tbltrans_historicalsalesDataTable
        Dim DTAInvoice_Manual As New DseLaporanAgingTableAdapters.tbltrans_historicalsalesTableAdapter
        Dim DTASR As New DseLaporanAgingTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBLSR As New DseLaporanAging.tbltrans_salesreturnDataTable
        Dim DTASR_Manual As New DseLaporanAgingTableAdapters.tbltrans_historicalsalesreturn_manualTableAdapter
        Dim TBLSR_Manual As New DseLaporanAging.tbltrans_historicalsalesreturn_manualDataTable
        Dim PrimaryInvoiceAcuan As Double = -1
        Dim Filter As String = ""
        Dim i As Double = 0
        Dim MinDate, MaxDate As Date
        Dim DTACDNote As New DseLaporanAgingTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TblCDNote As New DseLaporanAging.tbltrans_credit_noteDataTable
        Dim rowCDNote As DseLaporanAging.tbltrans_credit_noteRow

        'aul add
        Dim DtaManualDP As New DseLaporanAgingTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TblManualDP As New DseLaporanAging.tbltrans_historicalbalancemanual_dpDataTable

        Try
            'kumpulkan kartu ar sampai dengan tanggal yang bersangkutan
            If Me.clbCustomer.Items.Count = Me.clbCustomer.CheckedItems.Count Then
                DTAAR.Fill(TBLAR, TipeKartu, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriode.Value))
            Else
                DTAAR.ClearBeforeFill = False
                For Each DV In Me.clbCustomer.CheckedItems
                    rowCust = DV.Row
                    DTAAR.FillBy_PrimaryCustomer(TBLAR, TipeKartu, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriode.Value), rowCust.Primary)
                Next
            End If

            'isikan data yang akan digunakan untuk perhitung
            MinDate = Me.GetTanggalMinimum(TBLAR)
            MaxDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriode.Value)
            DTAInvoice.Fill(TBLInvoice, MaxDate, MinDate)
            DTAInvoice_Manual.Fill(TBLInvoice_Manual, MaxDate, MinDate)
            DTAInvoice_Manual.Fill(TBLInvoice_Manual, MaxDate, MinDate)
            DTASR.Fill(TBLSR, MaxDate, MinDate)
            DTASR_Manual.Fill(TBLSR_Manual, MaxDate, MinDate)
            DTACDNote.FillBy(TblCDNote, MinDate, MaxDate)
            DtaManualDP.Fill(TblManualDP, ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_Invoice, MinDate, MaxDate)
            'isikan no document asal

            'pindahkan ke tabel cetak
            Me.TBL.Rows.Clear()
            For Each rowAR In TBLAR.Rows
                Me.row = Me.TBL.Newbalance_invoice_viewRow
                'yang dimasukkan ke dalam balance, hanya invoice2/sr2 yang belum selesai pembayarannya
                With Me.row
                    .Primary = Me.GetPrimaryMax + 1
                    .Balance = rowAR.Saldo
                    .Customer = rowAR.Customer
                    .Due_Date = rowAR._Date + TimeSpan.FromDays(rowAR.Payment_Terms_Transaksi_Acuan)
                    .Invoice_Date = rowAR._Date
                    TipeTransaksiAsal = rowAR.Type_Transaksi_Asal
                    If TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual_Transfer Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis Then
                        .Invoice_Number = Me.GetInvoiceNumber(rowAR.Primary_Transaksi_Asal, TBLInvoice, TBLInvoice_Manual, TipeTransaksiAsal)
                    ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Manual_Transfer Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis Then
                        .Invoice_Number = Me.GetSRNumber(rowAR.Primary_Transaksi_Asal, TBLSR, TBLSR_Manual, TipeTransaksiAsal)
                    ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                        .Invoice_Number = Me.GetCDNoteNumber(rowAR.Primary_Transaksi_Asal, TblCDNote, TipeTransaksiAsal)
                    ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_DP_Manual Then 'aul add
                        .Invoice_Number = Me.GetManualDPNumber(rowAR.Primary_Transaksi_Asal, TblManualDP)
                    End If
                    .Payment_Terms = rowAR.Payment_Terms_Transaksi_Acuan
                    .Periode1 = 0
                    .Periode2 = 0
                    .Periode3 = 0
                    .Periode4 = 0
                    .Periode5 = 0
                    .Periode6 = 0
                    .Primary_Customer = rowAR.Primary_Customer
                    .Primary_Invoice = rowAR.Primary_Transaksi_Asal
                    .Remarks = ""
                    .Tipe = 0
                End With
                Me.TBL.Rows.Add(Me.row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAR.Dispose()
            TBLAR.Dispose()

            TBLInvoice.Dispose()
            DTAInvoice.Dispose()
            TBLInvoice_Manual.Dispose()
            DTAInvoice_Manual.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
            DTASR_Manual.Dispose()
            TBLSR_Manual.Dispose()
        End Try
    End Sub


    Private Function GetCDNoteNumber(ByVal vPrimaryCDNote As Double, ByVal vTBLCDNote As DseLaporanAging.tbltrans_credit_noteDataTable, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal) As String
        Dim rowCDNote As DseLaporanAging.tbltrans_credit_noteRow
        Dim Filter As String = ""

        Try
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                'dari Credit Note
                Filter = ""
                Filter = "[Primary] = " + vPrimaryCDNote.ToString & _
                         " And [Tipe Transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE & _
                         " And [TIpe Kartu] = " & ClsTransaksiCreditNote.EnumTipeKartu.Customer
                vTBLCDNote.DefaultView.RowFilter = Nothing
                vTBLCDNote.DefaultView.RowFilter = Filter
                If vTBLCDNote.DefaultView.Count > 0 Then
                    rowCDNote = vTBLCDNote.DefaultView(0).Row
                    Return rowCDNote.no_transksi
                End If

            Else
                'dari Credit Note
                Filter = ""
                Filter = "[Primary] = " + vPrimaryCDNote.ToString & _
                         " And [Tipe Transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                         " And [TIpe Kartu] = " & ClsTransaksiCreditNote.EnumTipeKartu.Customer
                vTBLCDNote.DefaultView.RowFilter = Nothing
                vTBLCDNote.DefaultView.RowFilter = Filter
                If vTBLCDNote.DefaultView.Count > 0 Then
                    rowCDNote = vTBLCDNote.DefaultView(0).Row
                    Return rowCDNote.no_transksi
                End If
            End If

            'kalo mpe sini maka kemablikan nilai default
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetPrimaryMax() As Double
        Dim i As DseLaporanAging.balance_invoice_viewRow

        Try
            If Me.TBL.Rows.Count > 0 Then
                i = Me.TBL.Rows(Me.TBL.Count - 1)
                Return i.Primary
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim Ketemu As Boolean = False
        Dim GT1, GT2, GT3, GT4, GT5, GT6, GT As Decimal
        Dim EndLicenseDate As Date = #1/1/1950#

        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpPeriode.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpPeriode.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If 

            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()
            DTA.ClearBeforeFill = False
            DTA2.ClearBeforeFill = False

            Me.IsiTBL()

            'tentukan periode dan payment terms
            Me.Periode1 = Me.dtpPeriode.Value
            Me.Periode2 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriode.Value + TimeSpan.FromDays(32))
            Me.Periode3 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriode.Value + TimeSpan.FromDays(62))
            Me.Periode4 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriode.Value + TimeSpan.FromDays(92))
            Me.Periode5 = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriode.Value + TimeSpan.FromDays(122))

            For Each row In Me.TBL.Rows
                If row.Due_Date.Date >= Me.Periode5.Date Then
                    row.Periode6 = row.Balance
                    row.Periode1 = 0
                    row.Periode2 = 0
                    row.Periode3 = 0
                    row.Periode4 = 0
                    row.Periode5 = 0
                ElseIf row.Due_Date.Date >= Me.Periode4.Date Then
                    row.Periode5 = row.Balance
                    row.Periode1 = 0
                    row.Periode2 = 0
                    row.Periode3 = 0
                    row.Periode4 = 0
                    row.Periode6 = 0
                ElseIf row.Due_Date.Date >= Me.Periode3.Date Then
                    row.Periode4 = row.Balance
                    row.Periode1 = 0
                    row.Periode2 = 0
                    row.Periode6 = 0
                    row.Periode3 = 0
                    row.Periode5 = 0
                ElseIf row.Due_Date.Date >= Me.Periode2.Date Then
                    row.Periode3 = row.Balance
                    row.Periode1 = 0
                    row.Periode2 = 0
                    row.Periode6 = 0
                    row.Periode4 = 0
                    row.Periode5 = 0
                ElseIf row.Due_Date.Date >= Me.Periode1.Date Then
                    row.Periode2 = row.Balance
                    row.Periode1 = 0
                    row.Periode6 = 0
                    row.Periode3 = 0
                    row.Periode4 = 0
                    row.Periode5 = 0
                Else
                    row.Periode1 = row.Balance
                    row.Periode6 = 0
                    row.Periode2 = 0
                    row.Periode3 = 0
                    row.Periode4 = 0
                    row.Periode5 = 0
                End If

                row.Payment_Terms = Math.Abs((row.Invoice_Date.Date - row.Due_Date.Date).Days)
                GT1 += row.Periode1
                GT2 += row.Periode2
                GT3 += row.Periode3
                GT4 += row.Periode4
                GT5 += row.Periode5
                GT6 += row.Periode6
            Next

            'tampilkan laporan
            Me.report.SetDataSource(DS)
            Me.report1.SetDataSource(DS)
            GT = GT1 + GT2 + GT3 + GT4 + GT5 + GT6

            If Me.cboFormat.SelectedIndex = 0 Then
                Me.report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
                Me.report.SetParameterValue(report.Parameter_Grand_Total.ParameterFieldName, GT)
                Me.report.SetParameterValue(report.Parameter_GT_Periode1.ParameterFieldName, GT1)
                Me.report.SetParameterValue(report.Parameter_GT_Periode2.ParameterFieldName, GT2)
                Me.report.SetParameterValue(report.Parameter_GT_Periode3.ParameterFieldName, GT3)
                Me.report.SetParameterValue(report.Parameter_GT_Periode4.ParameterFieldName, GT4)
                Me.report.SetParameterValue(report.Parameter_GT_Periode5.ParameterFieldName, GT5)
                Me.report.SetParameterValue(report.Parameter_GT_Periode6.ParameterFieldName, GT6)
                Me.report.SetParameterValue(report.Parameter_Periode2.ParameterFieldName, Format(Me.Periode1, "MMM yy"))
                Me.report.SetParameterValue(report.Parameter_Periode3.ParameterFieldName, Format(Me.Periode2, "MMM yy"))
                Me.report.SetParameterValue(report.Parameter_Periode4.ParameterFieldName, Format(Me.Periode3, "MMM yy"))
                Me.report.SetParameterValue(report.Parameter_Periode5.ParameterFieldName, Format(Me.Periode4, "MMM yy"))
                Me.report.SetParameterValue(report.Parameter_Period.ParameterFieldName, Me.dtpPeriode.Value)
                Me.report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "Aged Trial Balance Of Receivables")
                Me.report.SetParameterValue(report.Parameter_Document.ParameterFieldName, "Invoice No")
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                    Me.Negatif = 1
                Else
                    Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                End If
                Me.report.SetParameterValue(report.Parameter_Negatif.ParameterFieldName, Me.Negatif)
                Me.report.SetParameterValue(report.Parameter_FontSize.ParameterFieldName, Me.FontNumberSize)

                Me.crv1.ReportSource = Me.report
            Else
                Me.report1.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.Def_Prec)
                Me.report1.SetParameterValue(report.Parameter_Grand_Total.ParameterFieldName, GT)
                Me.report1.SetParameterValue(report.Parameter_GT_Periode1.ParameterFieldName, GT1)
                Me.report1.SetParameterValue(report.Parameter_GT_Periode2.ParameterFieldName, GT2)
                Me.report1.SetParameterValue(report.Parameter_GT_Periode3.ParameterFieldName, GT3)
                Me.report1.SetParameterValue(report.Parameter_GT_Periode4.ParameterFieldName, GT4)
                Me.report1.SetParameterValue(report.Parameter_GT_Periode5.ParameterFieldName, GT5)
                Me.report1.SetParameterValue(report.Parameter_GT_Periode6.ParameterFieldName, GT6)
                Me.report1.SetParameterValue(report.Parameter_Periode2.ParameterFieldName, Format(Me.Periode1, "MMM yy"))
                Me.report1.SetParameterValue(report.Parameter_Periode3.ParameterFieldName, Format(Me.Periode2, "MMM yy"))
                Me.report1.SetParameterValue(report.Parameter_Periode4.ParameterFieldName, Format(Me.Periode3, "MMM yy"))
                Me.report1.SetParameterValue(report.Parameter_Periode5.ParameterFieldName, Format(Me.Periode4, "MMM yy"))
                Me.report1.SetParameterValue(report.Parameter_Period.ParameterFieldName, Me.dtpPeriode.Value)
                Me.report1.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "Aged Trial Balance Of Receivables")
                Me.report1.SetParameterValue(report.Parameter_Document.ParameterFieldName, "Invoice No")
                If Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                    Me.Negatif = 1
                Else
                    Me.Negatif = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
                End If
                Me.report1.SetParameterValue(report.Parameter_Negatif.ParameterFieldName, Me.Negatif)
                Me.report1.SetParameterValue(report.Parameter_FontSize.ParameterFieldName, Me.FontNumberSize)

                Me.crv1.ReportSource = Me.report1
            End If

            Me.crv1.Refresh()
            Me.crv1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax(ByVal vTBL As DseLaporanAging.balance_invoice_viewDataTable) As Double
        Dim row1 As DseLaporanAging.balance_invoice_viewRow
        Dim Max As Double = 0
        Try
            If vTBL.Rows.Count > 0 Then
                For Each row1 In vTBL.Rows
                    If Max < row1.Primary Then
                        Max = row1.Primary
                    End If
                Next
            Else
                Return 0
            End If

            Return Max
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FrmLaporanAging_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.dtpPeriode.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriode.Value)
            ObjCust.IsiListBoxCustomer(Me.clbCustomer, TblCust)
            Me.cboFormat.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
                Me.clbCustomer.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbCustomer.Items.Count - 1
                Me.clbCustomer.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    'aul add
    Private Function GetManualDPNumber(ByVal vPrimaryManualDP As Double, ByVal vTBLManualDP As DseLaporanAging.tbltrans_historicalbalancemanual_dpDataTable) As String
        Dim rowManualDP As DseLaporanAging.tbltrans_historicalbalancemanual_dpRow
        Dim Filter As String = ""

        Try
            Filter = ""
            Filter = "[Primary] = " & vPrimaryManualDP
            vTBLManualDP.DefaultView.RowFilter = Nothing
            vTBLManualDP.DefaultView.RowFilter = Filter
            If vTBLManualDP.DefaultView.Count > 0 Then
                rowManualDP = vTBLManualDP.DefaultView(0).Row
                Return rowManualDP.DP_No
            End If

            'kalo mpe sini maka kemablikan nilai default
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class