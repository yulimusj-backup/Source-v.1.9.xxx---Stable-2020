Option Explicit On
Option Strict Off

Public Class FrmLaporanAgingAP
    Dim DS As New DseLaporanAging
    Dim DTA As New DseLaporanAgingTableAdapters.balance_purchase_viewTableAdapter
    Dim DTA2 As New DseLaporanAgingTableAdapters.balance_purchase_return_viewTableAdapter
    Dim TBL As DseLaporanAging.balance_purchase_viewDataTable = DS.balance_purchase_view
    Dim TBL2 As DseLaporanAging.balance_purchase_return_viewDataTable = DS.balance_purchase_return_view
    Dim row As DseLaporanAging.balance_purchase_viewRow
    Dim row2 As DseLaporanAging.balance_purchase_return_viewRow
    Dim TipeBalanceSR As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Automatic
    Dim Periode5, Periode4, Periode3, Periode2, Periode1 As Date
    Dim ObjCD As New Core_Module.ClsCekData
    Dim report As New CrpLaporanAging
    Dim report1 As New CrpLaporanAging_Landscape

    Dim ObjSup As New Core_Module.clsMasterSupplier
    Dim TBLSup As New Core_Module.dseMasterSupplier.tblmaster_supplier2DataTable
    Dim DV As DataRowView
    Dim rowSup As Core_Module.dseMasterSupplier.tblmaster_supplier2Row
    Dim row_inv As DseLaporanAging.balance_invoice_viewRow

    Dim Def_Prec As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim Negatif As Integer = 0
    Dim FontNumberSize As Integer = 9
    Dim ObjLicense As New Core_Module.ClsLicense

    Private Function GetPrimaryMax() As Double
        Dim i As DseLaporanAging.balance_purchase_viewRow

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

    Private Function GetTanggalMinimum(ByVal vTBL As DseLaporanAging.tbltransaksi_kartu_ar_ap1DataTable) As Date
        Dim rowAP As DseLaporanAging.tbltransaksi_kartu_ar_ap1Row
        Dim D As Date = Date.MaxValue

        Try
            For Each rowAP In vTBL.Rows
                If D.Date < rowAP._Date.Date Then
                    D = rowAP._Date
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

    Private Function GetPRNumber(ByVal vPrimaryPR As Double, ByVal vTBLPR As DseLaporanAging.tbltrans_purchasereturnDataTable, _
   ByVal vTBLPR_Manual As DseLaporanAging.tbltrans_historicalpurchasereturnbalancemanualDataTable, ByVal vTipeTransaksi As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal) As String
        Dim rowPR As DseLaporanAging.tbltrans_purchasereturnRow
        Dim rowPR_Manual As DseLaporanAging.tbltrans_historicalpurchasereturnbalancemanualRow
        Dim filter As String = ""

        Try
            If vTipeTransaksi = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Then
                'cari dalam tbl sr manual
                filter = "Primary = " + vPrimaryPR.ToString
                vTBLPR_Manual.DefaultView.RowFilter = filter
                If vTBLPR_Manual.DefaultView.Count > 0 Then
                    rowPR_Manual = vTBLPR_Manual.DefaultView(0).Row
                    Return rowPR_Manual.AOPR_Number
                End If
            ElseIf vTipeTransaksi = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual_Transfer Or vTipeTransaksi = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Then
                'cari dalam tbl sr
                filter = "Primary = " + vPrimaryPR.ToString
                vTBLPR.DefaultView.RowFilter = filter
                If vTBLPR.DefaultView.Count > 0 Then
                    rowPR = vTBLPR.DefaultView(0).Row
                    Return rowPR.No_Transaksi
                End If
            End If


            'kalo sampe sini kembalikan default value
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetAOPNumber(ByVal vPrimaryAOP As Double, ByVal vTBLAOP As DseLaporanAging.tbltrans_aopDataTable, _
    ByVal vTBLAOP_Manual As DseLaporanAging.tbltrans_historicalpurchasebalancemanualDataTable, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal) As String
        Dim rowAOP As DseLaporanAging.tbltrans_aopRow
        Dim rowAOP_Manual As DseLaporanAging.tbltrans_historicalpurchasebalancemanualRow
        Dim Filter As String = ""

        Try
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer Or _
            vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Then
                'dari aop biasa
                Filter = "Primary = " + vPrimaryAOP.ToString
                vTBLAOP.DefaultView.RowFilter = Filter
                If vTBLAOP.DefaultView.Count > 0 Then
                    rowAOP = vTBLAOP.DefaultView(0).Row
                    If Me.chkSupplierInvoice.Checked = True Then
                        Return rowAOP.Supplier_Invoice_No
                    Else
                        Return rowAOP.No_Transaksi
                    End If
                End If
            Else
                'dari aop manual
                Filter = "Primary = " + vPrimaryAOP.ToString
                vTBLAOP_Manual.DefaultView.RowFilter = Filter
                If vTBLAOP_Manual.DefaultView.Count > 0 Then
                    rowAOP_Manual = vTBLAOP_Manual.DefaultView(0).Row
                    Return rowAOP_Manual.AOP_Number
                End If
            End If

            'kalo mpe sini maka kemablikan nilai default
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetCDNoteNumber(ByVal vPrimaryCDNote As Double, ByVal vTBLCDNote As DseLaporanAging.tbltrans_credit_noteDataTable, ByVal vTipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal) As String
        Dim rowCDNote As DseLaporanAging.tbltrans_credit_noteRow
        Dim Filter As String = ""

        Try
            If vTipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Then
                'dari Credit Note
                Filter = ""
                Filter = "[Primary] = " + vPrimaryCDNote.ToString & _
                         " And [Tipe Transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.CREDIT_NOTE & _
                         " And [TIpe Kartu] = " & ClsTransaksiCreditNote.EnumTipeKartu.Supplier
                vTBLCDNote.DefaultView.RowFilter = Nothing
                vTBLCDNote.DefaultView.RowFilter = Filter
                If vTBLCDNote.DefaultView.Count > 0 Then
                    rowCDNote = vTBLCDNote.DefaultView(0).Row
                    If Me.chkSupplierInvoice.Checked = True Then
                        Return rowCDNote.customer_debet_note
                    Else
                        Return rowCDNote.no_transksi
                    End If
                End If

            Else
                'dari Credit Note
                Filter = ""
                Filter = "[Primary] = " + vPrimaryCDNote.ToString & _
                         " And [Tipe Transaksi] = " & ClsTransaksiCreditNote.EnumTipeTransaksi.DEBIT_NOTE & _
                         " And [TIpe Kartu] = " & ClsTransaksiCreditNote.EnumTipeKartu.Supplier
                vTBLCDNote.DefaultView.RowFilter = Nothing
                vTBLCDNote.DefaultView.RowFilter = Filter
                If vTBLCDNote.DefaultView.Count > 0 Then
                    rowCDNote = vTBLCDNote.DefaultView(0).Row
                    If Me.chkSupplierInvoice.Checked = True Then
                        Return rowCDNote.customer_debet_note
                    Else
                        Return rowCDNote.no_transksi
                    End If
                End If
            End If

            'kalo mpe sini maka kemablikan nilai default
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub IsiTabel()
        Dim TBLAP As New DseLaporanAging.tbltransaksi_kartu_ar_ap1DataTable
        Dim DTAAP As New DseLaporanAgingTableAdapters.tbltransaksi_kartu_ar_ap1TableAdapter
        Dim rowAP As DseLaporanAging.tbltransaksi_kartu_ar_ap1Row
        Dim TipeKartu As ClsTransaksi_Kartu_APAR.EnumJenisKartu = ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP
        Dim TipeTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Manual

        Dim TblAOP As New DseLaporanAging.tbltrans_aopDataTable
        Dim DTAAOP As New DseLaporanAgingTableAdapters.tbltrans_aopTableAdapter
        Dim TblAOP_Manual As New DseLaporanAging.tbltrans_historicalpurchasebalancemanualDataTable
        Dim DTAAOP_Manual As New DseLaporanAgingTableAdapters.tbltrans_historicalpurchasebalancemanualTableAdapter
        Dim DTAPR As New DseLaporanAgingTableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBLPR As New DseLaporanAging.tbltrans_purchasereturnDataTable
        Dim DTAPR_Manual As New DseLaporanAgingTableAdapters.tbltrans_historicalpurchasereturnbalancemanualTableAdapter
        Dim TBLPR_Manual As New DseLaporanAging.tbltrans_historicalpurchasereturnbalancemanualDataTable
        Dim MinDate, MaxDate As Date
        Dim DTACDNote As New DseLaporanAgingTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TblCDNote As New DseLaporanAging.tbltrans_credit_noteDataTable
        Dim rowCDNote As DseLaporanAging.tbltrans_credit_noteRow

        'aul add
        Dim DtaManualDP As New DseLaporanAgingTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TblManualDP As New DseLaporanAging.tbltrans_historicalbalancemanual_dpDataTable

        Try
            'kumpulkan kartu ar sampai dengan tanggal yang bersangkutan
            If Me.clbSupplier.Items.Count = Me.clbSupplier.CheckedItems.Count Then
                DTAAP.Fill(TBLAP, TipeKartu, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriode.Value))
            Else
                DTAAP.ClearBeforeFill = False
                For Each DV In Me.clbSupplier.CheckedItems
                    rowSup = DV.Row
                    DTAAP.FillBy_PrimarySupplier(TBLAP, TipeKartu, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriode.Value), rowSup.Primary)
                Next
            End If

            'isikan data yang akan digunakan untuk perhitung
            MinDate = Me.GetTanggalMinimum(TBLAP)
            MaxDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriode.Value)
            DTAAOP.Fill(TblAOP, MaxDate, MinDate)
            DTAAOP_Manual.Fill(TblAOP_Manual, MaxDate, MinDate)
            DTAAOP_Manual.Fill(TblAOP_Manual, MaxDate, MinDate)
            DTAPR.Fill(TBLPR, MaxDate, MinDate)
            DTAPR_Manual.Fill(TBLPR_Manual, MaxDate, MinDate)
            DTACDNote.FillBy(TblCDNote, MinDate, MaxDate)
            DtaManualDP.Fill(TblManualDP, ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP, MinDate, MaxDate)

            'pindahkan ke tabel cetak
            Me.TBL.Rows.Clear()
            For Each rowAP In TBLAP.Rows
                Me.row = Me.TBL.Newbalance_purchase_viewRow
                'yang dimasukkan ke dalam balance, hanya invoice2/sr2 yang belum selesai pembayarannya
                With Me.row
                    .Primary = Me.GetPrimaryMax + 1
                    .Balance = rowAP.Saldo
                    .Supplier = rowAP.Supplier
                    .Due_Date = rowAP._Date + TimeSpan.FromDays(rowAP.Payment_Terms_Transaksi_Acuan)
                    .AOP_Date = rowAP._Date

                    TipeTransaksiAsal = rowAP.Type_Transaksi_Asal
                    If TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Manual_Transfer Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis Then
                        .AOP_Number = Me.GetAOPNumber(rowAP.Primary_Transaksi_Asal, TblAOP, TblAOP_Manual, TipeTransaksiAsal)
                    ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Manual_Transfer Or _
                    TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.PR_Otomatis Then
                        .AOP_Number = Me.GetPRNumber(rowAP.Primary_Transaksi_Asal, TBLPR, TBLPR_Manual, TipeTransaksiAsal)
                    ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note Or _
                     TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note Then
                        .AOP_Number = Me.GetCDNoteNumber(rowAP.Primary_Transaksi_Asal, TblCDNote, TipeTransaksiAsal)
                    ElseIf TipeTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_DP_Manual Then 'aul add
                        .AOP_Number = Me.GetManualDPNumber(rowAP.Primary_Transaksi_Asal, TblManualDP)
                    End If
                    .Payment_Terms = rowAP.Payment_Terms_Transaksi_Acuan
                    .Periode1 = 0
                    .Periode2 = 0
                    .Periode3 = 0
                    .Periode4 = 0
                    .Periode5 = 0
                    .Periode6 = 0
                    .Primary_Supplier = rowAP.Primary_Supplier
                    .Primary_AOP = rowAP.Primary_Transaksi_Asal
                    .Remarks = ""
                    .Tipe = 0

                    If .IsAOP_NumberNull = True Then
                        MsgBox("")
                    End If
                End With
                Me.TBL.Rows.Add(Me.row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAP.Dispose()
            TBLAP.Dispose()

            TblAOP.Dispose()
            DTAAOP.Dispose()
            TblAOP_Manual.Dispose()
            DTAAOP_Manual.Dispose()
            DTAPR.Dispose()
            TBLPR.Dispose()
            DTAPR_Manual.Dispose()
            TBLPR_Manual.Dispose()
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim Ketemu As Boolean = False
        Dim GT1, GT2, GT3, GT4, GT5, GT6, GT As Decimal
        Dim PrimaryMin, PrimaryMax As Double
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

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue

            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()
            DTA.ClearBeforeFill = False
            DTA2.ClearBeforeFill = False

            Me.IsiTabel()

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

                row.Payment_Terms = Math.Abs((row.AOP_Date.Date - row.Due_Date.Date).Days)
                GT1 += row.Periode1
                GT2 += row.Periode2
                GT3 += row.Periode3
                GT4 += row.Periode4
                GT5 += row.Periode5
                GT6 += row.Periode6
            Next

            'pindahkan ke table invoice, report merujuk pada tabel invoice
            DS.balance_invoice_view.Rows.Clear()
            For Each row In Me.TBL.Rows
                row_inv = DS.balance_invoice_view.Newbalance_invoice_viewRow
                With row_inv
                    .Balance = row.Balance
                    .Customer = row.Supplier
                    .Due_Date = row.Due_Date
                    .Invoice_Date = row.AOP_Date
                    .Invoice_Number = row.AOP_Number
                    .Payment_Terms = row.Payment_Terms
                    .Periode1 = row.Periode1
                    .Periode2 = row.Periode2
                    .Periode3 = row.Periode3
                    .Periode4 = row.Periode4
                    .Periode5 = row.Periode5
                    .Periode6 = row.Periode6
                    .Primary = row.Primary
                    .Primary_Customer = row.Primary_Supplier
                    .Primary_Invoice = row.Primary_AOP
                    .Remarks = row.Remarks
                    .Tipe = row.Tipe
                End With
                DS.balance_invoice_view.Rows.Add(row_inv)
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
                Me.report.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "Aged Trial Balance Of Payables")
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
                Me.report1.SetParameterValue(report.Parameter_Judul_Laporan.ParameterFieldName, "Aged Trial Balance Of Payables")
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

    Private Function GetPrimaryMax(ByVal vTBL As DseLaporanAging.balance_purchase_viewDataTable) As Double
        Dim row1 As DseLaporanAging.balance_purchase_viewRow
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

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmLaporanAgingAP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.cboFormat.SelectedIndex = 0
            Me.Def_Prec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.dtpPeriode.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriode.Value)
            ObjSup.IsiCheckListBox(Me.clbSupplier, Me.TBLSup)
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