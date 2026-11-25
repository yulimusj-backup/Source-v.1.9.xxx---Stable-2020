Option Explicit On
Option Strict Off

Public Class FrmCetakFakturDP
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesInvoice

    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSO As New ClsTransaksiSalesOrder
    Public FormatDecimal As Integer = 2
    Dim FormatQty As Integer = 0
    Dim ObjCust As New Core_Module.clsMasterCustomer
    Dim ObjCompany As New Core_Module.ClsMasterCompanyData
    Dim ObjCurrency As New Core_Module.clsMasterCurrency

    Public primary As Double = -1
    Public CurrencyCode As String = ""
    Public TotalHargaJual As Double = 0
    Public TotalPPN As Double = 0
    Public TotalDitagih As Double = 0
    Public JenisFaktur As EnumJenisFakturPajak = EnumJenisFakturPajak.Standar
    Public NoFakturPajak As String = ""
    Public UserCode As String = ""
    Public FakturPajakStandarValas As Boolean = False
    Public RatePajak As Double = 1
    Public DecimalValas As Double = 0
    Public CurrencyCodeValas As String = ""
    Public PersenPajak As Double = 10
    Public mParentForm As FrmTransaksiSalesInvoiceDP

    Public Enum EnumJenisFakturPajak
        Standar
        Sederhana
    End Enum
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Cetak1()
        Dim DS As New DseCetakTransaksiSalesInvoice
        Dim DTA As New DseCetakTransaksiSalesInvoiceTableAdapters.TblCetak_VAT_ValasTableAdapter
        Dim report As New CrpVAT_Uang_Muka
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowHeader As DseCetakTransaksiSalesInvoice.TblCetak_VAT_ValasRow
        Dim rowItem As DseCetakTransaksiSalesInvoice.TblCetak_VAT_ValasRow
        Dim SubTotal As Double = 0
        Dim VDec As Integer = Me.DecimalValas
        Dim HargaJualValas, UangMukaValas, DasarPengenaanPajakValas, PPNValas As Double
        Dim DTAReturn As New DseTransaksiSalesInvoiceTableAdapters.DataTable9TableAdapter
        Dim TBLReturn As New DseTransaksiSalesInvoice.DataTable9DataTable
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim NoKMK As String = ""
        Dim TglKMK As Date = #1/1/1950#
        Dim ObjTax As New ClsTransaksiSalesInvoiceTaxRate
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim CustPONumber As String = ""
        'update NPWP 16 Digit 
        Dim NPWP_PKP As String = ""

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

            rowHeader = DS.TblCetak_VAT_Valas.NewTblCetak_VAT_ValasRow
            With rowHeader
                .No_Seri_Faktur_Pajak = mParentForm.txtNoSeriFakturPajak.Text
                .No_Faktur_Pajak = mParentForm.txtNoFakturPajak.Text
                .No_Transaksi = mParentForm.txtInvoiceNo.Text
                .Customer_Name = mParentForm.TxtCustomerName.Text
                .Tax_Add_Line2 = ""
                .Tax_Add_Line1 = ""
                Me.ObjInv.GetAlamatPajak(mParentForm.TxtPrimaryInvoice.Text, .Tax_Add_Line1, .Tax_Add_Line2)
                .No_NPWP = mParentForm.txtNoNPWP.Text
                .Product_Code = ""
                .Nett_Price = Me.TotalHargaJual
                .Total_Qty = 0
                .Unit_Code = ""
                .Amount = Me.TotalHargaJual
                .Sub_Total_Before_DP = Me.TotalHargaJual
                .Down_Payment = 0
                .Sub_Total = Me.TotalHargaJual
                .Tax_Amount = mParentForm.txtTaxAmount.Text
                .Tax_Exchange_Rate = Me.mParentForm.txtTaxExchangeRate.Text
                .Currency_Code = Me.mParentForm.txtCurrencyCode.Text
                .Tanggal_Transaksi = Me.mParentForm.dtpDate.Value
                CustPONumber = Me.ObjSO.GetCustPONumber(Me.mParentForm.txtSONo.Text)
                If CustPONumber = "" Then
                    .Description = "Down payment for SO: " + Me.mParentForm.txtSONo.Text
                Else
                    .Description = "Down payment for PO: " + CustPONumber
                End If
                .Primary_Currency = Me.mParentForm.txtPrimaryCurrency.Text
                .Tax = Me.mParentForm.txtTax.Text
                .Primary_SO = Me.mParentForm.txtPrimarySO.Text
                .Primary_DO = -1
                .Primary_Product = -1
                .DefaultPrecItem = Me.DecimalValas
                .City = Me.ObjCompany.GetCompany_City
                .Nama_Signed = mParentForm.TxtNamaSigned.Text
                .Jabatan_Signed = mParentForm.TxtJabatanSigned.Text
                .Primary_Units_Multi = -1
                'asumsi faktu pajak selalu menggunakan 0 decimal
                ._Decimal = DefPrecBase
            End With
            DS.TblCetak_VAT_Valas.Rows.Add(rowHeader)

            'Update NPWP 16 Digit
            Dim objSettingNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.mParentForm.dtpDate.Value.Date >= objSettingNPWP.Tglberlaku.Date Then
                NPWP_PKP = ObjCompany.GetCompany_NPWP
            Else
                NPWP_PKP = ObjCompany.GetCompany_NPWP_Old
            End If

            'lokasi yang mngeluarkan faktur pajak, secara default adalah cikarang
            For Each rowHeader In DS.TblCetak_VAT_Valas.Rows
                rowHeader.City = Me.ObjCompany.GetCompany_City
                'untuk PPN tidak dipungut, hitung kembali nilai tax amountnya
                If rowHeader.Tax > 0 And rowHeader.Tax_Amount = 0 Then
                    rowHeader.Tax_Amount = Math.Round(Math.Round(CDec(rowHeader.Sub_Total) * _
                    CDec(rowHeader.Tax_Exchange_Rate), DefPrecBase, MidpointRounding.AwayFromZero) * _
                    CDec(rowHeader.Tax) / 100, DefPrecBase, MidpointRounding.AwayFromZero)
                End If
            Next

            'cari no KMK dan tanggal KMK
            NoKMK = ObjTax.GetKMKNumber(rowHeader.Tanggal_Transaksi)
            TglKMK = ObjTax.GetKMKDate(rowHeader.Tanggal_Transaksi)

            If Me.FakturPajakStandarValas = True Then
                'jika faktur pajak valas maka
                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    rowItem.Amount = Math.Round(rowItem.Nett_Price * 1 * Me.RatePajak, _
                    DefPrecBase, MidpointRounding.AwayFromZero)
                    rowItem.Nett_Price = Math.Round(rowItem.Nett_Price * 1, _
                    VDec, MidpointRounding.AwayFromZero)
                    HargaJualValas += rowItem.Nett_Price
                    rowItem.Total_Qty = 0
                Next
                'hitung ulang sub total, nilai VAT
                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    SubTotal += rowItem.Amount
                Next
                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    rowItem.Sub_Total_Before_DP = SubTotal
                    rowItem.Sub_Total = rowItem.Sub_Total_Before_DP - rowItem.Down_Payment
                    rowItem.Tax_Amount = Math.Round(Me.mParentForm.txtSubTotal.Text * Me.mParentForm.txtTax.Text * _
                    Me.mParentForm.txtTaxExchangeRate.Text / 100, DefPrecBase, MidpointRounding.AwayFromZero)
                Next
            End If

            UangMukaValas = 0
            DasarPengenaanPajakValas = HargaJualValas
            PPNValas = Math.Round(HargaJualValas * Me.PersenPajak / 100, _
            VDec, MidpointRounding.AwayFromZero)

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_TipeVAT.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, DefPrecBase)
            report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, Me.FormatQty)
            report.SetParameterValue(report.Parameter_MataUangAsing.ParameterFieldName, Me.CurrencyCodeValas)
            report.SetParameterValue(report.Parameter_RatePajak.ParameterFieldName, Me.ObjCD.SetDecimal(Me.RatePajak, Me.DecimalValas))
            report.SetParameterValue(report.Parameter_Valas.ParameterFieldName, Me.FakturPajakStandarValas)
            report.SetParameterValue(report.Parameter_FormatDecimalValas.ParameterFieldName, Me.DecimalValas)
            report.SetParameterValue(report.Parameter_DPValas.ParameterFieldName, 0)
            report.SetParameterValue(report.Parameter_DasarPengenaanPajakValas.ParameterFieldName, CDec(Me.mParentForm.txtSubTotal.Text))
            report.SetParameterValue(report.Parameter_HargaJualValas.ParameterFieldName, CDec(Me.mParentForm.txtSubTotal.Text))
            report.SetParameterValue(report.Parameter_PPNValas.ParameterFieldName, Math.Round(Me.TotalHargaJual * _
            CDbl(Me.mParentForm.txtTax.Text) / 100, CInt(Me.DecimalValas), MidpointRounding.AwayFromZero))
            report.SetParameterValue(report.Parameter_TanggalVAT.ParameterFieldName, ObjCD.GetDateString_Indo_Panjang(mParentForm.dtpDate.Value))
            report.SetParameterValue(report.Parameter_NamaPKP.ParameterFieldName, ObjCompany.GetCompanyName)
            report.SetParameterValue(report.Parameter_AlamatPKP.ParameterFieldName, ObjCompany.GetCompany_TaxAdd1)
            report.SetParameterValue(report.Parameter_Alamt_PKP_2.ParameterFieldName, ObjCompany.GetCompany_TaxAdd2)
            'Update NPWP 16 Digit
            'report.SetParameterValue(report.Parameter_NPWP_PKP.ParameterFieldName, ObjCompany.GetCompany_NPWP)
            report.SetParameterValue(report.Parameter_NPWP_PKP.ParameterFieldName, NPWP_PKP)
            report.SetParameterValue(report.Parameter_TglPengukuhan.ParameterFieldName, Me.ObjCD.GetDateString_Indo(ObjCompany.GetCompany_TanggapPengukuhanPKP))
            report.SetParameterValue(report.Parameter_NoKMK.ParameterFieldName, NoKMK)
            report.SetParameterValue(report.Parameter_TglKMK.ParameterFieldName, TglKMK)
            report.SetParameterValue(report.Parameter_DefPrecIDR.ParameterFieldName, DefPrecBase)
            report.SetParameterValue(report.Parameter_PotonganHarga.ParameterFieldName, 0)

            Me.CRVfaktur.ReportSource = report
            Me.CRVfaktur.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
            ObjCD = Nothing
            ObjCompany = Nothing
            ObjTax = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub FrmCetakFaktur_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.FormatDecimal = Me.ObjCurrency.GetDecimal_DefaultCurrency

            If Me.JenisFaktur = EnumJenisFakturPajak.Standar Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool) = "1" Then
                    Me.ObjCustom.CetakCustomVAT_DP(Me.DecimalValas, mParentForm.txtNoSeriFakturPajak.Text, mParentForm.txtNoFakturPajak.Text, _
                    mParentForm.txtInvoiceNo.Text, mParentForm.TxtCustomerName.Text, mParentForm.TxtPrimaryInvoice.Text, mParentForm.txtNoNPWP.Text, _
                    Me.TotalHargaJual, mParentForm.txtTaxAmount.Text, Me.mParentForm.txtTaxExchangeRate.Text, Me.mParentForm.txtCurrencyCode.Text, _
                    Me.mParentForm.dtpDate.Value, Me.mParentForm.txtSONo.Text, Me.mParentForm.txtPrimaryCurrency.Text, Me.mParentForm.txtTax.Text, _
                    Me.mParentForm.txtPrimarySO.Text, Me.FakturPajakStandarValas, mParentForm.TxtNamaSigned.Text, mParentForm.TxtJabatanSigned.Text, _
                    Me.mParentForm.txtTaxExchangeRate.Text, Me.mParentForm.txtSubTotal.Text, Me.mParentForm.txtTax.Text, Me.FormatDay, _
                    Me.FormatMonth, Me.FormatYears, Me.CurrencyCodeValas, Me.FormatQty, Me.CRVfaktur)
                Else
                    Me.Cetak1()
                End If
            Else
                Me.ObjInv.CetakFakturPajakSederhana(Me.CRVfaktur, Me.primary, NoFakturPajak, Me.FormatDecimal, Me.TotalHargaJual, _
                Me.TotalPPN, Me.TotalDitagih)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class