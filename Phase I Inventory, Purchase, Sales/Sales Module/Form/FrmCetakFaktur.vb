Option Explicit On
Option Strict Off

Imports CrystalDecisions.CrystalReports.Engine

Public Class FrmCetakFaktur
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim ObjInv As New ClsTransaksiSalesInvoice
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjCD As New Core_Module.ClsCekData
    Public FormatDecimal As Integer = 2
    Dim FormatQty As Integer = 0
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
    Public TanggalStr As String = ""
    Public JenisFakturPajakStandar As EnumJenisFakturStandar = EnumJenisFakturStandar.HargaJual

    Public Enum EnumJenisFakturStandar
        HargaJual
        Penggantian
        UangMuka
        Termin
    End Enum

    Public Enum EnumJenisFakturPajak
        Standar
        Sederhana
    End Enum
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub KurangkanItemInvoice(ByVal rowReturn As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row, _
   ByVal vTBLIsi As DseCetakTransaksiSalesInvoice.TblCetak_VAT_ValasDataTable)
        Dim rowInv As DseCetakTransaksiSalesInvoice.TblCetak_VAT_ValasRow
        Dim counter As Integer = 0
        Dim HargaLama, HargaBaru As Decimal

        Try
            While rowReturn.Total_Qty > 0
                While counter < vTBLIsi.Rows.Count
                    rowInv = vTBLIsi.Rows(counter)
                    If rowInv.Product_Code = rowReturn.Product_Code And rowInv.Primary_SO = rowReturn.Primary_SO And _
                    rowInv.Primary_DO = rowReturn.Primary_DO Then
                        If rowInv.Total_Qty = rowReturn.Total_Qty Then
                            vTBLIsi.Rows.Remove(rowInv)
                            Exit Sub
                        ElseIf rowInv.Total_Qty < rowReturn.Total_Qty Then
                            rowReturn.Total_Qty -= rowInv.Total_Qty
                            vTBLIsi.Rows.Remove(rowInv)
                        ElseIf rowInv.Total_Qty > rowReturn.Total_Qty Then
                            rowInv.Total_Qty -= rowReturn.Total_Qty
                            'hitung ulang amount
                            HargaLama = rowInv.Amount
                            HargaBaru = Math.Round(rowInv.Amount * rowInv.Total_Qty / _
                            (rowInv.Total_Qty + rowReturn.Total_Qty), rowInv.DefaultPrecItem, MidpointRounding.AwayFromZero)
                            rowInv.Amount = HargaBaru
                            Exit Sub
                        End If
                    Else
                        counter += 1
                    End If
                End While
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak1()
        Dim DS As New DseCetakTransaksiSalesInvoice
        Dim DTA As New DseCetakTransaksiSalesInvoiceTableAdapters.TblCetak_VAT_ValasTableAdapter
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowHeader As DseCetakTransaksiSalesInvoice.TblCetak_VAT_ValasRow
        Dim rowItem, rowItemDummy As DseCetakTransaksiSalesInvoice.TblCetak_VAT_ValasRow
        Dim SubTotal As Double = 0
        Dim VDec As Integer = Me.DecimalValas
        Dim HargaJualValas, UangMukaValas, DasarPengenaanPajakValas, PPNValas As Double
        Dim DTAReturn As New DseTransaksiSalesInvoiceTableAdapters.DataTable9TableAdapter
        Dim TBLReturn As New DseTransaksiSalesInvoice.DataTable9DataTable
        Dim rowReturn As DseTransaksiSalesInvoice.DataTable9Row
        Dim AlamatPajak1, AlamatPajak2 As String
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim City As String = ""
        Dim dummy As Double = 0
        Dim NoKMK As String = ""
        Dim TglKMK As Date = #1/1/1950#
        Dim ObjTax As New ClsTransaksiSalesInvoiceTaxRate
        Dim Ongkir As Decimal = 0
        Dim DefPrecBase As Integer = 0
        'Update NPWP 16 Digit
        Dim NPWP_PKP As String = ""

        Try
            DefPrecBase = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.FormatDecimal = DefPrecBase
            DTA.FillBy_PrimaryCurrency(DS.TblCetak_VAT_Valas, primary, CurrencyCode)

            'kurangkan isi yang terjadi dengan sales return
            If DTAReturn.FillBy_PrimaryInvoice(TBLReturn, Me.primary) > 0 Then
                Dim DTAReturn_Detail As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
                Dim TBLReturn_Detail As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
                Dim rowReturnDetail As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row

                For Each rowReturn In TBLReturn.Rows
                    DTAReturn_Detail.FillBy_PrimaryTransaksi(TBLReturn_Detail, rowReturn.Primary_SR)
                    For Each rowReturnDetail In TBLReturn_Detail.Rows
                        Me.KurangkanItemInvoice(rowReturnDetail, DS.TblCetak_VAT_Valas)
                    Next
                Next

                DTAReturn_Detail.Dispose()
                TBLReturn_Detail.Dispose()
            End If

            'tambahkan ongkos kirim
            Ongkir = Me.ObjInv.GetOngkosKirim(primary)
            If Ongkir <> 0 And Me.ObjInv.IsShipmentWithTax(primary) = True Then
                rowItem = DS.TblCetak_VAT_Valas.NewTblCetak_VAT_ValasRow
                rowItemDummy = DS.TblCetak_VAT_Valas.Rows(0)

                With rowItem
                    .Amount = Ongkir
                    .City = rowItemDummy.City
                    .Currency_Code = rowItemDummy.Currency_Code
                    .Customer_Name = rowItemDummy.Customer_Name
                    .DefaultPrecItem = rowItemDummy.DefaultPrecItem
                    .Description = "Ongkos Kirim"
                    .Down_Payment = rowItemDummy.Down_Payment
                    .Jabatan_Signed = rowItemDummy.Jabatan_Signed
                    .Nama_Signed = rowItemDummy.Nama_Signed
                    .Nett_Price = Ongkir
                    .No_Faktur_Pajak = rowItemDummy.No_Faktur_Pajak
                    .No_NPWP = rowItemDummy.No_NPWP
                    .No_Seri_Faktur_Pajak = rowItemDummy.No_Seri_Faktur_Pajak
                    .No_Transaksi = rowItemDummy.No_Transaksi
                    .Primary_Currency = rowItemDummy.Primary_Currency
                    .Primary_DO = -1
                    .Primary_Product = -1
                    .Primary_SO = -1
                    .Product_Code = "OK"
                    .Sub_Total = rowItemDummy.Sub_Total
                    .Sub_Total_Before_DP = rowItemDummy.Sub_Total_Before_DP
                    .Tanggal_Transaksi = rowItemDummy.Tanggal_Transaksi
                    .Tax = rowItemDummy.Tax
                    .Tax_Add_Line1 = rowItemDummy.Tax_Add_Line1
                    .Tax_Add_Line2 = rowItemDummy.Tax_Add_Line2
                    .Tax_Amount = rowItemDummy.Tax_Amount
                    .Tax_Exchange_Rate = rowItemDummy.Tax_Exchange_Rate
                    .Total_Qty = 1
                    .Unit_Code = "UNIT"
                    ._Decimal = 0
                    .disc_amount = rowItemDummy.disc_amount
                End With

                DS.TblCetak_VAT_Valas.Rows.Add(rowItem)

                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    rowItem.Sub_Total_Before_DP += Ongkir
                Next
            End If

            If Me.ObjInv.GetTanggalTransaksi(primary) < #4/1/2013# Then
                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    rowItem.No_Seri_Faktur_Pajak = rowItem.No_Seri_Faktur_Pajak + "-"
                Next
            End If

            'lokasi yang mngeluarkan faktur pajak, secara default adalah cikarang
            City = ObjCompany.GetCompany_City
            For Each rowHeader In DS.TblCetak_VAT_Valas.Rows
                rowHeader.City = City
            Next

            'update data alamat pajak
            AlamatPajak1 = ""
            AlamatPajak2 = ""
            Me.ObjInv.GetAlamatPajak(primary, AlamatPajak1, AlamatPajak2)
            For Each rowHeader In DS.TblCetak_VAT_Valas.Rows
                rowHeader.Tax_Add_Line1 = AlamatPajak1
                rowHeader.Tax_Add_Line2 = AlamatPajak2
                'untuk PPN tidak dipungut, hitung kembali nilai tax amountnya
                If rowHeader.Tax > 0 And rowHeader.Tax_Amount = 0 Then
                    rowHeader.Tax_Amount = Math.Round(Math.Round(CDec(rowHeader.Sub_Total) * _
                    CDec(rowHeader.Tax_Exchange_Rate), DefPrecBase, MidpointRounding.AwayFromZero) * _
                    CDec(rowHeader.Tax) / 100, DefPrecBase, MidpointRounding.AwayFromZero)
                End If
            Next

            'update data invoice IDR
            For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                rowItem.disc_amount_IDR = Math.Round(rowItem.disc_amount * RatePajak, _
                DefPrecBase, MidpointRounding.AwayFromZero)
                rowItem.Sub_Total -= rowItem.disc_amount_IDR
            Next

            'Update NPWP 16 Digit
            Dim objSettingNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.ObjInv.GetTanggalTransaksi(primary).Date >= objSettingNPWP.Tglberlaku.Date Then
                NPWP_PKP = ObjCompany.GetCompany_NPWP
            Else
                NPWP_PKP = ObjCompany.GetCompany_NPWP_Old
            End If

            If Me.FakturPajakStandarValas = True Then
                'jika faktur pajak valas maka
                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    rowItem.Nett_Price = Math.Round(rowItem.Amount, _
                    VDec, MidpointRounding.AwayFromZero)
                    rowItem.Amount = Math.Round(rowItem.Amount * RatePajak, _
                    DefPrecBase, MidpointRounding.AwayFromZero)
                    HargaJualValas += rowItem.Nett_Price
                Next

                'hitung ulang sub total, nilai VAT
                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    SubTotal += rowItem.Amount
                Next
                For Each rowItem In DS.TblCetak_VAT_Valas.Rows
                    rowItem.Sub_Total_Before_DP = SubTotal
                    rowItem.Sub_Total = rowItem.Sub_Total_Before_DP - rowItem.Down_Payment - rowItem.disc_amount
                    rowItem.Tax_Amount = Math.Round(rowItem.Sub_Total * 0.1, DefPrecBase, MidpointRounding.AwayFromZero)
                Next
                rowHeader = DS.TblCetak_VAT_Valas.Rows(0)
                UangMukaValas = rowHeader.Down_Payment
                dummy = Math.Round(UangMukaValas * Me.RatePajak, DefPrecBase, MidpointRounding.AwayFromZero)
                DasarPengenaanPajakValas = HargaJualValas - UangMukaValas - rowHeader.disc_amount
                PPNValas = Math.Round(DasarPengenaanPajakValas * Me.PersenPajak / 100, _
                VDec, MidpointRounding.AwayFromZero)

                'cari no KMK dan tanggal KMK
                NoKMK = ObjTax.GetKMKNumber(rowHeader.Tanggal_Transaksi)
                TglKMK = ObjTax.GetKMKDate(rowHeader.Tanggal_Transaksi)

                For Each rowHeader In DS.TblCetak_VAT_Valas.Rows
                    rowHeader.Down_Payment = dummy
                    rowHeader.Sub_Total = Math.Round(DasarPengenaanPajakValas * Me.RatePajak, DefPrecBase, MidpointRounding.AwayFromZero)
                    rowHeader.Tax_Amount = Math.Round((rowHeader.Sub_Total * Me.PersenPajak) / 100, DefPrecBase, MidpointRounding.AwayFromZero)
                    rowHeader.Sub_Total_Before_DP = Math.Round(HargaJualValas * Me.RatePajak, DefPrecBase, MidpointRounding.AwayFromZero)
                Next
            Else
                UangMukaValas = 0
            End If

            If Me.JenisFakturPajakStandar = EnumJenisFakturStandar.HargaJual Then
                Dim report As New CrpVAT_HargaJual
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
                report.SetParameterValue(report.Parameter_TipeVAT.ParameterFieldName, 0)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, Me.FormatQty)
                report.SetParameterValue(report.Parameter_MataUangAsing.ParameterFieldName, Me.CurrencyCodeValas)
                report.SetParameterValue(report.Parameter_RatePajak.ParameterFieldName, Me.ObjCD.SetDecimal(Me.RatePajak, Me.DecimalValas))
                report.SetParameterValue(report.Parameter_Valas.ParameterFieldName, Me.FakturPajakStandarValas)
                report.SetParameterValue(report.Parameter_FormatDecimalValas.ParameterFieldName, Me.DecimalValas)
                report.SetParameterValue(report.Parameter_DPValas.ParameterFieldName, UangMukaValas)
                report.SetParameterValue(report.Parameter_DasarPengenaanPajakValas.ParameterFieldName, DasarPengenaanPajakValas)
                report.SetParameterValue(report.Parameter_HargaJualValas.ParameterFieldName, HargaJualValas)
                report.SetParameterValue(report.Parameter_PPNValas.ParameterFieldName, PPNValas)
                report.SetParameterValue(report.Parameter_TanggalVAT.ParameterFieldName, TanggalStr)
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

                Me.CRVfaktur.ReportSource = report
            ElseIf Me.JenisFakturPajakStandar = EnumJenisFakturStandar.Penggantian Then
                Dim report As New CrpVAT_Penggantian
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
                report.SetParameterValue(report.Parameter_TipeVAT.ParameterFieldName, 0)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, Me.FormatQty)
                report.SetParameterValue(report.Parameter_MataUangAsing.ParameterFieldName, Me.CurrencyCodeValas)
                report.SetParameterValue(report.Parameter_RatePajak.ParameterFieldName, Me.ObjCD.SetDecimal(Me.RatePajak, Me.DecimalValas))
                report.SetParameterValue(report.Parameter_Valas.ParameterFieldName, Me.FakturPajakStandarValas)
                report.SetParameterValue(report.Parameter_FormatDecimalValas.ParameterFieldName, Me.DecimalValas)
                report.SetParameterValue(report.Parameter_DPValas.ParameterFieldName, UangMukaValas)
                report.SetParameterValue(report.Parameter_DasarPengenaanPajakValas.ParameterFieldName, DasarPengenaanPajakValas)
                report.SetParameterValue(report.Parameter_HargaJualValas.ParameterFieldName, HargaJualValas)
                report.SetParameterValue(report.Parameter_PPNValas.ParameterFieldName, PPNValas)
                report.SetParameterValue(report.Parameter_TanggalVAT.ParameterFieldName, TanggalStr)
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

                Me.CRVfaktur.ReportSource = report
            ElseIf Me.JenisFakturPajakStandar = EnumJenisFakturStandar.Termin Then
                Dim report As New CrpVAT_Termin
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
                report.SetParameterValue(report.Parameter_TipeVAT.ParameterFieldName, 0)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, Me.FormatQty)
                report.SetParameterValue(report.Parameter_MataUangAsing.ParameterFieldName, Me.CurrencyCodeValas)
                report.SetParameterValue(report.Parameter_RatePajak.ParameterFieldName, Me.ObjCD.SetDecimal(Me.RatePajak, Me.DecimalValas))
                report.SetParameterValue(report.Parameter_Valas.ParameterFieldName, Me.FakturPajakStandarValas)
                report.SetParameterValue(report.Parameter_FormatDecimalValas.ParameterFieldName, Me.DecimalValas)
                report.SetParameterValue(report.Parameter_DPValas.ParameterFieldName, UangMukaValas)
                report.SetParameterValue(report.Parameter_DasarPengenaanPajakValas.ParameterFieldName, DasarPengenaanPajakValas)
                report.SetParameterValue(report.Parameter_HargaJualValas.ParameterFieldName, HargaJualValas)
                report.SetParameterValue(report.Parameter_PPNValas.ParameterFieldName, PPNValas)
                report.SetParameterValue(report.Parameter_TanggalVAT.ParameterFieldName, TanggalStr)
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

                Me.CRVfaktur.ReportSource = report
            ElseIf Me.JenisFakturPajakStandar = EnumJenisFakturStandar.UangMuka Then
                Dim report As New CrpVAT_Uang_Muka
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
                report.SetParameterValue(report.Parameter_TipeVAT.ParameterFieldName, 0)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
                report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, Me.FormatQty)
                report.SetParameterValue(report.Parameter_MataUangAsing.ParameterFieldName, Me.CurrencyCodeValas)
                report.SetParameterValue(report.Parameter_RatePajak.ParameterFieldName, Me.ObjCD.SetDecimal(Me.RatePajak, Me.DecimalValas))
                report.SetParameterValue(report.Parameter_Valas.ParameterFieldName, Me.FakturPajakStandarValas)
                report.SetParameterValue(report.Parameter_FormatDecimalValas.ParameterFieldName, Me.DecimalValas)
                report.SetParameterValue(report.Parameter_DPValas.ParameterFieldName, UangMukaValas)
                report.SetParameterValue(report.Parameter_DasarPengenaanPajakValas.ParameterFieldName, DasarPengenaanPajakValas)
                report.SetParameterValue(report.Parameter_HargaJualValas.ParameterFieldName, HargaJualValas)
                report.SetParameterValue(report.Parameter_PPNValas.ParameterFieldName, PPNValas)
                report.SetParameterValue(report.Parameter_TanggalVAT.ParameterFieldName, TanggalStr)
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

                Me.CRVfaktur.ReportSource = report
            End If


            Me.CRVfaktur.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTAReturn.Dispose()
            ObjCompany = Nothing
            TBLReturn.Dispose()
            ObjTax = Nothing
        End Try
    End Sub

    Private Sub FrmCetakFaktur_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ObjInv As New CustomSMESS.ClsTransaksiSalesInvoice

        Try
            Me.FormatDecimal = Me.ObjCur.GetDecimal_DefaultCurrency
            If Me.JenisFaktur = EnumJenisFakturPajak.Standar Then
                If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomTaxInvoice_Bool) = "1" Then
                    ObjInv.CetakCustomVAT(Me.DecimalValas, Me.primary, Me.CurrencyCode, Me.FakturPajakStandarValas, Me.RatePajak, Me.PersenPajak, _
                    Me.JenisFakturPajakStandar, Me.FormatDay, Me.FormatMonth, Me.FormatYears, Me.FormatQty, Me.CurrencyCodeValas, Me.FormatDecimal, _
                    Me.TanggalStr, Me.ObjCD, Me.CRVfaktur, Me.ObjInv.IsShipmentWithTax(Me.primary))
                Else
                    Me.Cetak1()
                End If
            Else
                Me.ObjInv.CetakFakturPajakSederhana(Me.CRVfaktur, Me.primary, NoFakturPajak, Me.FormatDecimal, Me.TotalHargaJual, _
                Me.TotalPPN, Me.TotalDitagih)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjInv = Nothing
        End Try
    End Sub
End Class