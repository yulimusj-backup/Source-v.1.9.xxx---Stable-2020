Option Explicit On
Option Strict Off

Imports Core_Module
Imports CustomSMESS
Public Class FrmTransaksiSalesInvoiceQuick
    Dim mTblMultiUnit As New DseTransaksiSO.tblmaster_product_item_multiunitDataTable
    Dim tblCboUnit As New DseTransaksiSO.CboUnitDataTable
    Dim ObjUnit As New Core_Module.ClsMasterUnit

    Dim objCustom As New CustomSMESS.ClsTransaksiSalesInvoice

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public CheckLicense As Boolean = True

    Dim baru As Boolean
    Dim ObjInvoice As New ClsTransaksiSalesInvoice
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjCustomer_sProductCode As New clsTransaksiSettingcustomerProductCodeNo
    Dim objDO As New ClsTransaksiDO
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjSalesPerson As New Core_Module.ClsMasterSalesPerson
    Dim ObjInvoiceItem As New ClsTransaksiSalesInvoiceItem
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSR As New ClsTransaksiSalesReturn
    Dim ObjSRItem As New ClsTransaksiSalesReturnItem
    Dim ObjStockOutSales As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjLicense As New ClsLicense
    Dim ObjLocation As New clsMasterStorageArea
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjTaxInvoiceSetting As New ClsTransaksiSalesInvoice_SettingTaxInvoice
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate

    Dim DefaultPrecision As Integer = 0
    Dim DefaultPrecision_DefCur As Integer = 0

    Dim buffer As String = ""
    Dim SedangRevisi As Boolean = False
    Dim TBL As New DseTransaksiDO.tbltrans_do_view2DataTable
    Dim TBL2 As New DseTransaksiSalesInvoice.DataTable7DataTable
    Dim TBLSR As New dseTransaksiSalesReturn.DataTable3DataTable
    ''' <summary>
    ''' Perhatikan!!! KALO PAKE METODE UNIT PRICE INCLUDE ATAU SO DILAPORAN SALES ANALISIS 
    ''' DAN SALES STATISTIK AKAN NGACO PERHITUNGANNYA
    ''' </summary>
    ''' <remarks></remarks>
    'Public MetodeHitungInclude As EnumHitungInclude = EnumHitungInclude.NettPriceInclude

    Dim TBL_Item As New Core_Module.dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2_Item As New Core_Module.dseMasterProduct.stock_location_viewDataTable
    Dim TBL3_Item As New Core_Module.dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4_Item As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item2DataTable
    Dim TBL5_Item As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item3DataTable

    Dim Buffer2 As String = ""

    'terkait dengan transfer quick invoice
    Dim listOfTransfered As New List(Of String)
    Dim listOfSkiped As New List(Of String)
    Dim listOfFailed As New List(Of String)
    Dim Msg As String = ""
    Dim Lanjut As Boolean = False
    Dim FrmErr As New FrmTransferBeginingBalance_Error
    Dim vErrorMessage As String = ""
    Dim SudahSelesaiCek As Boolean = False
    Dim vStartDate, vEndDate As Date
    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Dim DseAccount As New DseMultiAccountTableAdapters.tblisi_cbo_ar_akunTableAdapter
    Dim tblAccount As New DseMultiAccount.tblisi_cbo_ar_akunDataTable
    Dim rowAccount As DseMultiAccount.tblisi_cbo_ar_akunRow

    Dim DseCurrTax As New DseMultiAccountTableAdapters.tblCurrencyTaxTableAdapter
    Dim tblCurrTax As New DseMultiAccount.tblCurrencyTaxDataTable
    Dim rowCurrTax As DseMultiAccount.tblCurrencyTaxRow

    Public TBLHeaderCost As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataDataTable
    Public rowHeaderCost As DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataRow
    Public PrimaryCustomerShipTo As Double = -1
    Public IdentifierAddressCustomerShipTo As String = ""
    Public TBLitemCost As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costDataTable
    Public rowitemCost As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costRow
    Dim OptionNIKPASS As Core_Module.clsMasterCustomer.EnumOptionNikPass = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK
    Public TblCustomDummy As New CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable
    Dim TblCustomInOutDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable  'new
    Dim ObjCustom2 As New ClsTransaksiSalesInvoiceQuick_AddByCustomProduct

    'ab
    Public vFromQuick As Boolean = False
    Public vApakahDatabenar As Boolean = True

    '2025 Core Tax Tahap 2.1
    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
    '2025 Core Tax Tahap 2.2
    Dim SPPBDate As Date = #1/1/1950#
    Dim AddInfo As Double = -1
    Dim FacilityStamp As Double = -1

    Public Sub IsiCBOArAkunByTipePerkiraan1()
        '- isicbo customer
        Dim tp1 As New Core_Module.ClsMasterAkun1.EnumTipeAkun
        Dim tp2 As New Core_Module.ClsMasterAkun1.EnumTipeAkun2
        tp1 = Core_Module.ClsMasterAkun1.EnumTipeAkun.Asset
        tp2 = Core_Module.ClsMasterAkun1.EnumTipeAkun2.Account_Receivable
        Dim objMulti As New ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboAccount.DataSource = Nothing
            Me.cboAccount.Items.Clear()
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryAccount.Items.Clear()
            Me.tblAccount.Rows.Clear()

            objMulti.IsiTBLAPARAcount(tp1, tp2, Me.tblAccount, Me.txtPrimaryCurrency.Text)

            Me.cboAccount.DataSource = Me.tblAccount.DefaultView
            Me.cboAccount.DisplayMember = "Account Code"
            Me.cboPrimaryAccount.DataSource = Me.tblAccount.DefaultView
            Me.cboPrimaryAccount.DisplayMember = "Primary"
            Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCBOCurrencyTax()
        '- isicbo customer
        Dim objMulti As New ClsMultiAccount
        Try
            ' If Me.TxtPrimaryTransaksi.Text = "-1" Then
            Me.cboCurrTax.DataSource = Nothing
            Me.cboCurrTax.Items.Clear()
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboPrimaryCurrTax.Items.Clear()
            objMulti.IsiTBLCurrTax(Me.txtPrimaryCurrency.Text, Me.tblCurrTax)

            'Dim I As New Form5
            'I.DataGridView1.DataSource = Nothing
            'I.DataGridView1.DataSource = Me.tblAR
            'I.ShowDialog()

            Me.cboCurrTax.DataSource = Me.tblCurrTax.DefaultView
            Me.cboCurrTax.DisplayMember = "Currency Code"
            Me.cboPrimaryCurrTax.DataSource = Me.tblCurrTax.DefaultView
            Me.cboPrimaryCurrTax.DisplayMember = "Primary"
            Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/2050#
        Dim DOStr As String = ""
        Dim StockOutStr As String = ""
        Dim ObjCL As New ClsTransaksiPurchaseClosing
        Dim row As DataGridViewRow
        Dim ListPrimaryDO As New List(Of Double)
        Dim ListNoStockOut As New List(Of String)
        Dim Dummy As String = ""
        Dim ConvertedCurr As String = ""
        Dim row_dgv As DataGridViewRow
        Dim i As Integer = 0
        Dim DefTaxRate As Decimal = 0
        'Dim rowDGV As DataGridViewRow

        Try
            If Me.objCustom.ApakahDataBenarCustom(Me.dgvItem, Me.ColCodeNo.Index, Me.ColPrimaryCodeNo.Index, Me.chkTax.Checked, _
            Me.ColPriceByMultiUnit.Index, Me.ColPrimaryUnitsMulti.Index, Me.ColPrimaryUnit.Index) = False Then
                Return False
            End If

            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '~ab~
            'Remarks : tambahan batasan 
            '25/02/2020
            Dim DTA_Convt As New DseTransaksiPurchaseClosingTableAdapters.DataTable1TableAdapter
            If DTA_Convt.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA_Convt.ScalarQuery_GetConvertionDate.Value.Date > Me.dtpDate.Value Then
                    MsgBox("Please specify transaction date after convertion date", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox("Please specify transaction date after convertion date", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '

            'cek data yang harus diisi
            KoleksiNamaData.Add("invoice number")
            KoleksiTextBox.Add(Me.txtInvoiceNo)
            KoleksiNamaData.Add("customer code")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            KoleksiNamaData.Add("currency")
            KoleksiTextBox.Add(Me.txtCurrencyCode)
            KoleksiNamaData.Add("sales person code")
            KoleksiTextBox.Add(Me.TxtSalesPersonCode)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2025 Core Tax Tahap 2.2
            'berkaitan dengan data e-faktur yang tidak lengkap
            'If Me.chkTax.Checked = True Then
            '    Me.Txt_E_Faktur_Blok.Text = Me.Txt_E_Faktur_Blok.Text.Trim
            '    Me.Txt_E_Faktur_Jalan.Text = Me.Txt_E_Faktur_Jalan.Text.Trim
            '    Me.Txt_E_Faktur_Kelurahan.Text = Me.Txt_E_Faktur_Kelurahan.Text.Trim
            '    Me.Txt_E_Faktur_KodePos.Text = Me.Txt_E_Faktur_KodePos.Text.Trim
            '    Me.Txt_E_Faktur_No.Text = Me.Txt_E_Faktur_No.Text.Trim
            '    Me.Txt_E_Faktur_NoTelepon.Text = Me.Txt_E_Faktur_NoTelepon.Text.Trim
            '    Me.Txt_E_Faktur_Propinsi.Text = Me.Txt_E_Faktur_Propinsi.Text.Trim
            '    Me.Txt_E_Faktur_RT.Text = Me.Txt_E_Faktur_RT.Text.Trim
            '    Me.Txt_E_Faktur_RW.Text = Me.Txt_E_Faktur_RW.Text.Trim
            '    Me.Txt_E_Faktur_Kabupaten.Text = Me.Txt_E_Faktur_Kabupaten.Text.Trim
            '    Me.Txt_E_Faktur_Kecamatan.Text = Me.Txt_E_Faktur_Kecamatan.Text.Trim

            '    If Me.Txt_E_Faktur_Blok.Text = "" Or _
            '    Me.Txt_E_Faktur_Jalan.Text = "" Or _
            '    Me.Txt_E_Faktur_Kelurahan.Text = "" Or _
            '    Me.Txt_E_Faktur_KodePos.Text = "" Or _
            '    Me.Txt_E_Faktur_No.Text = "" Or _
            '    Me.Txt_E_Faktur_NoTelepon.Text = "" Or _
            '    Me.Txt_E_Faktur_Propinsi.Text = "" Or _
            '    Me.Txt_E_Faktur_RT.Text = "" Or _
            '    Me.Txt_E_Faktur_RW.Text = "" Or _
            '    Me.Txt_E_Faktur_Kecamatan.Text = "" Or _
            '    Me.Txt_E_Faktur_Kabupaten.Text = "" Then
            '        MsgBox("Incomplete e-Faktur data, please fix.", MsgBoxStyle.Information, "Information")
            '        Me.ShowDetailFakturPajak()
            '        If Me.Txt_E_Faktur_Blok.Text = "" Or _
            '           Me.Txt_E_Faktur_Jalan.Text = "" Or _
            '           Me.Txt_E_Faktur_Kelurahan.Text = "" Or _
            '           Me.Txt_E_Faktur_KodePos.Text = "" Or _
            '           Me.Txt_E_Faktur_No.Text = "" Or _
            '           Me.Txt_E_Faktur_NoTelepon.Text = "" Or _
            '           Me.Txt_E_Faktur_Propinsi.Text = "" Or _
            '           Me.Txt_E_Faktur_RT.Text = "" Or _
            '           Me.Txt_E_Faktur_RW.Text = "" Or _
            '           Me.Txt_E_Faktur_Kecamatan.Text = "" Or _
            '           Me.Txt_E_Faktur_Kabupaten.Text = "" Then
            '            Return False
            '        End If
            '    End If
            'End If

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjInvoice.TimeStamp <> Me.ObjInvoice.GetTimestamp(Me.ObjInvoice.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjInvoice.CariDetail(Me.ObjInvoice.Primary, Me.dgvItem, Me.dgvDP, _
                        Me.dgvItemSR, Me.dgvExchageRates, Me, Me.TblCustomDummy, Me.TblCustomInOutDummy)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Me.RboCustomerShipTo.Checked And (Me.TxtPrimaryCustomerShipTo.Text = "" Or Me.TxtPrimaryCustomerShipTo.Text = "-1" Or Me.TxtIdentifierAddressCustomerShipTo.Text = "") Then
                MsgBox("If you want to ship to other customer, you have to define a customer and address.", MsgBoxStyle.Information, "Information")
                Me.TxtCustomerCodeShipTo.Focus()
                Return False
            End If

            If Me.RboCustomerShipTo.Checked And (Me.txtPrimaryCustomer.Text = Me.TxtPrimaryCustomerShipTo.Text) Then
                MsgBox("If you want to ship to other customer, please define different customer for goods shipping.", MsgBoxStyle.Information, "Information")
                Me.TxtCustomerCodeShipTo.Focus()
                Return False
            End If

            'cek jangan sampai ada DP yang terlupakan, khusus untuk order balance yang sudah habis
            Dim ij As New List(Of Double)
            Dim DPNeeded As String = ""

            'tambahkan item DP
            Me.dgvDPDummy.Rows.Clear()
            For Each row In Me.dgvItem.Rows
                If ij.Contains(Me.dgvItem.Item(Me.ColPrimarySO.Index, row.Index).Value) = False Then
                    ij.Add(Me.dgvItem.Item(Me.ColPrimarySO.Index, row.Index).Value)
                End If
            Next
            Me.ObjInvoice.IsiDgvDP(ij, Me.dgvDPDummy, Me.txtPrimaryCurrency.Text, Me.txtPrimaryInvoice.Text)
            If Me.dgvDPDummy.Rows.Count <> Me.dgvDP.Rows.Count Then
                If MsgBox("There are down payment refering to this invoice, continue and ignore the downpayment?", _
                MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            'cek berkaitan dengan tanggal closing
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.ObjInvoice.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjInvoice.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    'Me.txtInvoiceNo.Text = Me.ObjInvoice.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjInvoice.TanggalTransaksi
                        Me.txtInvoiceNo.Text = Me.ObjInvoice.NoTransaksi
                        Return False
                    End If
                End If
            End If

            'kalo ada sales return maka berikan peringatan
            Dummy = Me.GetSR
            If Dummy <> "" Then
                If MsgBox("Sales return: " + Dummy + " refers to this invoice, ignore?", MsgBoxStyle.YesNo, _
                "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            'data salah jika, tidak ada kurs pajak untuk tgl ybs
            If Me.txtPrimaryCurrency.Text <> "1" Then
                If Me.dtpDate.Value.Date >= Me.ObjTaxInvoiceSetting.BatasStartDate Then
                    If Me.ObjTaxRate.ApakahSudahAda(Me.dtpDate.Value.Date) = False Then
                        MsgBox("Please specify tax invoice rate data current transaction date.", MsgBoxStyle.Information, "Information")
                        Me.ShowTaxSetting()
                        Return False
                    End If
                End If

                'jika kursnya tidak sama, maka berikan peringatan
                DefTaxRate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value, Me.txtPrimaryCurrency.Text)
                If CDec(Me.txtTaxExchangeRate.Text) <> DefTaxRate Then
                    If MsgBox("The tax exchange rate differs from default tax exchange rate for " + Me.txtCurrencyCode.Text + " on " + _
                    Format(Me.dtpDate.Value, "dd MMM yyyy") + " which is " + Me.ObjCD.DigitGrouping(DefTaxRate, Me.ObjTaxRate.DefaultPrec) + " continue save?", MsgBoxStyle.YesNo, _
                    "Confirm") <> MsgBoxResult.Yes Then
                        Me.txtTaxExchangeRate.Focus()
                        Return False
                    End If
                End If
            End If

            'jika tidak ada no faktur pajak standar
            If Me.CboFakturPajak.SelectedIndex = 1 Then
                If Me.txtNoFakturPajak.Text = "" And Me.chkTax.Checked = True Then
                    '2025 Core Tax Tahap 2.2
                    'Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    'If Me.txtNoFakturPajak.Text = "" Then
                    '    MsgBox("There are no new tax invoice number available. Please specify  new tax invoice number in your global setting.", _
                    '    MsgBoxStyle.Information, "Information")
                    '    Return False
                    'End If

                    'MsgBox("Please enter tax invoice number", MsgBoxStyle.Information, "Information")
                    'Return False
                End If
            End If

            'berkaitan dengan tax exchange rate
            If Me.txtPrimaryCurrency.Text = "1" Then
                If CDbl(Me.txtTaxExchangeRate.Text) <> 1 Then
                    MsgBox("For invoice with " + Me.ObjCurrency.GetCurrencyCode(1) + " currency, the tax exchange rate value must be 1.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtTaxExchangeRate.Focus()
                    Return False
                End If
            Else
                'jika bukan IDR dan tax ratenya bernilai 1 maka berikan peringatan
                If CDbl(Me.txtTaxExchangeRate.Text) = 1 Then
                    MsgBox("This invoice use non-" + Me.ObjCurrency.GetCurrencyCode(1) + " currency and the tax exchange rate value = 1. This is wrong data input.", _
                    MsgBoxStyle.YesNo, "Confirm")
                    Me.txtTaxExchangeRate.Focus()
                    Return False
                End If
            End If

            If Me.dgvExchageRates.RowCount > 0 Then
                Dummy = ""
                i = 0
                row_dgv = Me.dgvExchageRates.Rows(0)
                If Me.txtCurrencyCode.Text.ToUpper <> row_dgv.Cells(Me.ColExchangedCurrency.Index).Value.ToString.ToUpper Then
                    'berkaitan dengan exchange rate data, jika currecy yang digunakan tidak sama dengan exchange rate data, maka seharusnya nilainya
                    'tidak nol dan tidak 1
                    For Each row_dgv In Me.dgvExchageRates.Rows
                        If CDbl(row_dgv.Cells(Me.ColExchangeRate.Index).Value) = 1 Then
                            i += 1
                            If Dummy = "" Then
                                Dummy = "The following document have exchange rate value = 1:" + vbCrLf + _
                                i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            Else
                                Dummy = Dummy + i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            End If
                        End If
                    Next
                Else
                    'sebaliknya, jika mata uangnya sama tetapi kemudian exchange rate yang digunakan adalah 1 maka data juga salah
                    For Each row_dgv In Me.dgvExchageRates.Rows
                        If CDbl(row_dgv.Cells(Me.ColExchangeRate.Index).Value) <> 1 Then
                            i += 1
                            If Dummy = "" Then
                                Dummy = "The following document have exchange rate value <> 1:" + vbCrLf + _
                                i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            Else
                                Dummy = Dummy + i.ToString + ". " + row_dgv.Cells(Me.ColDocumentNo.Index).Value.ToString + vbCrLf
                            End If
                        End If
                    Next
                End If

                If Dummy <> "" Then
                    Dummy = Dummy + vbCrLf + "This is wrong data input."
                    MsgBox(Dummy, MsgBoxStyle.YesNo, "Confirm")
                    Return False
                End If
            End If

            'semua invoice DP harus memiliki currency yang sama dengan currency invoice ybs
            Dummy = ""
            For Each row_dgv In Me.dgvDP.Rows
                If Me.ObjInvoice.Get_Primary_Currency(row_dgv.Cells(Me.ColPrimaryInvDP.Index).Value) _
                <> Me.txtPrimaryCurrency.Text Then
                    If Dummy = "" Then
                        Dummy = row_dgv.Cells(Me.ColDPNo.Index).Value.ToString
                    Else
                        Dummy = Dummy + ", " + row_dgv.Cells(Me.ColDPNo.Index).Value.ToString
                    End If
                End If
            Next
            If Dummy <> "" Then
                MsgBox("The following downpayment invoice have different currency: " _
                + Dummy + ". We cannot enter the data.", MsgBoxStyle.Information, "Information")

                Return False
            End If

            'update payment term jika masih tidak sama
            'Me.IsiPaymentTerm()

            'nilai maksimal total sales adalah: 999,999,999,999
            If CDec(Me.TxtTotal.Text) * CDec(Me.txtTaxExchangeRate.Text) > 999999999999 Then
                MsgBox("Maximum sales value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'nilai pajak maksimal adalah 999,999,999,999
            If CDec(Me.txtTaxAmount.Text) > 999999999999 Then
                MsgBox("Maximum VAT value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If CDec(Me.txtTaxAmount1.Text) > 999999999999 Then
                MsgBox("Maximum VAT value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'If Me.chkTax.Checked = True Then
            '    If CDec(Me.txtTax.Text) <= 0 Then
            '        MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
            '        Me.TabControl1.SelectedIndex = 0
            '        Me.txtTax.Focus()
            '        Return False
            '    End If
            '    If CDec(Me.txtTax.Text) <> 10 Then
            '        If MsgBox("Tax rate <> 10%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
            '            Me.txtTax.Focus()
            '            Return False
            '        End If
            '    End If
            'End If

            If Me.chkTax.Checked = True Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtTax.Focus()
                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If

                If CDec(Me.txtTax.Text) <> CDec(Dummy) Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtTax.Focus()
                    If MsgBox("Tax rate <> " + Me.ObjCD.DigitGrouping(Dummy, 2) + "%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If

            'kalo pake pajak, maka no seri faktur pajak harus diisikan
            '2025 Core Tax Tahap 2.2
            'Me.txtNoSeriFakturPajak.Text = Me.txtNoSeriFakturPajak.Text.Trim
            'If Me.chkTax.Checked = True And Me.txtNoSeriFakturPajak.Text = "" And Me.CboFakturPajak.SelectedIndex = 1 Then
            '    MsgBox("Please specify tax invoice serial number.", MsgBoxStyle.Information, "Information")
            '    Me.txtNoSeriFakturPajak.Focus()
            '    Return False
            'End If

            If Me.chkTax.Checked = True Then
                'kalo ada VAT, maka akun VAT out harus ada
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify VAT out account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If CDbl(Me.txtDutyStamp.Text) > 0 Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify dutystamp expense account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'cek setiap baris item, jika sub total = 0, maka berikan peringatan bahwa nanti akan ada trouble ketika import e-Faktur
            KoleksiNamaData.Clear()
            Dummy = ""
            For Each row_dgv In Me.dgvItem.Rows
                If CDec(row_dgv.Cells(Me.ColAmount.Index).Value) = 0 Then
                    If KoleksiNamaData.Contains(row_dgv.Cells(Me.ColCodeNo.Index).Value) = False Then
                        KoleksiNamaData.Add(row_dgv.Cells(Me.ColCodeNo.Index).Value)
                    End If
                End If
            Next
            For Each CodeNo As String In KoleksiNamaData
                If Dummy = "" Then
                    Dummy = CodeNo
                Else
                    Dummy = Dummy + ", " + CodeNo
                End If
            Next
            If Dummy <> "" Then
                If MsgBox("The following code no has amount = 0, this will cause exception in e-Faktur import process. Continue save?" + vbCrLf + Dummy, _
                MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If

            '2025 Core Tax Tahap 2.2
            'no seri faktur pajak harus: 010, 020, 030, 040, 060, 070, 080, 090
            'If Me.chkTax.Checked = True Then
            '    If Me.txtNoSeriFakturPajak.Text <> "010" And _
            '    Me.txtNoSeriFakturPajak.Text <> "020" And _
            '    Me.txtNoSeriFakturPajak.Text <> "030" And _
            '    Me.txtNoSeriFakturPajak.Text <> "040" And _
            '    Me.txtNoSeriFakturPajak.Text <> "060" And _
            '    Me.txtNoSeriFakturPajak.Text <> "070" And _
            '    Me.txtNoSeriFakturPajak.Text <> "080" And _
            '    Me.txtNoSeriFakturPajak.Text <> "090" And _
            '     Me.txtNoSeriFakturPajak.Text <> "011" And _
            '    Me.txtNoSeriFakturPajak.Text <> "021" And _
            '    Me.txtNoSeriFakturPajak.Text <> "031" And _
            '    Me.txtNoSeriFakturPajak.Text <> "041" And _
            '    Me.txtNoSeriFakturPajak.Text <> "061" And _
            '    Me.txtNoSeriFakturPajak.Text <> "071" And _
            '    Me.txtNoSeriFakturPajak.Text <> "081" And _
            '    Me.txtNoSeriFakturPajak.Text <> "091" Then
            '        txtNoSeriFakturPajak.Text = txtNoSeriFakturPajak.Text.Trim
            '        MsgBox("The serial number of tax invoice must be chosen from one of the following alternatives: 010, 011, 020, 021, 030, 031, 040, 041, 060, 061, 070, 071, 080, 081, 091 or 090.", _
            '        MsgBoxStyle.Information, "Information")
            '        Me.txtNoSeriFakturPajak.Focus()
            '        Return False
            '    End If

            '    If Me.txtNoSeriFakturPajak.Text = "070" Or _
            '    Me.txtNoSeriFakturPajak.Text = "080" Or _
            '    Me.txtNoSeriFakturPajak.Text = "071" Or _
            '    Me.txtNoSeriFakturPajak.Text = "081" Then
            '        Dim ijk As New FrmTransaksiSalesInvoice_KeteranganTambahanEFaktur
            '        ijk.TxtTaxInvoiceSeries.Text = Me.txtNoSeriFakturPajak.Text
            '        ijk.vIndexAcuan = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
            '        ijk.ShowDialog()
            '        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = ijk.cboKeteranganTambahan.SelectedIndex
            '    End If

            '    'kalo IDR, pastikan tidak ada decimal
            '    If Me.txtPrimaryCurrency.Text = "1" Then
            '        If Me.ObjCD.GetJmlPrecision(Me.txtSubTotalAfterDP.Text) > 0 Then
            '            If MsgBox("The invoice is in " + Me.ObjCurrency.GetCurrencyCode(1) + _
            '            " currency and sub total amount after downpayment is in decimal value, this will cause failure in e-Faktur data import. Continue save?", _
            '            MsgBoxStyle.YesNo, "Confirm") _
            '            <> MsgBoxResult.Yes Then
            '                Return False
            '            End If
            '        End If
            '    End If
            'End If

            ''cek bisi ada qty default=0
            Dim ListError As String = ""
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If CDec(.Cells(Me.ColTotalQty.Index).Value) <= 0 Then
                        ListError += .Cells(Me.ColCodeNo.Index).Value & ", "
                    End If
                End With
            Next
            If ListError <> "" Then
                ListError = ListError.Substring(0, ListError.Length - 3) & "."
                MsgBox("The following product has zero default qty: " & vbNewLine & ListError & vbNewLine & "Please change qty or set ratio.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.chkTax.Checked = True Then
                Dim TaxAdd1, TaxAdd2 As String
                TaxAdd1 = ""
                TaxAdd2 = ""

                Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, TaxAdd1, TaxAdd2)
                If TaxAdd1.Trim = "" And TaxAdd2.Trim = "" Then
                    MsgBox("Please specify a valid tax address for the specified customer", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If CDec(Me.txtDutyStamp.Text) > 0 Then
                If Me.txtPrimaryCurrency.Text <> "1" Then
                    MsgBox("Duty stamp is for IDR invoices only.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Dim ObjCust As New CustomSMESS.ClsTransaksiStockInOut
            Dim rd As DataGridViewRow
            Dim TblDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
            TblDummy = Me.TblCustomInOutDummy.Copy
            'If TblDummy.Rows.Count > 0 Then
            For ijk As Integer = 0 To Me.dgvItem.RowCount - 1
                rd = Me.dgvItem.Rows(ijk)
                With rd
                    Dim PrimaryProduct As Double = .Cells(Me.ColPrimaryCodeNo.Index).Value
                    Dim PrimaryPOSO As Double = .Cells(Me.ColPrimarySO.Index).Value
                    Dim PrimaryLocation As Double = .Cells(Me.ColPrimaryLocation.Index).Value
                    Dim Batch As String = .Cells(Me.ColLotNumber.Index).Value
                    Dim QtyDef As Decimal = .Cells(Me.ColTotalQty.Index).Value
                    Dim QtyMulti As Decimal = .Cells(Me.ColTotalQtyMulti.Index).Value
                    Dim ProdCode As String = .Cells(Me.ColCodeNo.Index).Value
                    If ObjCust.CompareQtySMESS_Custom(PrimaryProduct, PrimaryPOSO, PrimaryLocation, Batch, QtyDef, QtyMulti, TblDummy) = False Then
                        If String.IsNullOrEmpty(Batch) = False Then
                            ProdCode &= " batch: " & Batch
                        End If
                        MsgBox("Inconsisten data, please cek product: " & ProdCode & "between main form and custom form.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.ClearSelection()
                        rd.Selected = True
                        Return False
                    End If
                End With
            Next
            'End If

            'Me.txtNoNPWP.Text = Me.txtNoNPWP.Text.Trim
            'If Me.ObjCustomer.ApakahFormatNPWPBenar(Me.txtNoNPWP.Text, True) = False Then
            '    Return False
            'End If
            'Comment: Perubahan NPWP 16 Digit
            If chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                '2025 Core Tax Tahap 2.1
                If ObjSetNPWP.SetNPWP_New(Me.dtpDate.Value.Date, Me.txtPrimaryCustomer.Text, Me.txtNoNPWP.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Customer, _
                Me.txtNoNIKPass.Text, frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionSales, Me.txtPrimaryAccount.Text, Me.UserCode) = False Then
                    Return False
                End If
            End If

            '2025 Core Tax Tahap 2.2, show info pajak, ketika pilihan 08, nanti diganti menjadi yang bawah, ketika full
            If Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax07 Or Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax08 Then
                If ShowInfoPajak() = False Then
                    Return False
                End If
            End If

            '2025 Coretax Tahap 2.2, cek apakah boleh 01 atau tidak
            If chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                If Me.dtpDate.Value <= ObjSetNPWP.TglAkhirTax04 And Me.cboTaxtype.SelectedIndex = 0 Then
                    MsgBox("The tax type 01 can't be used until " & ObjSetNPWP.TglAkhirTax04.AddDays(1) & ".", MsgBoxStyle.Information, _
                            "Information")
                    Return False
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            ObjCL = Nothing
            ListNoStockOut = Nothing
            ListPrimaryDO = Nothing
        End Try
    End Function

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TLayoutPanel2.Controls, True, False, True, False, True, False, True, False, True, False) '//#2732

            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit

            Me.dgvDP.AllowUserToDeleteRows = False
            Me.dgvItemSR.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.dgvExchageRates.ReadOnly = True

            'Me.llbAddress.Enabled = False
            Me.llbAddDP.Enabled = False
            Me.llbAddSR.Enabled = False
            'Me.llbEditNoFakturPajak.Enabled = False
            'Me.llbEditNoInvoice.Enabled = False
            Me.llbUpDateNPWP.Enabled = False
            'Me.lblOtherCharge.Enabled = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtReadOnly.BackColor, False)

            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False
            Me.llbImport.Enabled = True

            Me.RboAddressShipTo.Enabled = False
            Me.RboCustomerShipTo.Enabled = False

            '2025 Core Tax Tahap 2.2
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
            Me.LlbViewDetailTax.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage4.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer6.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer6.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.gbo_EFaktur.Controls, True, False, False)


            Me.txtPrimaryAccount.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtPrimaryInvoice.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.TxtPrimarySalesPerson.Text = "-1"
            Me.txtRevisiNo.Text = "0"
            Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.txtTax.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.txtTaxExchangeRate.Text = "1.00"
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur)
            Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur) 'new tax amount1
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur)
            Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.TxtDiscString.Text = "0"
            Me.txtExchangeRateToUSD.Text = "1"
            Me.txtExhangeRateToIDR.Text = "1"
            Me.txtPaymentTerm.Text = "0"
            Me.TxtQtyDefault.Text = "0"
            Me.TxtQtyMulti.Text = "0"
            Me.txtPrimaryCustomer.Text = "-1"
            Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
            Me.TxtPrimaryCurrency1.Text = Core_Module.PrimaryBaseCurrency
            Me.TxtNamaSigned.Clear()
            Me.TxtJabatanSigned.Clear()

            Me.Txt_E_Faktur_Jalan.Text = "-"
            Me.Txt_E_Faktur_Blok.Text = "-"
            Me.Txt_E_Faktur_No.Text = "-"
            Me.Txt_E_Faktur_RT.Text = "0"
            Me.Txt_E_Faktur_RW.Text = "0"
            Me.Txt_E_Faktur_Kecamatan.Text = "-"
            Me.Txt_E_Faktur_Kelurahan.Text = "-"
            Me.Txt_E_Faktur_Propinsi.Text = "-"
            Me.Txt_E_Faktur_Kabupaten.Text = "-"
            Me.Txt_E_Faktur_KodePos.Text = "00000"
            Me.Txt_E_Faktur_NoTelepon.Text = "-"
            Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = "-1"

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.LblPrintStatus.Text = "Print Status: -"

            Me.dgvItem.Rows.Clear()
            Me.dgvDP.Rows.Clear()
            Me.dgvExchageRates.Rows.Clear()
            Me.dgvItemSR.Rows.Clear()
            Me.dgvSR.Rows.Clear()

            Me.chkTax.Checked = True
            Me.cboCustomerAddress.DataSource = Nothing
            Me.cboCustomerAddress.Text = ""
            Me.cboCustomerAddress.Items.Clear()
            Me.chkAlias.Checked = False

            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.cboCurrTax.DataSource = Nothing

            Me.TBLHeaderCost.Rows.Clear()
            Me.TBLitemCost.Rows.Clear()
            Me.cboPrimaryAccount.Text = ""
            Me.cboAccount.Text = ""
            Me.cboCurrTax.Text = ""

            Me.cboPrimaryAccount.Items.Clear()
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.cboAccount.Items.Clear()
            Me.cboCurrTax.Items.Clear()

            Me.CboAddressShipTo.Items.Clear()
            Me.CboAddressCustomerShipTo.Items.Clear()
            Me.TxtPrimaryCustomerShipTo.Text = "-1"

            Me.PrimaryCustomerShipTo = -1
            Me.IdentifierAddressCustomerShipTo = ""

            Me.TblCustomDummy.Rows.Clear()
            Me.TblCustomInOutDummy.Rows.Clear()

            '2025 Core Tax Tahap 2.2
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
            Me.txtNoFakturPajak.Text = ""

            '2025 Core Tax Tahap 2.2 
            Me.AddInfo = -1
            Me.FacilityStamp = -1
            Me.SPPBDate = #1/1/1950#
            'Me.TxtSPPBNo.Text = ""
            Me.txtDPPNilaiLain.Text = "0"
            Me.txtRevisedInvoice.Text = "00" '//#2732
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer6.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TLayoutPanel2.Controls, True, True, True, True, True, True, True, True, True, True) '//#2732
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvDP.AllowUserToDeleteRows = True
            Me.dgvItemSR.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            Me.dgvExchageRates.ReadOnly = False

            'Me.llbAddress.Enabled = True
            Me.llbAddDP.Enabled = True
            Me.llbAddSR.Enabled = True
            'Me.llbEditNoFakturPajak.Enabled = True
            'Me.llbEditNoInvoice.Enabled = True
            Me.llbUpDateNPWP.Enabled = True
            'Me.lblOtherCharge.Enabled = True

            If Me.dgvItem.RowCount > 0 Then
                'Me.txtCustomerCode.ReadOnly = True
                'Me.btnFindCustomer.Enabled = False
                ''    Me.txtCustomerCode.BackColor = Me.TxtTotal.BackColor
            Else
                'Me.txtCustomerCode.ReadOnly = False
                'Me.btnFindCustomer.Enabled = True
            End If

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.txtPrimaryInvoice.BackColor, False)

            Me.DtpExpired.Enabled = False

            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                Me.txtInvoiceNo.ReadOnly = True
                Me.txtInvoiceNo.BackColor = Me.TxtReadOnly.BackColor
                Me.txtInvoiceNo.TabStop = False
                Me.dtpDate.Focus()
                Me.llbEditNoInvoice.Visible = True
            Else
                Me.txtInvoiceNo.ReadOnly = False
                Me.txtInvoiceNo.BackColor = Me.txtPrimaryInvoice.BackColor
                Me.txtInvoiceNo.TabStop = True
                Me.txtInvoiceNo.Focus()
                Me.llbEditNoInvoice.Visible = False
            End If

            'set default precision untuk idr
            Me.DefaultPrecision_DefCur = Me.ObjCurrency.GetDecimal_DefaultCurrency
            Me.SetTaxRate(False)

            'set llb faktur pajak
            If Me.chkTax.Checked = True Then
                Me.llbEditNoFakturPajak.Enabled = True
            Else
                Me.llbEditNoFakturPajak.Enabled = False
            End If
            Me.llbImport.Enabled = False


            Me.cboAccount.Enabled = True
            Me.cboCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
            End If

            '//TR : Tambahan baru Nik/Pas
            If Me.cboNIKPASS.Visible = True Then
                Me.cboNIKPASS.TabStop = True
            Else
                Me.cboNIKPASS.TabStop = False
            End If

            Me.RboAddressShipTo.Enabled = True
            Me.RboCustomerShipTo.Enabled = True
            If Me.RboAddressShipTo.Checked Then
                CboAddressShipTo.Enabled = True
                LlbViewAddressShipTo.Enabled = True
                TxtCustomerCodeShipTo.Enabled = False
                BtnFindCustomerShipTo.Enabled = False
                CboAddressCustomerShipTo.Enabled = False
                LlbViewAddressCustomerShipTo.Enabled = False
            Else
                CboAddressShipTo.Enabled = False
                LlbViewAddressShipTo.Enabled = False
                TxtCustomerCodeShipTo.Enabled = True
                BtnFindCustomerShipTo.Enabled = True
                CboAddressCustomerShipTo.Enabled = True
                LlbViewAddressCustomerShipTo.Enabled = True
            End If

            '1
            Me.LinkLabel1.Enabled = True
            Atur1()

            '2025 Core Tax Tahap 2.2
            Me.llbEditNoFakturPajak.Visible = False
            Me.txtNoSeriFakturPajak.Visible = False
            Me.txtNoFakturPajak.Visible = True
            Me.LlbViewDetailTax.Enabled = False
            If Me.baru Then
                Me.chkRevisedInvoice.Enabled = False
            Else
                Me.chkRevisedInvoice.Enabled = True
            End If
            'Me.lblSPPBNo.Visible = False
            'Me.TxtSPPBNo.Visible = False
            Me.txtDPPNilaiLain.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim TaxPercentSetting, UseTaxSetting As String

        Try
            If Me.CheckLicense = True Then
                If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                    MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.baru = True
            Me.SedangRevisi = False
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TabControl1.SelectedIndex = 0
            Me.dtpDate.Focus()
            Me.ObjInvoice.ResetObject(Me.dgvItem, Me.dgvDP, Me.dgvExchageRates, Me.dgvItemSR)
            Me.txtInvoiceNo.Text = Me.ObjInvoice.GetNoBaru1(Me.dtpDate.Value)
            Me.CboFakturPajak.SelectedIndex = 1
            '2025 Core Tax Tahap 2.2
            'Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
            Me.ObjInvoice.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice

            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            Me.SetDefaultPrecision()
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiCBOCurrencyTax()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
            Else
                If UseTaxSetting = "0" Then
                    Me.chkTax.Checked = False
                    Me.CboFakturPajak.SelectedIndex = 0
                Else
                    Me.chkTax.Checked = True
                    TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(TaxPercentSetting) = True Then
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping(TaxPercentSetting, 2)
                    End If
                End If
            End If

            Me.TxtNamaSigned.Text = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.SignedByName)
            Me.TxtJabatanSigned.Text = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.SignedByJabatan)

            Me.SetTaxRate()
            Me.SetCalculateBy()

            If Me.txtInvoiceNo.Text = "" And Me.txtInvoiceNo.Enabled = True Then
                Me.txtInvoiceNo.Focus()
            End If

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.ShipmentTax)
            If UseTaxSetting = "" Or UseTaxSetting = "0" Then
                Me.cboShipmentTax.SelectedIndex = 1
            Else
                Me.cboShipmentTax.SelectedIndex = 0
            End If
            Me.cboPPNDipungut.SelectedIndex = 0
            Me.Hitung_Total()

            '2025 Core Tax Tahap 2.2
            Me.cboTaxtype.SelectedIndex = Me.ObjInvoice.getLastTaxType()
            Me.txtNoFakturPajak.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Public Sub Hitung_Total()
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Double = 0
        Dim TBL1 As New DseTransaksiDO.DataTable10DataTable
        Dim DTA1 As New DseTransaksiDOTableAdapters.DataTable10TableAdapter
        Dim DummyPrimaryDO As Double = -1
        Dim PrimaryDO As Double = -1
        Dim PrimarySO, PrimaryProduct As Double
        Dim PriceExclude As Boolean = True
        Dim Unit_Price_SO As Double = 0
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ED As Date = #1/1/1950#
        Dim DefPrec As Integer = 0
        Dim DiscAmount As Decimal = 0
        Dim BaseNumber As Decimal = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Total_MiddlePoint()
                Exit Sub
            End If
            If Me.txtPrimaryCustomer.Text = "" Then Me.txtPrimaryCustomer.Text = "-1"
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            If IsNumeric(Me.txtTaxAmount.Text) = False Then
                Me.txtTaxAmount.Text = 0
            End If
            If IsNumeric(Me.TxtShipment.Text) = False Then
                Me.TxtShipment.Text = 0
            End If
            If IsNumeric(Me.txtDutyStamp.Text) = False Then
                Me.txtDutyStamp.Text = 0
            End If
            If IsNumeric(txtTaxAmount.Text) = False Then
                txtTaxAmount.Text = 0
            End If
            If IsNumeric(Me.txtTaxAmount1.Text) = False Then 'new tax amount1
                Me.txtTaxAmount1.Text = 0
            End If

            If Me.dgvItem.Rows.Count > 0 Then
                'hitung sub total
                For Each row In Me.dgvItem.Rows
                    PrimaryDO = row.Cells(Me.ColPrimaryDO.Index).Value
                    PrimarySO = row.Cells(Me.ColPrimarySO.Index).Value
                    PrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                    If DummyPrimaryDO <> PrimaryDO Then
                        DTA1.Fill_ByPrimaryDO(TBL1, PrimaryDO)
                        DummyPrimaryDO = PrimaryDO
                    End If
                    PriceExclude = Me.Get_Price_Exclude(PrimarySO, PrimaryProduct, TBL1)
                    Unit_Price_SO = row.Cells(Me.ColUnitPrice.Index).Value

                    'yang diperhitungkan hanya yang tidak pake secondary unit
                    If Me.cboCalculatesBy.SelectedIndex = 0 Then
                        'berdasarkan sales invoice
                        exchangRate = Me.dgvExchageRates.Rows(0).Cells(3).Value
                    ElseIf Me.cboCalculatesBy.SelectedIndex = 1 Then
                        'berdasarkan delivery order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryDO.Index).Value)
                    ElseIf Me.cboCalculatesBy.SelectedIndex = 2 Then
                        'berdasarkan saler order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimarySO.Index).Value)
                    End If

                    'hitung sales return jika ada
                    Dim DefaultPrecisionRow As Double = 1
                    Dim SubTotal1 As Double = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    Dim SecQty As Double
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                        SecQty = row.Cells(Me.ColSecQty.Index).Value
                    Else
                        SecQty = -1
                    End If

                    If IsDate(row.Cells(Me.ColExpiry.Index).Value) = True Then
                        ED = row.Cells(Me.ColExpiry.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    'Dim JmlReturn As Double = Me.GetJmlDGVSR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'row.Cells(Me.ColPrimaryDO.Index).Value, _
                    'SecQty, row.Cells(Me.ColPrimarySO.Index).Value, _
                    'row.Cells(Me.ColLotNumber.Index).Value, ED) -> untuk quick inv tidak ada return

                    'If Me.MetodeHitungInclude = EnumHitungInclude.UnitPriceInclude Then
                    '    If PriceExclude = True Then
                    '        If SecQty = -1 Then
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            row.Cells(Me.ColNettPrice.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        Else
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        End If
                    '    Else
                    '        If SecQty = -1 Then
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            exchangRate * (Unit_Price_SO * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        Else
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            row.Cells(Me.ColSecQty.Index).Value * exchangRate * (Unit_Price_SO * _
                    '            ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        End If
                    '    End If
                    'ElseIf Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                    If PriceExclude = True Then
                        If SecQty = -1 Then
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, CInt(Me.DefaultPrecision), ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, CInt(Me.DefaultPrecision), ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        Else
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    Else
                        If SecQty = -1 Then
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        Else
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                                row.Cells(Me.ColSecQty.Index).Value * exchangRate * row.Cells(Me.ColNettPrice.Index).Value, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * _
                                row.Cells(Me.ColSecQty.Index).Value * exchangRate * row.Cells(Me.ColNettPrice.Index).Value, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    End If
                    'End If
                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
                Next

                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
                Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(Me.TxtShipment.Text, DefaultPrecision)
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.dgvDP.Rows
                    Me.txtDownPayment.Text = CDbl(Me.txtDownPayment.Text) + CDbl(row.Cells(3).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtDownPayment.Text, DefaultPrecision))


                '------------------------------------------------------------------------------------------------------------------------'
                '------------------------------------------------------------------------------------------------------------------------'
                'diganti berdasarkan complain error ci fitri no 50 - error sales smess
                'Keputusan : with or without tax, shipment tetap dihitung, tapi saat perhitungan pajak, shipment tidak dihitung
                '2016/07/25
                '------------------------------------------------------------------------------------------------------------------------'
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                Dim DecDasarPerhitunganPajak As Decimal = 0
                If Me.cboShipmentTax.SelectedIndex = 0 Then
                    DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                Else
                    DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                End If


                'If Me.cboShipmentTax.SelectedIndex = 0 Then
                '    'shipment with tax
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDbl(Me.txtDownPayment.Text), DefaultPrecision))
                'Else
                '    'shipment without tax
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDbl(Me.txtDownPayment.Text), DefaultPrecision))
                'End If



                'hitung global disc
                BaseNumber = 1 / (10 ^ DefaultPrecision)
                Me.TxtDiscAmount.Text = Me.ObjCD.Get_Disc_Amount(Me.TxtDiscString.Text, CDec(Me.txtSubTotalAfterDP.Text), _
                DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint, BaseNumber, Me.ObjCD)

                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                    DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Else
                    DefPrecTax = DefaultPrecision
                End If

                If Me.chkTax.Checked = True Then
                    'If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    'ganti dengan cara perhitungan yang lebih stabil
                    'by: iwan - 2017-01-16
                    'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                    'If Me.txtPrimaryCurrency.Text = "1" Then
                    '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    '   DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    'Else
                    '    If Me.txtPrimaryCurrTax.Text = "1" Then
                    '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                    '                              DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    '    Else

                    '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    '                               DecDasarPerhitunganPajak / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    '    End If
                    'End If
                    DecDasarPerhitunganPajak = 0
                    For Each row In Me.dgvItem.Rows
                        DecDasarPerhitunganPajak += CDec(row.Cells(Me.ColVAT2.Index).Value)
                    Next
                    If Me.cboShipmentTax.SelectedIndex = 0 Then
                        DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(DecDasarPerhitunganPajak + _
                        CDec(Me.ObjCD.DigitGrouping(((CDec(Me.TxtShipment.Text) * CDec(Me.txtTax.Text)) / 100), 0, ClsCekData.EnumTipeRounding.Down)), DefaultPrecision))
                    End If
                    If Me.txtPrimaryCurrency.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(DecDasarPerhitunganPajak, DefPrecTax)
                    Else
                        If Me.txtPrimaryCurrTax.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * _
                            DecDasarPerhitunganPajak, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                        Else

                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(DecDasarPerhitunganPajak, DefPrecTax)
                        End If
                    End If
                    'End If

                    Me.txtTaxAmount1.Text = Me.txtTaxAmount.Text 'new tax amount1
                    If Me.cboPPNDipungut.SelectedIndex = 1 Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                    Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax) 'new tax amount1
                End If
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DefPrecTax)
            Else
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.dgvDP.Rows.Clear()
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text)) 'new tax amount1
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            End If

            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.txtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                    decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.txtDutyStamp.Text) Then
                    decTotal += CDec(Me.txtDutyStamp.Text)
                End If
                If IsNumeric(Me.txtOtherCharge.Text) Then
                    decTotal += CDec(Me.txtOtherCharge.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
            Else
                If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.txtDutyStamp.Text) Then
                        decTotal += CDec(Me.txtDutyStamp.Text)
                    End If
                    If IsNumeric(Me.txtOtherCharge.Text) Then
                        decTotal += CDec(Me.txtOtherCharge.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtOtherCharge.Text) Then
                        decTotal += CDec(Me.txtOtherCharge.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Public Sub Hitung_Total_MiddlePoint()
        Dim row As DataGridViewRow
        Dim SubTotal As Double = 0
        Dim exchangRate As Double = 0
        Dim TBL1 As New DseTransaksiDO.DataTable10DataTable
        Dim DTA1 As New DseTransaksiDOTableAdapters.DataTable10TableAdapter
        Dim DummyPrimaryDO As Double = -1
        Dim PrimaryDO As Double = -1
        Dim PrimarySO, PrimaryProduct As Double
        Dim PriceExclude As Boolean = True
        Dim Unit_Price_SO As Double = 0
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ED As Date = #1/1/1950#
        Dim DefPrec As Integer = 0
        Dim DiscAmount As Decimal = 0
        Dim BaseNumber As Decimal = 0
        Dim DecDasarPerhitunganPajak As Decimal = 0
        Dim DecDasarPerhitunganPajak2 As Decimal = 0
        Try
            If Me.txtPrimaryCustomer.Text = "" Then Me.txtPrimaryCustomer.Text = "-1"
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            If IsNumeric(Me.txtTaxAmount.Text) = False Then
                Me.txtTaxAmount.Text = 0
            End If
            If IsNumeric(Me.TxtShipment.Text) = False Then
                Me.TxtShipment.Text = 0
            End If
            If IsNumeric(Me.txtDutyStamp.Text) = False Then
                Me.txtDutyStamp.Text = 0
            End If
            If IsNumeric(txtTaxAmount.Text) = False Then
                txtTaxAmount.Text = 0
            End If
            If IsNumeric(Me.txtTaxAmount1.Text) = False Then 'new tax amount1
                Me.txtTaxAmount1.Text = 0
            End If

            If Me.dgvItem.Rows.Count > 0 Then
                'hitung sub total
                For Each row In Me.dgvItem.Rows
                    PrimaryDO = row.Cells(Me.ColPrimaryDO.Index).Value
                    PrimarySO = row.Cells(Me.ColPrimarySO.Index).Value
                    PrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                    If DummyPrimaryDO <> PrimaryDO Then
                        DTA1.Fill_ByPrimaryDO(TBL1, PrimaryDO)
                        DummyPrimaryDO = PrimaryDO
                    End If
                    PriceExclude = Me.Get_Price_Exclude(PrimarySO, PrimaryProduct, TBL1)
                    Unit_Price_SO = row.Cells(Me.ColUnitPrice.Index).Value

                    'yang diperhitungkan hanya yang tidak pake secondary unit
                    If Me.cboCalculatesBy.SelectedIndex = 0 Then
                        'berdasarkan sales invoice
                        exchangRate = Me.dgvExchageRates.Rows(0).Cells(3).Value
                    ElseIf Me.cboCalculatesBy.SelectedIndex = 1 Then
                        'berdasarkan delivery order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimaryDO.Index).Value)
                    ElseIf Me.cboCalculatesBy.SelectedIndex = 2 Then
                        'berdasarkan saler order no
                        exchangRate = Me.GetExchangeRate(row.Cells(Me.ColPrimarySO.Index).Value)
                    End If

                    'hitung sales return jika ada
                    Dim DefaultPrecisionRow As Double = 1
                    Dim SubTotal1 As Double = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    Dim SecQty As Double
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                        SecQty = row.Cells(Me.ColSecQty.Index).Value
                    Else
                        SecQty = -1
                    End If

                    If IsDate(row.Cells(Me.ColExpiry.Index).Value) = True Then
                        ED = row.Cells(Me.ColExpiry.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    'Dim JmlReturn As Double = Me.GetJmlDGVSR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    'row.Cells(Me.ColPrimaryDO.Index).Value, _
                    'SecQty, row.Cells(Me.ColPrimarySO.Index).Value, _
                    'row.Cells(Me.ColLotNumber.Index).Value, ED) -> untuk quick inv tidak ada return

                    'If Me.MetodeHitungInclude = EnumHitungInclude.UnitPriceInclude Then
                    '    If PriceExclude = True Then
                    '        If SecQty = -1 Then
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            row.Cells(Me.ColNettPrice.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        Else
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        End If
                    '    Else
                    '        If SecQty = -1 Then
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            exchangRate * (Unit_Price_SO * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        Else
                    '            SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                    '            row.Cells(Me.ColSecQty.Index).Value * exchangRate * (Unit_Price_SO * _
                    '            ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '        End If
                    '    End If
                    'ElseIf Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                    If PriceExclude = True Then
                        If SecQty = -1 Then
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, CInt(Me.DefaultPrecision), ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, CInt(Me.DefaultPrecision), ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        Else
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    Else
                        If SecQty = -1 Then
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * row.Cells(Me.ColNettPrice.Index).Value * _
                                exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        Else
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                                row.Cells(Me.ColSecQty.Index).Value * exchangRate * row.Cells(Me.ColNettPrice.Index).Value, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                SubTotal1 = ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value) * _
                                row.Cells(Me.ColSecQty.Index).Value * exchangRate * row.Cells(Me.ColNettPrice.Index).Value, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    End If
                    'End If
                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
                Next

                If Me.txtPrimaryCurrency.Text = "1" Then
                    Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(SubTotal, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefaultPrecision)
                Else
                    Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
                End If

                'Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
                Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(Me.TxtShipment.Text, DefaultPrecision)
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.dgvDP.Rows
                    Me.txtDownPayment.Text = CDbl(Me.txtDownPayment.Text) + CDbl(row.Cells(3).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtDownPayment.Text, DefaultPrecision))


                '------------------------------------------------------------------------------------------------------------------------'
                '------------------------------------------------------------------------------------------------------------------------'
                'diganti berdasarkan complain error ci fitri no 50 - error sales smess
                'Keputusan : with or without tax, shipment tetap dihitung, tapi saat perhitungan pajak, shipment tidak dihitung
                '2016/07/25
                '------------------------------------------------------------------------------------------------------------------------'
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(Me.txtSubTotalBeforeDP.Text) + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                'Dim DecDasarPerhitunganPajak As Decimal = 0
                If Me.cboShipmentTax.SelectedIndex = 0 Then
                    DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                Else
                    DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                End If
                DecDasarPerhitunganPajak2 = DecDasarPerhitunganPajak

                'If Me.cboShipmentTax.SelectedIndex = 0 Then
                '    'shipment with tax
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDbl(Me.txtDownPayment.Text), DefaultPrecision))
                'Else
                '    'shipment without tax
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDbl(Me.txtDownPayment.Text), DefaultPrecision))
                'End If



                'hitung global disc
                BaseNumber = 1 / (10 ^ DefaultPrecision)
                Me.TxtDiscAmount.Text = Me.ObjCD.Get_Disc_Amount(Me.TxtDiscString.Text, CDec(Me.txtSubTotalAfterDP.Text), _
                DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint, BaseNumber, Me.ObjCD)

                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                    DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Else
                    DefPrecTax = DefaultPrecision
                End If

                If Me.chkTax.Checked = True Then
                    'If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    'ganti dengan cara perhitungan yang lebih stabil
                    'by: iwan - 2017-01-16
                    'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                    'If Me.txtPrimaryCurrency.Text = "1" Then
                    '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    '   DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    'Else
                    '    If Me.txtPrimaryCurrTax.Text = "1" Then
                    '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                    '                              DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    '    Else

                    '        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                    '                               DecDasarPerhitunganPajak / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    '    End If
                    'End If
                    DecDasarPerhitunganPajak = 0
                    For Each row In Me.dgvItem.Rows
                        DecDasarPerhitunganPajak += CDec(row.Cells(Me.ColVAT2.Index).Value)
                    Next
                    If Me.cboShipmentTax.SelectedIndex = 0 Then
                        DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(DecDasarPerhitunganPajak + _
                        CDec(Me.ObjCD.DigitGrouping(((CDec(Me.TxtShipment.Text) * CDec(Me.txtTax.Text)) / 100), 0, ClsCekData.EnumTipeRounding.Middlepoint)), DefaultPrecision))
                    End If

                    If Me.txtPrimaryCurrency.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(DecDasarPerhitunganPajak, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                    Else
                        If Me.txtPrimaryCurrTax.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * _
                            DecDasarPerhitunganPajak, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                        Else

                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(DecDasarPerhitunganPajak, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                        End If
                    End If
                    'End If

                    Me.txtTaxAmount1.Text = Me.txtTaxAmount.Text 'new tax amount1
                    If Me.cboPPNDipungut.SelectedIndex = 1 Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                    Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax) 'new tax amount1
                End If
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DefPrecTax)
            Else
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.dgvDP.Rows.Clear()
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text)) 'new tax amount1
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            End If

            '2025 Core Tax Tahap 2.2
            Me.UbahDPPNilaiLain(DecDasarPerhitunganPajak2)

            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If Me.txtPrimaryCurrency.Text = "1" Then
                'kalo IDR, duty stamp diperhitungkan
                If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                    decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                End If
                If IsNumeric(Me.txtTaxAmount.Text) Then
                    decTotal += CDec(Me.txtTaxAmount.Text)
                End If
                If IsNumeric(Me.txtDutyStamp.Text) Then
                    decTotal += CDec(Me.txtDutyStamp.Text)
                End If
                If IsNumeric(Me.txtOtherCharge.Text) Then
                    decTotal += CDec(Me.txtOtherCharge.Text)
                End If
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefPrec))
            Else
                If Me.txtPrimaryCurrency.Text = Me.txtPrimaryCurrTax.Text Then
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        decTotal += CDec(Me.txtTaxAmount.Text)
                    End If
                    If IsNumeric(Me.txtDutyStamp.Text) Then
                        decTotal += CDec(Me.txtDutyStamp.Text)
                    End If
                    If IsNumeric(Me.txtOtherCharge.Text) Then
                        decTotal += CDec(Me.txtOtherCharge.Text)
                    End If
                    'kalo Currencynya sama diperhitungkan duty stamp dan tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                Else
                    'kalo IDR, duty stamp diperhitungkan
                    If IsNumeric(Me.txtSubTotalAfterDP.Text) Then
                        decTotal += CDec(Me.txtSubTotalAfterDP.Text)
                    End If
                    If IsNumeric(Me.txtOtherCharge.Text) Then
                        decTotal += CDec(Me.txtOtherCharge.Text)
                    End If
                    'kalo bukan IDR duty stamp dan pajak tidak diperhitungkan
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(decTotal, DefaultPrecision))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub


    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision()
                    Me.cboCalculatesBy.Focus()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision()
                        Me.cboCalculatesBy.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If

                Me.IsiCBOCurrencyTax()
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.SetTaxRate()
                Me.SetCalculateBy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtDutyStamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.Leave
        Dim defPrec As Decimal = 0
        Try
            If Me.TxtRemarks.Enabled = True Then
                defPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, defPrec)
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim i As New FrmCariMasterCurrency

        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision()
                        Me.cboCalculatesBy.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
            End If

            Me.IsiCBOCurrencyTax()
            Me.IsiCBOArAkunByTipePerkiraan1()

            Me.SetTaxRate()
            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                Me.cboCalculatesBy.Focus()
                Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                Me.SetDefaultPrecision()
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.txtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision()
                    Me.cboCalculatesBy.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetDefaultPrecision()
                End If
            End If

            Me.IsiCBOCurrencyTax()
            Me.IsiCBOArAkunByTipePerkiraan1()

            Me.SetTaxRate()
            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub


    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If Me.TxtRemarks.Enabled = True AndAlso _
            IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso _
            (Me.txtPrimaryCurrency.Text) > 0 Then
                Me.IsiCBOCurrencyTax()
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.txtOtherCharge.Text = 0
                Me.TBLHeaderCost.Rows.Clear()
                Me.TBLitemCost.Rows.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub cboCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            'Me.txtPrimaryAccount.Visible = True
            If Me.TxtRemarks.Enabled = True AndAlso IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.txtDutyStamp.Text) Then
                    Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DeffPrec)
                End If

                If IsNumeric(Me.txtTaxAmount.Text) Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                End If

                Me.lblCurrTax.Text = Me.cboCurrTax.Text
                Me.Hitung_Total()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub cboPrimaryCurrTax_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryCurrTax.SelectedIndexChanged
        Dim DeffPrec As Decimal = 0

        Try
            If Me.TxtRemarks.Enabled = True Then
                'Me.txtPrimaryAccount.Visible = True
                If IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                    Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                    DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                    If IsNumeric(Me.txtDutyStamp.Text) Then
                        Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DeffPrec)
                    End If

                    If IsNumeric(Me.txtTaxAmount.Text) Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTaxAmount.Text, DeffPrec)
                    End If

                    Me.lblCurrTax.Text = Me.cboCurrTax.Text
                    Me.Hitung_Total()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryCurrTax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrTax.TextChanged
        Dim PrimaryCurrTax As Double = -1

        Try
            If Me.Visible = False Then Exit Sub
            If IsNumeric(Me.txtPrimaryCurrTax.Text) Then
                PrimaryCurrTax = CDbl(Me.txtPrimaryCurrTax.Text)
            End If

            If Me.tblCurrTax.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblCurrTax.Rows.Count - 1
                    Me.rowCurrTax = Me.tblCurrTax.Rows(IntA)
                    With Me.rowCurrTax
                        If .Primary = PrimaryCurrTax Then
                            Me.cboPrimaryCurrTax.SelectedIndex = IntA
                            Me.cboCurrTax.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAccount.SelectedIndexChanged
        Try
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPrimaryAccount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryAccount.SelectedIndexChanged
        Try
            ' Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryAccount.Text) AndAlso Me.cboPrimaryAccount.Text > -1 Then
                Me.txtPrimaryAccount.Text = Me.cboPrimaryAccount.Text
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPrimaryAccount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryAccount.TextChanged
        Dim PrimaryARAkun As Double = -1

        Try
            If Me.Visible = False Then Exit Sub
            If IsNumeric(Me.txtPrimaryAccount.Text) Then
                PrimaryARAkun = CDbl(Me.txtPrimaryAccount.Text)
            End If

            If Me.tblAccount.Rows.Count > 0 Then
                For IntA As Integer = 0 To Me.tblAccount.Rows.Count - 1
                    Me.rowAccount = Me.tblAccount.Rows(IntA)
                    With Me.rowAccount
                        If .Primary = PrimaryARAkun Then
                            Me.cboPrimaryAccount.SelectedIndex = IntA
                            Me.cboAccount.SelectedIndex = IntA
                        End If
                    End With
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================
    Public Enum EnumHitungInclude As Integer
        UnitPriceSO
        UnitPriceInclude
        NettPriceInclude
    End Enum

    Private Sub FrmTransaksiAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FrmTransaksiSalesInvoiceQuick.CheckForIllegalCrossThreadCalls = False

        Try
            '~ab
            '3/31/2020
            If Me.vFromQuick = True Then
                Me.btnFind.Visible = False
                Me.btnFind.TabStop = False
                Me.btnNew.Visible = False
                Me.btnNew.TabStop = False
                Me.btnEdit.Visible = False
                Me.btnEdit.TabStop = False
                Me.btnDelete.Visible = False
                Me.btnDelete.TabStop = False
                Me.btnExit.Visible = False
                Me.btnExit.TabStop = False
                Me.btnPrint.Visible = False
                Me.btnPrint.TabStop = False
                Me.SimpanLangsung()
            End If

            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CboFakturPajak.SelectedIndex = 0
            Me.cboCalculatesBy.SelectedIndex = 0
            Me.dgvDP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvExchageRates.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemSR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Text = "Quick Sales"
            Me.SplitContainer4.Panel2Collapsed = True
            Control.CheckForIllegalCrossThreadCalls = False
            Me.DefaultPrecision_DefCur = Me.ObjCurrency.GetDecimal_DefaultCurrency
            Me.cboShipmentTax.SelectedIndex = 0
            Me.cboPPNDipungut.SelectedIndex = 0

            '-balikin lagi
            If Me.objCustom.ApakahBolehImportInvoice = False Then
                Me.llbImport.Visible = False
                Me.ProgressBar1.Visible = False
            Else
                Me.llbImport.Visible = True
                Me.ProgressBar1.Visible = True
            End If

            'If Me.UserCode <> "USER" Then
            '    Me.LblPrintStatus.Visible = False
            'End If
            Me.dtpDate.Left = Me.SplitContainer6.Panel1.Width - 5 - Me.dtpDate.Width
            Me.llbEditNoInvoice.Left = Me.dtpDate.Left - 5 - Me.llbEditNoInvoice.Width
            Me.txtInvoiceNo.Width = Me.llbEditNoInvoice.Left - 5 - Me.txtInvoiceNo.Left
            Me.TxtCustomerName.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.TxtCustomerName.Width
            Me.btnFindCustomer.Left = Me.TxtCustomerName.Left - 5 - Me.btnFindCustomer.Width
            Me.txtCustomerCode.Width = Me.btnFindCustomer.Left - 5 - Me.txtCustomerCode.Left
            Me.btnFindSalesPerson.Left = Me.btnFindCustomer.Left
            Me.TxtSalesPersonName.Left = Me.TxtCustomerName.Left
            Me.TxtSalesPersonName.Width = Me.TxtCustomerName.Width
            Me.TxtSalesPersonCode.Width = Me.btnFindSalesPerson.Left - 5 - Me.TxtSalesPersonCode.Left

            Me.llbUpDateNPWP.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.llbUpDateNPWP.Width
            Me.txtNoNPWP.Width = Me.llbUpDateNPWP.Left - 5 - Me.txtNoNPWP.Left
            Me.TLayoutPanel2.Width = Me.llbUpDateNPWP.Left - 3 - Me.TLayoutPanel2.Left '//#2732

            Me.txtNoNIKPass.Width = Me.txtNoNPWP.Width
            Me.llbAddress.Left = Me.llbUpDateNPWP.Left
            Me.cboCustomerAddress.Width = Me.txtNoNPWP.Width
            Me.llbEditNoFakturPajak.Left = Me.llbAddress.Left
            Me.LlbViewDetailTax.Left = Me.llbAddress.Left
            Me.txtNoFakturPajak.Left = Me.llbEditNoFakturPajak.Left - 5 - Me.txtNoFakturPajak.Width
            Me.Label8.Left = Me.txtNoFakturPajak.Left - 5 - Me.Label8.Width
            Me.txtNoSeriFakturPajak.Width = Me.Label8.Left - 5 - Me.txtNoSeriFakturPajak.Left
            Me.txtNoFakturPajakSederhana.Width = Me.dtpDate.Left + Me.dtpDate.Width - Me.txtNoFakturPajakSederhana.Left
            Me.SplitContainer3.Width = Me.txtNoFakturPajakSederhana.Width
            'Me.btnFindCustomer.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.btnFindCustomer.Width
            Me.TxtNamaSigned.Top = 0
            Me.TxtJabatanSigned.Top = 0
            Me.TxtNamaSigned.Width = Me.SplitContainer3.Panel1.Width - Me.TxtNamaSigned.Left
            Me.TxtJabatanSigned.Width = Me.SplitContainer3.Panel2.Width - Me.TxtJabatanSigned.Left
            Me.cboCalculatesBy.Width = Me.txtNoFakturPajakSederhana.Width
            Me.cboAccount.Width = Me.txtNoFakturPajakSederhana.Width
            Me.cboPrimaryAccount.Left = Me.cboAccount.Left
            Me.cboPrimaryAccount.Visible = False
            Me.btnFindCurrency.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.btnFindCurrency.Width
            Me.txtCurrencyCode.Width = Me.btnFindCurrency.Left - 5 - Me.txtCurrencyCode.Left

            Me.txtSubTotalAfterDP.Width = Me.SplitContainer5.Panel2.Width - 5 - Me.txtSubTotalAfterDP.Left
            Me.txtSubTotalBeforeDP.Width = Me.txtSubTotalAfterDP.Width
            Me.TxtShipment.Left = Me.txtSubTotalAfterDP.Left + Me.txtSubTotalAfterDP.Width - Me.TxtShipment.Width
            Me.cboShipmentTax.Width = Me.TxtShipment.Left - 5 - Me.cboShipmentTax.Left
            Me.txtDownPayment.Width = Me.txtSubTotalAfterDP.Width
            Me.txtTaxExchangeRate.Left = Me.txtSubTotalAfterDP.Width + Me.txtSubTotalAfterDP.Left - Me.txtTaxExchangeRate.Width
            Me.Label17.Left = Me.txtTaxExchangeRate.Left - 5 - Me.Label17.Width
            Me.Label25.Left = Me.Label17.Left - 10 - Me.Label25.Width
            Me.txtTax.Width = Me.Label25.Left - 5 - Me.txtTax.Left
            Me.txtTaxAmount.Left = Me.txtSubTotalAfterDP.Width + Me.txtSubTotalAfterDP.Left - Me.txtTaxAmount.Width
            Me.cboPPNDipungut.Width = Me.txtTaxAmount.Left - 5 - Me.cboPPNDipungut.Left
            Me.txtOtherCharge.Width = Me.txtSubTotalAfterDP.Width
            Me.TextBox2.Width = Me.txtSubTotalAfterDP.Width
            Me.TextBox2.Top = Me.txtOtherCharge.Top
            Me.txtDutyStamp.Width = Me.txtSubTotalAfterDP.Width
            Me.TxtTotal.Width = Me.txtSubTotalAfterDP.Width
            Me.TxtRemarks.Width = Me.SplitContainer5.Panel1.Width - Me.TxtRemarks.Left

            Me.BtnFindCodeNo.Left = Me.SplitContainer5.Panel1.Width - Me.BtnFindCodeNo.Width
            Me.BtnFindPONo.Left = Me.BtnFindCodeNo.Left
            Me.BtnFindPONo.Width = Me.BtnFindCodeNo.Width
            Me.TxtCodeNo.Width = Me.BtnFindCodeNo.Left - 5 - Me.TxtCodeNo.Left
            Me.txtDesc.Width = Me.SplitContainer5.Panel1.Width - Me.txtDesc.Left
            Me.cboBatchNo.Width = Me.txtDesc.Width
            Me.DtpExpired.Width = Me.cboBatchNo.Width
            Me.cboLocation.Width = Me.txtDesc.Width

            Me.txtAmount.Width = Me.SplitContainer5.Panel2.Width - Me.txtAmount.Left
            Me.btnInput.Left = Me.txtAmount.Left + Me.txtAmount.Width - Me.btnInput.Width
            Me.TxtRemarkItem.Width = Me.btnInput.Left - 5 - Me.TxtRemarkItem.Left
            Me.txtNettPrice.Left = Me.txtAmount.Left + Me.txtAmount.Width - Me.txtNettPrice.Width
            Me.TxtDiscount.Left = Me.txtNettPrice.Left
            Me.TxtQtyMulti.Left = Me.txtNettPrice.Left
            Me.LblDisc.Left = Me.TxtDiscount.Left - 5 - Me.LblDisc.Width
            Me.LblQty.Left = Me.LblDisc.Left
            Me.Label33.Left = Me.LblQty.Left
            Me.CboUnit.Left = Me.txtAmount.Left + Me.txtAmount.Width - Me.CboUnit.Width
            Me.btnFindCurrency2.Left = Me.LblDisc.Left - Me.btnFindCurrency2.Width - 5
            Me.TxtCurrencyCode1.Width = Me.btnFindCurrency2.Left - 5 - Me.TxtCurrencyCode1.Left
            Me.TxtUnitPrice.Width = Me.Label33.Left - 5 - Me.TxtUnitPrice.Left
            Me.TxtDiscAmount1.Width = Me.TxtUnitPrice.Width

            Me.cboAccount.Width = Me.dtpDate.Left + Me.dtpDate.Width - Me.cboAccount.Left
            Me.CboAddressShipTo.Width = Me.cboCustomerAddress.Width + Me.cboCustomerAddress.Left - Me.CboAddressShipTo.Left
            Me.CboAddressCustomerShipTo.Width = Me.CboAddressShipTo.Width
            Me.LlbViewAddressCustomerShipTo.Left = Me.llbAddress.Left
            Me.LlbViewAddressShipTo.Left = Me.llbAddress.Left
            Me.TxtCustomerDescShipTo.Width = Me.TxtCustomerName.Left + Me.TxtCustomerName.Width - Me.TxtCustomerDescShipTo.Left

            Me.ColVAT.Visible = False
            Me.ColVAT2.HeaderText = "VAT"
            Me.btnFind.Focus()
            Me.dtpDate.ShowUpDown = False
            Me.dtpDate.CustomFormat = "dd MMM yyyy"

            Me.txtSubTotalAfterDP.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.txtSubTotalAfterDP.Left
            Me.txtSubTotalBeforeDP.Width = Me.txtSubTotalAfterDP.Width
            Me.TxtShipment.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.TxtShipment.Left
            Me.txtTaxExchangeRate.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.txtTaxExchangeRate.Left
            Me.txtTaxAmount.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.txtTaxAmount.Left
            Me.txtTaxAmount1.Width = Me.txtTaxAmount.Width 'Me.SplitContainer6.Panel2.Width - 5 - Me.txtTaxAmount.Left
            Me.txtTaxAmount1.Left = Me.txtTaxAmount.Left 'Me.SplitContainer6.Panel2.Width - 5 - Me.txtTaxAmount.Left
            Me.txtDutyStamp.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.txtDutyStamp.Left
            Me.txtOtherCharge.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.txtOtherCharge.Left
            Me.TxtRemarks.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.TxtRemarks.Left
            Me.SplitContainer3.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.SplitContainer3.Left
            Me.TxtTotal.Width = Me.SplitContainer6.Panel2.Width - 5 - Me.TxtTotal.Left

            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.SplitContainer3.TabStop = False
            Me.SplitContainer4.TabStop = False
            Me.SplitContainer5.TabStop = False
            Me.SplitContainer6.TabStop = False

            Me.btnUP.BackgroundImage = My.Resources.arrow_up
            Me.btnDown.BackgroundImage = My.Resources.arrow_down
            'Me.txtPriceByMultiUnit.Visible = True
            'Me.ColPriceByMultiUnit.Visible = True
            Atur1()

            'Me.ColAmount.Visible = True
            'Me.ColAmount2.Visible = True
            'Me.ColVAT.Visible = True
            'Me.ColVAT2.Visible = True

            '2025 Core Tax Tahap 2.2
            Me.llbEditNoFakturPajak.Visible = False
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
            'Me.lblSPPBNo.Visible = False
            'Me.TxtSPPBNo.Visible = False
            Me.btnUP.BackgroundImage = Global.Sales_Module.My.Resources.Resources.arrow_up
            Me.btnDown.BackgroundImage = Global.Sales_Module.My.Resources.Resources.arrow_down
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtExchangeRateToUSD_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchangeRateToUSD.GotFocus
        txtExchangeRateToUSD.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtExchangeRateToUSD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExchangeRateToUSD.KeyDown
        Me.ObjCD.Number_KeyDown(txtExchangeRateToUSD, sender, e)
    End Sub

    Private Sub txtExchangeRateToUSD_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExchangeRateToUSD.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtExchangeRateToUSD, sender, e)
    End Sub

    Private Sub txtExchangeRateToUSD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchangeRateToUSD.LostFocus
        txtExchangeRateToUSD.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub

    Private Sub txtExhangeRateToIDR_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.GotFocus
        Me.txtExhangeRateToIDR.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtExhangeRateToIDR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExhangeRateToIDR.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtExhangeRateToIDR, sender, e)
    End Sub

    Private Sub txtExhangeRateToIDR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExhangeRateToIDR.LostFocus
        txtExhangeRateToIDR.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub

    Private Sub TxtRemark_GotFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        TxtRemarks.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        TxtRemarks.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                Dim ij As New FrmCetakInvoice
                If Me.ObjInvoice.CetakInvoice(ij.crv1, Me.txtPrimaryInvoice.Text, Me.dtpDate.Value, _
                Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text), Me.UserCode) = True Then
                    ij.ShowDialog()

                    If MsgBox("Is invoice printed correctly?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        'update printed invoice
                        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
                        DTA.Connection.ConnectionString = Core_Module.GetConstring
                        DTA.UpdateQuery_SetPrintStatus(Me.UserCode, Now, Me.txtPrimaryInvoice.Text)
                        DTA.Dispose()

                        Me.LblPrintStatus.Text = "Print Status: " + Me.ObjInvoice.GetPrintStatus(Me.txtPrimaryInvoice.Text)
                        If Me.LblPrintStatus.Text = "Print Status: " Then
                            Me.LblPrintStatus.Text = "Print Status: -"
                        End If
                        Me.LblPrintStatus.Left = Me.gboDetail.Width - 5 - Me.LblPrintStatus.Width
                    End If
                End If
                ij.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFaktur_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFaktur.Click
        Dim ij As New FrmCetakFaktur
        Dim rowDGV As DataGridViewRow
        Dim JenisFakturPajakStandar As FrmCetakFaktur.EnumJenisFakturStandar = FrmCetakFaktur.EnumJenisFakturStandar.HargaJual

        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.CboFakturPajak.SelectedIndex = 0 Then
                    MsgBox("This feature is only for standard tax invoice.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    Dim ijk As New FrmTransaksiSalesInvoice_JenisFakturPajakStandar
                    If ijk.ShowDialog = Windows.Forms.DialogResult.OK Then
                        JenisFakturPajakStandar = ijk.JenisFakturPajakStandar
                    Else
                        Exit Sub
                    End If
                End If

                ij.JenisFakturPajakStandar = JenisFakturPajakStandar
                ij.primary = Me.txtPrimaryInvoice.Text
                ij.NoFakturPajak = Me.txtNoSeriFakturPajak.Text + "-" + Me.txtNoFakturPajak.Text
                ij.CurrencyCode = Me.txtCurrencyCode.Text
                ij.TotalDitagih = Me.TxtTotal.Text
                ij.TotalPPN = Me.txtTaxAmount.Text
                ij.TotalHargaJual = Me.txtSubTotalBeforeDP.Text
                ij.RatePajak = Me.txtTaxExchangeRate.Text
                ij.TanggalStr = Me.ObjCD.GetDateString_Indo_Panjang(Me.dtpDate.Value)
                rowDGV = Me.dgvItem.Rows(0)
                If rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value <> 1 Then
                    ij.FakturPajakStandarValas = True
                Else
                    ij.FakturPajakStandarValas = False
                End If
                ij.DecimalValas = Me.ObjCurrency.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                ij.CurrencyCodeValas = rowDGV.Cells(Me.ColCurrency.Index).Value

                If Me.CboFakturPajak.SelectedIndex = 0 Then
                    ij.JenisFaktur = FrmCetakFaktur.EnumJenisFakturPajak.Sederhana
                    'ij.Text = "Faktur Pajak Sederhana"
                Else
                    ij.JenisFaktur = FrmCetakFaktur.EnumJenisFakturPajak.Standar
                    'ij.Text = "Faktur Pajak Standar"
                End If
                ij.FormatDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub


    Private Sub KosongkanItem()
        Dim row As DataGridViewRow

        Try
            If Me.dgvItem.RowCount = 0 Then
                If IsNumeric(Me.txtPrimaryCurrency.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = Me.txtPrimaryCurrency.Text
                Else
                    Me.TxtPrimaryCurrency1.Text = "1"
                End If
            Else
                row = Me.dgvItem.Rows(0)
                Me.TxtPrimaryCurrency1.Text = row.Cells(Me.ColPrimaryCurrency.Index).Value
            End If
            Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)

            Me.TxtUnitPrice.Text = "0"
            Me.TxtDiscount.Text = "0"
            Me.txtNettPrice.Text = "0"
            Me.TxtQtyDefault.Text = "0"
            Me.TxtQtyMulti.Text = "0"
            Me.txtAmount.Text = "0"
            Me.TxtDiscAmount1.Text = "0"
            Me.TxtDiscFromPerct.Text = "1"
            Me.TxtRemarkItem.Text = ""
            Me.TxtCodeNo.Text = ""
            Me.txtDesc.Text = ""

            Me.SetGboItem()

            'kosongkan juga cbo lokasi dan cbo batch number
            Me.cboBatchNo.DataSource = Nothing
            Me.cboBatchNo.Refresh()
            Me.cboLocation.DataSource = Nothing
            Me.cboLocation.Refresh()

            Me.cboPriceByMultiUnit.Visible = False
            Me.cboPriceByMultiUnit.TabStop = False
            Me.txtPriceByMultiUnit.Text = "1"
            Me.cboPriceByMultiUnit.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowTaxSetting()
        Dim ObjHakAkses As New Core_Module.ClsUser

        Try
            If ObjHakAkses.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoiceSettingPajak") = True Then
                Dim ij As Sales_Module.FrmTransaksiSalesInvoiceSettingPajak
                Dim j As Windows.Forms.Form

                For Each j In Me.MdiChildren
                    ij = TryCast(j, Sales_Module.FrmTransaksiSalesInvoiceSettingPajak)
                    If IsNothing(ij) = False Then
                        ij.BringToFront()
                        ij.Focus()
                        Exit Sub
                    End If
                Next

                Dim i As New Sales_Module.FrmTransaksiSalesInvoiceSettingPajak
                i.UserCode = Me.UserCode
                ' Set the caption to be unique.
                i.BackColorAll = Me.BackColorAll
                i.MdiParent = Me.ParentForm
                i.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjHakAkses = Nothing
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjInvoice
                .CalculatesBy = Me.cboCalculatesBy.SelectedIndex
                .DownPayment = Me.txtDownPayment.Text
                .DutyStamp = Me.txtDutyStamp.Text
                .ExchangRateToIDR = Me.txtExhangeRateToIDR.Text
                .ExhangeRateToUSD = Me.txtExchangeRateToUSD.Text
                .NoFakturPajak = Me.txtNoFakturPajak.Text
                .NoNPWP = Me.txtNoNPWP.Text
                .NoSeriFakturPajak = Me.txtNoSeriFakturPajak.Text
                .Revisi = Me.txtRevisiNo.Text
                .SubTotalB4DP = Me.txtSubTotalBeforeDP.Text
                .OngkosKirim = Me.TxtShipment.Text
                .TaxExchangeRate = Me.txtTaxExchangeRate.Text
                .CustIdentifierAdd = Me.txtCustomerAddIdentifier.Text
                .NoTransaksi = Me.txtInvoiceNo.Text
                .PaymentTerm = Me.txtPaymentTerm.Text
                .PrimaryCurrency = Me.txtPrimaryCurrency.Text
                .PrimaryCustomer = Me.txtPrimaryCustomer.Text
                .SubTotal = Me.txtSubTotalAfterDP.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .TaxAmount1 = Me.txtTaxAmount1.Text 'new tax amount1
                .TaxPercentage = Me.txtTax.Text
                .Total = Me.TxtTotal.Text
                .UseTax = Me.chkTax.Checked
                .Remarks = Me.TxtRemarks.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .JenisFakturPajak = Me.CboFakturPajak.SelectedIndex
                .NoFakturPajakSederhana = Me.txtNoFakturPajakSederhana.Text
                .NamaSigned = Me.TxtNamaSigned.Text
                .JabatanSigned = Me.TxtJabatanSigned.Text
                If Me.cboShipmentTax.SelectedIndex = 0 Then
                    .ShipmentWithTax = True
                Else
                    .ShipmentWithTax = False
                End If
                If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    .VATDipungut = True
                Else
                    .VATDipungut = False
                End If
                If .JenisFakturPajak = ClsTransaksiSalesInvoice.EnumJenisFakturPajak.FakturPajakStandar Then
                    .UrutanFakturPajak = 0
                Else
                    .UrutanFakturPajak = 0
                End If
                .PrimarySalesPerson = Me.TxtPrimarySalesPerson.Text
                .DiscString = Me.TxtDiscString.Text
                .DiscAmount = Me.TxtDiscAmount.Text

                .EFak_Blok = Me.Txt_E_Faktur_Blok.Text
                .EFak_NamaJalan = Me.Txt_E_Faktur_Jalan.Text
                .EFak_Kelurahan = Me.Txt_E_Faktur_Kelurahan.Text
                .EFak_KodePos = Me.Txt_E_Faktur_KodePos.Text
                .EFak_No = Me.Txt_E_Faktur_No.Text
                .EFak_Telepon = Me.Txt_E_Faktur_NoTelepon.Text
                .EFak_Propinsi = Me.Txt_E_Faktur_Propinsi.Text
                .EFak_RT = Me.Txt_E_Faktur_RT.Text
                .EFak_RW = Me.Txt_E_Faktur_RW.Text
                .EFak_Kabupaten = Me.Txt_E_Faktur_Kabupaten.Text
                .EFak_Kecamatan = Me.Txt_E_Faktur_Kecamatan.Text
                .EFak_KodeJenisTransaksiPajak = Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text
                .PrimaryCurrTax = Me.txtPrimaryCurrTax.Text
                .PrimaryAccount = Me.txtPrimaryAccount.Text
                .OtherCharge = Me.txtOtherCharge.Text

                'jika datanya baru, maka update data alamat pajak 1 dan alamat pajak 2
                If Me.baru = True Then
                    Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, .AlamatPajak1, .AlamatPajak2)
                Else
                    'jika ada perubahan data customer dari invoice lama maka update data alamat customer
                    If Me.ObjInvoice.Get_Primary_Cust(Me.txtPrimaryInvoice.Text) <> Me.txtPrimaryCustomer.Text Then
                        Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, .AlamatPajak1, .AlamatPajak2)
                    End If
                End If
                '//TR : Tambahan baru Nik/Pas
                .OptionNIKPASS = OptionNIKPASS
                .NoNikPass = Me.txtNoNIKPass.Text

                If Me.TBLHeaderCost.Rows.Count = 0 Then
                    Me.rowHeaderCost = Me.TBLHeaderCost.Newtbltrans_salesinvoice_add_dataRow
                    Me.rowHeaderCost.primary = -1
                    Me.rowHeaderCost.primary_inv = Me.txtPrimaryInvoice.Text
                    Me.rowHeaderCost.primary_weight_unit = -1
                    Me.rowHeaderCost.primary_volume_unit = -1
                    Me.rowHeaderCost.Total_Gross_Volume = 0
                    Me.rowHeaderCost.Total_Gross_Weight = 0
                    Me.rowHeaderCost.Total_Nett_Volume = 0
                    Me.rowHeaderCost.Total_Nett_Weight = 0
                    Me.rowHeaderCost.ETA_Port = String.Empty
                    Me.rowHeaderCost.ETA_Date = #1/1/1950#
                    Me.rowHeaderCost.ETD_Port = String.Empty
                    Me.rowHeaderCost.ETD_Date = #1/1/1950#
                    Me.rowHeaderCost.transportation_mode = String.Empty
                    Me.rowHeaderCost.var1_name = String.Empty
                    Me.rowHeaderCost.var2_name = String.Empty
                    Me.rowHeaderCost.var3_name = String.Empty
                    Me.rowHeaderCost.var4_name = String.Empty
                    Me.rowHeaderCost.var5_name = String.Empty
                    Me.rowHeaderCost.var1_value = String.Empty
                    Me.rowHeaderCost.var2_value = String.Empty
                    Me.rowHeaderCost.var3_value = String.Empty
                    Me.rowHeaderCost.var4_value = String.Empty
                    Me.rowHeaderCost.var5_value = String.Empty
                    Me.rowHeaderCost.invoice_type = String.Empty
                    Me.rowHeaderCost.shipping_mark = String.Empty
                    Me.rowHeaderCost.primary_cust_ship_to = Me.PrimaryCustomerShipTo
                    Me.rowHeaderCost.identifier_cust_ship_to = Me.IdentifierAddressCustomerShipTo
                    Me.TBLHeaderCost.Rows.Add(Me.rowHeaderCost)
                End If

                '.SPPBNo = Me.TxtSPPBNo.Text
                '2025 Core Tax Tahap 2.2
                .SPPBDate = Me.SPPBDate
                '.TaxRevised = Me.chkRevisedInvoice.Checked
                '//#2732
                If Me.txtRevisedInvoice.Text = "01" Then
                    .TaxRevised = True
                Else
                    .TaxRevised = False
                End If
                If Me.chkTax.Checked = False Then
                    .NoFakturPajak = ""
                End If


                .PrimaryAddInfo = Me.AddInfo
                .PrimaryFacilityStamp = Me.FacilityStamp
                '2025 Core Tax Tahap 2.2, diupdate ketika semua type bisa dipilih, jadi selectedindex, tanpa percabangan
                .TaxType = Me.cboTaxtype.SelectedIndex
                .NoSeriFakturPajak = Me.cboTaxtype.SelectedItem.ToString
                .DPPNilaiLain = Me.txtDPPNilaiLain.Text

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjInvoice
                Me.cboCalculatesBy.SelectedIndex = .CalculatesBy
                Me.txtNoFakturPajak.Text = .NoFakturPajak
                '//TR : Tambahan baru Nik/Pas
                'Update NPWP 16 Digit
                'Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(.PrimaryCustomer)
                Me.txtNoNPWP.Text = .NoNPWP
                '2025 Core Tax Tahap 2.1
                If txtNoNPWP.Text = "00.000.000.0-000.000" Or txtNoNPWP.Text = "" Or _
                Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" Then
                    'Me.llbUpDateNPWP.Visible = False
                    Me.cboNIKPASS.Visible = True
                    Me.cboNIKPASS.TabStop = True
                    Me.txtNoNIKPass.Visible = True
                    Me.cboNIKPASS.SelectedIndex = .OptionNIKPASS
                    Me.txtNoNIKPass.Text = .NoNikPass
                    Me.txtNoNPWP.Visible = False

                Else
                    'Me.llbUpDateNPWP.Visible = True
                    Me.cboNIKPASS.Visible = False
                    Me.cboNIKPASS.TabStop = False
                    Me.txtNoNIKPass.Visible = False
                    Me.txtNoNPWP.Visible = True
                End If

                Me.txtNoSeriFakturPajak.Text = .NoSeriFakturPajak
                Me.txtRevisiNo.Text = .Revisi

                Me.CboFakturPajak.SelectedIndex = .JenisFakturPajak
                Me.txtPrimaryInvoice.Text = .Primary
                Me.txtInvoiceNo.Text = .NoTransaksi
                Me.txtPaymentTerm.Text = .PaymentTerm
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision(False)
                Me.txtPrimaryCustomer.Text = .PrimaryCustomer
                Me.txtCustomerAddIdentifier.Text = .CustIdentifierAdd
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                Me.chkTax.Checked = .UseTax
                Me.TxtRemarks.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtNoFakturPajakSederhana.Text = .NoFakturPajakSederhana
                Me.TxtNamaSigned.Text = .NamaSigned
                Me.TxtJabatanSigned.Text = .JabatanSigned
                If .ShipmentWithTax = True Then
                    Me.cboShipmentTax.SelectedIndex = 0
                Else
                    Me.cboShipmentTax.SelectedIndex = 1
                End If
                If .VATDipungut = True Then
                    Me.cboPPNDipungut.SelectedIndex = 0
                Else
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If
                Me.TxtPrimarySalesPerson.Text = .PrimarySalesPerson
                Me.TxtSalesPersonCode.Text = Me.ObjSalesPerson.GetCode(Me.TxtPrimarySalesPerson.Text)
                Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
                Me.TxtDiscString.Text = .DiscString

                Me.Txt_E_Faktur_Blok.Text = .EFak_Blok
                Me.Txt_E_Faktur_Jalan.Text = .EFak_NamaJalan
                Me.Txt_E_Faktur_Kelurahan.Text = .EFak_Kelurahan
                Me.Txt_E_Faktur_KodePos.Text = .EFak_KodePos
                Me.Txt_E_Faktur_No.Text = .EFak_No
                Me.Txt_E_Faktur_NoTelepon.Text = .EFak_Telepon
                Me.Txt_E_Faktur_Propinsi.Text = .EFak_Propinsi
                Me.Txt_E_Faktur_RT.Text = .EFak_RT
                Me.Txt_E_Faktur_RW.Text = .EFak_RW
                Me.Txt_E_Faktur_Kabupaten.Text = .EFak_Kabupaten
                Me.Txt_E_Faktur_Kecamatan.Text = .EFak_Kecamatan
                Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = .EFak_KodeJenisTransaksiPajak


                Me.IsiCBOArAkunByTipePerkiraan1()
                If .PrimaryAccount > -1 Then
                    Me.txtPrimaryAccount.Text = .PrimaryAccount
                    Me.cboPrimaryAccount.Text = .PrimaryAccount
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryAccount.DataSource = Nothing
                    Me.txtPrimaryAccount.Text = -1
                End If

                Me.IsiCBOCurrencyTax()
                If .PrimaryCurrTax > -1 Then
                    Me.txtPrimaryCurrTax.Text = .PrimaryCurrTax
                    Me.cboPrimaryCurrTax.Text = .PrimaryCurrTax
                Else
                    Me.cboAccount.DataSource = Nothing
                    Me.cboPrimaryCurrTax.DataSource = Nothing
                    Me.txtPrimaryCurrTax.Text = -1
                End If
                Me.lblCurrTax.Text = Me.cboCurrTax.Text

                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax))
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping(.TaxAmount1, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax)) 'new tax amount1
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(.DutyStamp, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax))
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotalB4DP, Me.DefaultPrecision))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping(.DiscAmount, Me.DefaultPrecision)
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxExchangeRate, 2))
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.DownPayment, Me.DefaultPrecision))
                Me.txtExchangeRateToUSD.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExhangeRateToUSD, Me.DefaultPrecision))
                Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(.OngkosKirim, Me.DefaultPrecision)
                Me.txtTax.Text = .TaxPercentage
                Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.OtherCharge, Me.DefaultPrecision))
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))

                '2025 Core Tax Tahap 2.2
                Me.SPPBDate = .SPPBDate

                'Me.chkRevisedInvoice.Checked = .TaxRevised
                '//#2732
                If .TaxRevised = True Then
                    Me.txtRevisedInvoice.Text = "01"
                Else
                    Me.txtRevisedInvoice.Text = "00"
                End If

                Me.AddInfo = .PrimaryAddInfo
                Me.FacilityStamp = .PrimaryFacilityStamp
                Me.cboTaxtype.SelectedIndex = .TaxType
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(.DPPNilaiLain, Me.DefaultPrecision)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With

            If Me.ApakahAdaAlias = True Then
                Me.chkAlias.Checked = True
            Else
                Me.chkAlias.Checked = False
            End If

            Me.FormatDGV()

            'print status
            Me.LblPrintStatus.Text = "Print Status: " + Me.ObjInvoice.GetPrintStatus(Me.txtPrimaryInvoice.Text)
            If Me.LblPrintStatus.Text = "Print Status: " Then
                Me.LblPrintStatus.Text = "Print Status: -"
            End If
            Me.LblPrintStatus.Left = Me.gboDetail.Width - 5 - Me.LblPrintStatus.Width

            If Me.TBLHeaderCost.Rows.Count > 0 Then
                Me.rowHeaderCost = Me.TBLHeaderCost.Rows(0)
                Me.PrimaryCustomerShipTo = Me.rowHeaderCost.primary_cust_ship_to
                Me.IdentifierAddressCustomerShipTo = Me.rowHeaderCost.identifier_cust_ship_to
            End If

            If Me.PrimaryCustomerShipTo = ObjInvoice.PrimaryCustomer Then
                Me.RboAddressShipTo.Checked = True
                Me.TxtIdentifierAddressShipTo.Text = Me.IdentifierAddressCustomerShipTo
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressShipTo.Text)
            Else
                Me.RboCustomerShipTo.Checked = True
                Me.TxtPrimaryCustomerShipTo.Text = Me.PrimaryCustomerShipTo
                Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomerShipTo.Text)
                Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                Me.TxtIdentifierAddressCustomerShipTo.Text = Me.IdentifierAddressCustomerShipTo
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressCustomerShipTo.Text)
            End If
            
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaAlias() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(Me.ColAliasCode.Index).Value.ToString) = False Or _
                String.IsNullOrEmpty(rowDGV.Cells(Me.ColAliasDesc.Index).Value.ToString) = False Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    'Try
    '    '    Me.baru = True
    '    '    Me.NonAktifkanDetail()
    '    '    Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
    '    '    Me.ObjInvoice.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice
    '    '    Me.Text = "Quick Invoice1" 

    '    '    If Me.objCustom.ApakahBolehImportInvoice = True Then
    '    '        Me.llbImport.Visible = True
    '    '        Me.ProgressBar1.Visible = True
    '    '    End If

    '    '    Me.TxtNamaSigned.Width = Me.SplitContainer3.Panel1.Width - 2
    '    '    Me.TxtJabatanSigned.Width = Me.SplitContainer3.Panel2.Width - 2
    '    '    Me.TxtJabatanSigned.Left = 1

    '    '    Me.LlbViewAddressCustomerShipTo.Left = Me.LlbViewAddressCustomerShipTo.Parent.Width - 5 - Me.LlbViewAddressCustomerShipTo.Width
    '    '    Me.LlbViewAddressShipTo.Left = Me.LlbViewAddressCustomerShipTo.Left
    '    '    Me.CboAddressShipTo.Width = Me.LlbViewAddressCustomerShipTo.Left - 5 - Me.CboAddressShipTo.Left
    '    '    Me.CboAddressCustomerShipTo.Width = Me.CboAddressShipTo.Width
    '    '    Me.TxtCustomerDescShipTo.Width = Me.TxtCustomerDescShipTo.Parent.Width - 5 - Me.TxtCustomerDescShipTo.Left
    '    '    Me.ShowIcon = False

    '    '    '3
    '    '    Me.LinkLabel1.Parent = Me.TabPage1
    '    '    Me.LinkLabel1.BringToFront()
    '    '    Me.txtPrimaryInvoice.Text = "-1"
    '    '    Me.txtPrimaryCustomer.Text = "-1"
    '    '    Me.txtPrimaryCurrency.Text = "-1"
    '    '    Me.txtTax.Text = "0"

    '    '    '2.
    '    '    Me.Atur1()
    '    '    Me.btnUP.BackgroundImage = My.Resources.arrow_up
    '    '    Me.btnDown.BackgroundImage = My.Resources.arrow_down 
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    'End Try
    'End Sub

    Public Function Simpan(ByVal FrmDO As FrmTransaksiSalesInvoiceQuick_DetailNumber) As Boolean
        Try
            Me.SalinKeObjek()
            If Me.ObjInvoice.SimpanBaru(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
            Me.dgvExchageRates, Me.UserCode, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy) Then
                Me.SalinKeField()
                Me.NonAktifkanDetail()
                Me.btnFind.Focus()
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim FrmDO As New FrmTransaksiSalesInvoiceQuick_DetailNumber
        Dim Lanjut As Boolean = True

        Try
            FrmDO.TxtDONumber.Text = ""
            FrmDO.PrimaryInvQuickAcuan = Me.txtPrimaryInvoice.Text
            If Me.txtPrimaryInvoice.Text <> "" And Me.txtPrimaryInvoice.Text <> "-1" Then
                FrmDO.NoDO_Lama = Me.objDO.GetDONO_ByPrimaryAOPQuick(Me.txtPrimaryInvoice.Text)
            Else
                FrmDO.NoDO_Lama = ""
            End If
            FrmDO.BackColorAll = Me.BackColorAll
            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                Lanjut = True
            Else
                If FrmDO.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Lanjut = True
                Else
                    Lanjut = False
                End If
            End If

            'pengecekan no do agar tidak double
            If Lanjut = True Then
                If Me.ApakahDataBenar() = True Then
                    Me.KosongkanItem()
                    Me.SalinKeObjek()
                    If baru = True Then
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.SedangRevisi = True Then
                                If Me.ObjInvoice.SimpanRevisi(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                                Me.dgvExchageRates, Me.UserCode, Me.ObjInvoice.Primary, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy) Then
                                    MsgBox("Invoice Revised", MsgBoxStyle.Information, "Information")
                                    Me.NonAktifkanDetail()
                                    Me.SalinKeField()
                                    Me.btnFind.Focus()
                                    Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                    Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                    Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                    Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                                End If
                            Else
                                If Me.ObjInvoice.SimpanBaru(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                                Me.dgvExchageRates, Me.UserCode, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy) Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.NonAktifkanDetail()
                                    Me.SalinKeField()
                                    Me.btnFind.Focus()
                                    Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                    Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                    Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                    Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                                End If
                            End If
                        End If
                    Else
                        'konfirmasi disimpan atau tidak
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjInvoice.SimpanPerubahan(Me.dgvItem, Me.dgvDP, _
                            Me.dgvItemSR, Me.dgvExchageRates, Me.UserCode, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.NonAktifkanDetail()
                                Me.SalinKeField()
                                Me.btnEdit.Focus()
                                Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            FrmDO.Dispose()
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtRemarks.Text = "" And Me.dgvItem.RowCount = 0 Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.ObjInvoice.CariDetail(Me.txtPrimaryInvoice.Text, Me.dgvItem, _
                        Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me, Me.TblCustomDummy, Me.TblCustomInOutDummy)
                        Me.NonAktifkanDetail()
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiSalesInvoice

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice
            hj.IsQuickInvoice = True

            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjInvoice.CariDetail(hj.PrimaryTransaksi, _
                Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me, Me.TblCustomDummy, Me.TblCustomInOutDummy) _
                = True Then

                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim mPrimaryCurrency As Double = -1
        Dim rowDGV As DataGridViewRow

        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.ObjInvoice.ApakahBolehEdit(Me.ObjInvoice.Primary, Me.UserCode, True) = True Then
                    If Me.ObjInvoice.ApakahBolehEdit_QuickInvoice(Me.ObjInvoice.Primary) = False Then
                        Exit Sub
                    End If
                    If Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = False Then
                        MsgBox("This invoice is not quick invoice. Use sales invoice form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.baru = False
                    Me.SedangRevisi = False
                    Me.ObjInvoice.CariDetail(Me.ObjInvoice.Primary, Me.dgvItem, Me.dgvDP, _
                    Me.dgvItemSR, Me.dgvExchageRates, Me, Me.TblCustomDummy, Me.TblCustomInOutDummy)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()

                    'Exit Sub
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                    'primary currency untuk item mengikuti currency baris pertama dari item
                    If Me.dgvItem.RowCount > 0 Then
                        rowDGV = Me.dgvItem.Rows(0)
                        mPrimaryCurrency = rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value
                        Me.TxtPrimaryCurrency1.Text = mPrimaryCurrency
                        Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = False Then
                    MsgBox("This invoice is not quick invoice. Use sales invoice form", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If Me.ObjInvoice.ApakahBolehHapus(Me.ObjInvoice.Primary, True, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    'Dim NoInvLama As String = Me.txtInvoiceNo.Text.Substring(0, 18) --> berkaitan dengan revisi, untuk invoice quick tidak ada revisi
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjInvoice.Hapus(Me.txtPrimaryInvoice.Text, Me.UserCode, True, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim FrmDO As New FrmTransaksiSalesInvoiceQuick_DetailNumber

        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    FrmDO.PrimaryInvQuickAcuan = Me.txtPrimaryInvoice.Text
                    FrmDO.BackColorAll = Me.BackColorAll
                    FrmDO.TxtDONumber.Text = ""

                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.SedangRevisi = True Then
                                If Me.ObjInvoice.SimpanRevisi(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                                Me.dgvExchageRates, Me.UserCode, Me.ObjInvoice.Primary, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy) = True Then
                                    MsgBox("Invoice Revised", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                    Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                    Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                    Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                    Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                                Else
                                    e.Cancel = True
                                End If
                            Else
                                If Me.ObjInvoice.SimpanBaru(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                                Me.dgvExchageRates, Me.UserCode, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy) = True Then
                                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                    Me.SalinKeField()
                                    Me.NonAktifkanDetail()
                                    Me.btnFind.Focus()
                                    Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                    Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                    Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                    Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                                Else
                                    e.Cancel = True
                                End If
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjInvoice.SimpanPerubahan(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                            Me.dgvExchageRates, Me.UserCode, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                                Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                                Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                                Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                                Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmDO.Dispose()
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim DefaultPrecItem As Double = 0

        Try
            Me.FormatDGV2()

            'untuk exchange rate
            For Each row In Me.dgvExchageRates.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Try
            '4
            If Me.dgvItem.ReadOnly = True Then
                Exit Sub
            End If
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                Dim FrmRatio As New FrmTransaksiQuotation_SetQtyRatio
                Dim row As DataGridViewRow

                If Me.dgvItem.Item(Me.ColPrimaryUnit.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value Or _
                    Me.TxtRemarks.Enabled = False Then
                    Exit Sub
                End If
                Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value.ToString
                Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value.ToString
                FrmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                FrmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                FrmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                FrmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                FrmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                FrmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                If FrmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = FrmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = FrmRatio.TxtRatioDefault.Text
                    row = Me.dgvItem.Rows(e.RowIndex)
                    Me.SetRatioOnDgv(row)
                    Me.SetQtyDefaultOnDgv1(row, FrmRatio.vDefPrecDefault)

                    ''batal jika qty default jadi 0
                    If CDec(Me.dgvItem.Item(Me.ColTotalQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Ratio restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.SetRatioOnDgv(row)
                        Me.SetQtyDefaultOnDgv1(row, FrmRatio.vDefPrecDefault)
                        Exit Sub
                    End If

                    ''batal jika perubahan ratio menyebabkan stok gudang minus
                    If Me.ApakahEditSah(row.Index, False) = False Then
                        MsgBox("Ratio restored to previous value, because it make qty will be minus.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.SetRatioOnDgv(row)
                        Me.SetQtyDefaultOnDgv1(row, FrmRatio.vDefPrecDefault)
                        Exit Sub
                    Else
                        'hitung ulang row
                        Me.Hitung_Amount(Me.dgvItem.Rows(e.RowIndex))
                        'hitung total
                        Me.Hitung_Total()
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColPriceByMultiUnit_Str.Index And Me.TxtRemarks.Enabled = True Then
                If Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value Then
                    Dim FrmRef As New Core_Module.FrmTrans_PriceByMultiUnit
                    FrmRef.UnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                    FrmRef.UnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    FrmRef.Buffer_PriceByMultiUnit = Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value

                    If FrmRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = FrmRef.cboPriceByMultiUnit.SelectedIndex
                        If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = "1" Then
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value.ToString
                        Else
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString
                        End If
                        'hitung ulang row
                        Me.Hitung_Amount(Me.dgvItem.Rows(e.RowIndex))
                        'hitung total
                        Me.Hitung_Total()
                    End If

                    FrmRef = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)
            'hitung ulang nilai total
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
            If Me.chkAlias.Checked = True Then
                If Me.ApakahAdaAlias = True Then
                    Me.chkAlias.Checked = True
                Else
                    Me.chkAlias.Checked = False
                End If
            End If
            If Me.dgvItem.RowCount <= 0 Then
                'Me.txtCustomerCode.ReadOnly = False
                'Me.btnFindCustomer.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            Me.TxtRemarks.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            Me.TxtRemarks.BackColor = Me.txtPrimaryInvoice.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtInvoiceNo.ReadOnly = False And Me.txtInvoiceNo.Enabled = True Then
                    txtInvoiceNo.BackColor = BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtInvoiceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtInvoiceNo.ReadOnly = False And Me.txtInvoiceNo.Enabled = True Then
                    txtInvoiceNo.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        If Me.txtCustomerCode.ReadOnly = False And Me.txtCustomerCode.Enabled = True Then
            txtCustomerCode.BackColor = BackColorAll
        Else
            txtCustomerCode.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub txtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        If Me.txtCustomerCode.ReadOnly = False And Me.txtCustomerCode.Enabled = True Then
            txtCustomerCode.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        txtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        txtCurrencyCode.BackColor = Me.txtPrimaryInvoice.BackColor
    End Sub

    Private Sub txtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentTerm.GotFocus
        txtPaymentTerm.BackColor = BackColorAll
    End Sub

    Private Sub txtPaymentTerm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentTerm.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtPaymentTerm, sender, e, 1000)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentTerm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentTerm.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtPaymentTerm, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentTerm.LostFocus
        txtPaymentTerm.BackColor = Me.txtPrimaryInvoice.BackColor
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    'Private Function GetJmlDGVSR(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryDO As Double, _
    ' ByVal vSecQty As Double, ByVal vPrimarySO As Double, ByVal vBatchNumber As String, ByVal vED As Date) As Double
    '    Try
    '        Dim row As DataGridViewRow
    '        Dim Jml As Double = 0

    '        For Each row In Me.dgvSR.Rows
    '            If vED = #1/1/1950# Then
    '                If IsDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = False Then
    '                    If IsNumeric(row.Cells(Me.ColSecQty_SR.Index).Value) = True Then
    '                        If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo _
    '                        And row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
    '                        row.Cells(Me.ColSecQty_SR.Index).Value = vSecQty And _
    '                        row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
    '                        row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber Then
    '                            Jml = Jml + CDbl(row.Cells(Me.COlTotalQtySR.Index).Value)
    '                        End If
    '                    Else
    '                        If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo And _
    '                        row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
    '                        row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
    '                        row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber Then
    '                            Jml = Jml + CDbl(row.Cells(Me.COlTotalQtySR.Index).Value)
    '                        End If
    '                    End If
    '                End If
    '            Else
    '                If IsDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = True Then
    '                    If IsNumeric(row.Cells(Me.ColSecQty_SR.Index).Value) = True Then
    '                        If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo _
    '                        And row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
    '                        row.Cells(Me.ColSecQty_SR.Index).Value = vSecQty And _
    '                        row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
    '                        row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber And _
    '                        CDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = vED Then
    '                            Jml = Jml + CDbl(row.Cells(Me.COlTotalQtySR.Index).Value)
    '                        End If
    '                    Else
    '                        If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo And _
    '                        row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
    '                        row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
    '                        row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber And _
    '                        CDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = vED Then
    '                            Jml = Jml + CDbl(row.Cells(Me.COlTotalQtySR.Index).Value)
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        Return Jml
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Private Function GetExchangeRate(ByVal vPrimaryTransaksi As Double) As Double
        Dim row As DataGridViewRow
        Try
            For Each row In Me.dgvExchageRates.Rows
                If row.Cells(5).Value = vPrimaryTransaksi Then
                    Return row.Cells(3).Value
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function Get_Price_Exclude(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vTBL As DseTransaksiDO.DataTable10DataTable) As Boolean
        Dim rowDO As DseTransaksiDO.DataTable10Row

        Try
            For Each rowDO In vTBL.Rows
                If rowDO.Primary_Product = vPrimaryProduct And rowDO.Primary_SO = vPrimarySO Then
                    Return rowDO.Price_Exclude
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function Get_Nett_Price_SO(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vTBL As DseTransaksiDO.tbltrans_do_view3_newDataTable) As Double
        Dim rowDO As DseTransaksiDO.tbltrans_do_view3_newRow

        Try
            For Each rowDO In vTBL.Rows
                If rowDO.Primary_Product = vPrimaryProduct And rowDO.Primary_SO = vPrimarySO Then
                    Return rowDO.Nett_Price
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub HitungDiscAmount()
    '    Dim DiscAmount As Decimal
    '    Dim DecimalPlace As Integer

    '    Try
    '        If Me.TxtPrimaryCurrency1.Text <> "" And Me.TxtPrimaryCurrency1.Text <> "-1" Then
    '            DecimalPlace = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)
    '        Else
    '            DecimalPlace = 0
    '        End If

    '        DiscAmount = Math.Round((CDec(Me.TxtUnitPrice.Text) * (CDec(Me.TxtDiscount.Text)) / 100), DecimalPlace, MidpointRounding.AwayFromZero)
    '        Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(DiscAmount, DecimalPlace)
    '        Me.TxtDiscFromPerct.Text = 1
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub HitungDiscPerct()
    '    Dim DiscPerct As Decimal

    '    Try
    '        DiscPerct = Math.Round((CDec(Me.TxtDiscAmount1.Text) / CDec(Me.TxtUnitPrice.Text) * 100), 2, MidpointRounding.AwayFromZero)
    '        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(DiscPerct, 2)
    '        Me.TxtDiscFromPerct.Text = 0
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub Hitung_Amount2()
        Dim DecimalPlace As Integer
        Dim Amount As Decimal = 0
        Dim NettPrice As Decimal = 0
        Dim DiscPerct, DiscAmount As Decimal

        Try
            If Me.TxtPrimaryCurrency1.Text <> "" And Me.TxtPrimaryCurrency1.Text <> "-1" Then
                DecimalPlace = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)
            Else
                DecimalPlace = 0
            End If

            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, DecimalPlace)

            If Me.TxtDiscFromPerct.Text = "1" Then
                DiscAmount = Math.Round((CDec(Me.TxtUnitPrice.Text) * (CDec(Me.TxtDiscount.Text)) / 100), DecimalPlace, MidpointRounding.AwayFromZero)
                Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(DiscAmount, DecimalPlace)
            Else
                DiscPerct = Math.Round((CDec(Me.TxtDiscAmount1.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2, MidpointRounding.AwayFromZero)
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(DiscPerct, 2)
            End If

            NettPrice = Math.Round(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount1.Text), DecimalPlace, MidpointRounding.AwayFromZero)
            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, DecimalPlace)

            If Me.txtPriceByMultiUnit.Text = "" Then
                Me.txtPriceByMultiUnit.Text = "1"
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtNettPrice.Text) * _
                CDbl(Me.TxtQtyMulti.Text), DecimalPlace), DecimalPlace, ClsCekData.EnumTipeRounding.Middlepoint)
            Else
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtNettPrice.Text) * _
                CDbl(Me.TxtQtyDefault.Text), DecimalPlace), DecimalPlace, ClsCekData.EnumTipeRounding.Middlepoint)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Hitung_Amount(ByVal row As DataGridViewRow)
        Dim DefaultPrecisionRow As Integer
        Dim DiscPerct As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Amount_MiddlePoint(row)
                Exit Sub
            End If
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            DiscPerct = row.Cells(Me.ColDiscFromPrect.Index).Value
            If DiscPerct = 1 Then
                row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecisionRow))
            Else
                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)), 2))
            End If
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(row.Cells(Me.ColUnitPrice.Index).Value) - CDec(row.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecisionRow))

            'hitung amount
            'If Me.MetodeHitungInclude = EnumHitungInclude.UnitPriceInclude Then
            '    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
            '        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
            '        (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), _
            '        DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
            '        DefaultPrecisionRow))
            '    Else
            '        If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
            '            row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
            '            (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100) _
            '            * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow, MidpointRounding.AwayFromZero), DefaultPrecisionRow))
            '        Else
            '            row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
            '            (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
            '            DefaultPrecisionRow))
            '        End If
            '    End If
            'ElseIf Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
            If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, _
                    DefaultPrecisionRow))
                Else
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQty.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, _
                    DefaultPrecisionRow))
                End If
            Else
                If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value _
                        * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQty.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value _
                        * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow))
                    End If
                Else
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value, _
                        DefaultPrecisionRow))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQty.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value, _
                        DefaultPrecisionRow))
                    End If
                End If
            End If
            'End If

            With row
                .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((CDec(.Cells(Me.ColNettPrice.Index).Value) * CDec(Me.txtTax.Text)) / 100, DefaultPrecisionRow)
                If DefaultPrecisionRow = 0 Then
                    If .Cells(Me.ColNettPrice.Index).Value.ToString.EndsWith("55") Then
                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT.Index).Value) - 1, DefaultPrecisionRow)
                    End If
                ElseIf DefaultPrecisionRow = 2 Then
                    If .Cells(Me.ColNettPrice.Index).Value.ToString.EndsWith("55") Then
                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT.Index).Value) - 0.01, DefaultPrecisionRow)
                    End If
                End If
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecisionRow)
                Else
                    .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQty.Index).Value), DefaultPrecisionRow)
                End If
                .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), DefaultPrecisionRow)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Public Sub Hitung_Amount_MiddlePoint(ByVal row As DataGridViewRow)
        Dim DefaultPrecisionRow As Integer
        Dim DiscPerct As Integer = 0

        Try
            DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
            'hitung nett price
            DiscPerct = row.Cells(Me.ColDiscFromPrect.Index).Value
            If DiscPerct = 1 Then
                row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecisionRow))
            Else
                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)), 2))
            End If
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(row.Cells(Me.ColUnitPrice.Index).Value) - CDec(row.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecisionRow))

            'hitung amount
            'If Me.MetodeHitungInclude = EnumHitungInclude.UnitPriceInclude Then
            '    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
            '        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
            '        (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), _
            '        DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
            '        DefaultPrecisionRow))
            '    Else
            '        If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
            '            row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
            '            (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100) _
            '            * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow, MidpointRounding.AwayFromZero), DefaultPrecisionRow))
            '        Else
            '            row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
            '            (row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
            '            DefaultPrecisionRow))
            '        End If
            '    End If
            'ElseIf Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
            If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, _
                    DefaultPrecisionRow))
                Else
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQty.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, _
                    DefaultPrecisionRow))
                End If
            Else
                If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value _
                        * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQty.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value _
                        * CDbl(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow))
                    End If
                Else
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQtyMulti.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value, _
                        DefaultPrecisionRow))
                    Else
                        row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(row.Cells(Me.ColTotalQty.Index).Value) * _
                        row.Cells(Me.ColNettPrice.Index).Value, _
                        DefaultPrecisionRow))
                    End If
                End If
            End If
            'End If

            With row
                '.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((CDec(.Cells(Me.ColNettPrice.Index).Value) * CDec(Me.txtTax.Text)) / 100, DefaultPrecisionRow)
                'If DefaultPrecisionRow = 0 Then
                '    If .Cells(Me.ColNettPrice.Index).Value.ToString.EndsWith("55") Then
                '        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT.Index).Value) - 1, DefaultPrecisionRow)
                '    End If
                'ElseIf DefaultPrecisionRow = 2 Then
                '    If .Cells(Me.ColNettPrice.Index).Value.ToString.EndsWith("55") Then
                '        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT.Index).Value) - 0.01, DefaultPrecisionRow)
                '    End If
                'End If
                'If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                '    .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecisionRow)
                'Else
                '    .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQty.Index).Value), DefaultPrecisionRow)
                'End If

                .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping((CDec(.Cells(Me.ColAmount.Index).Value) * CDec(Me.txtTax.Text)) / 100, DefaultPrecisionRow)
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) / CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecisionRow)
                Else
                    .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) / CDec(.Cells(Me.ColTotalQty.Index).Value), DefaultPrecisionRow)
                End If
                .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), DefaultPrecisionRow)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTaxRate(Optional ByVal vSetNewTaxRate As Boolean = True)
        Dim Rate As Decimal = 0

        Try
            If Me.chkTax.Checked = True Then
                If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                    Me.txtTaxExchangeRate.Enabled = False
                    Me.txtTaxExchangeRate.ReadOnly = True
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Me.txtTaxExchangeRate.TabStop = False
                    If vSetNewTaxRate = True Then
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Me.ObjCurrency.GetDefaultPrecission(1))
                    End If
                Else
                    Me.txtTaxExchangeRate.Enabled = True
                    Me.txtTaxExchangeRate.ReadOnly = False
                    Me.txtTaxExchangeRate.BackColor = Me.txtPrimaryInvoice.BackColor
                    Me.txtTaxExchangeRate.TabStop = True
                    If vSetNewTaxRate = True Then
                        Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                        If Rate = 0 Then
                            Rate = 1
                        End If
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Me.ObjTaxRate.DefaultPrec)
                    End If
                End If
            Else
                If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                    Me.txtTaxExchangeRate.Enabled = False
                    Me.txtTaxExchangeRate.ReadOnly = True
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Me.txtTaxExchangeRate.TabStop = False
                    If vSetNewTaxRate = True Then
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Me.ObjCurrency.GetDefaultPrecission(1))
                    End If
                Else
                    Me.txtTaxExchangeRate.Enabled = True
                    Me.txtTaxExchangeRate.ReadOnly = False
                    Me.txtTaxExchangeRate.BackColor = Me.txtPrimaryInvoice.BackColor
                    Me.txtTaxExchangeRate.TabStop = True
                    If vSetNewTaxRate = True Then
                        Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                        If Rate = 0 Then
                            Rate = 1
                        End If
                        Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Me.ObjTaxRate.DefaultPrec)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDefaultPrecision(Optional ByVal vHitungTotal As Boolean = True)
        Try
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
            Me.FormatDGV()
            If vHitungTotal = True Then
                Me.Hitung_Total()
                Me.HitungVATItem()
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotalBeforeDP.Text, DefaultPrecision))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotalAfterDP.Text, DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxAmount.Text, 0))
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxAmount1.Text, 0)) 'new tax amount1
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim Dummy As String = ""
        Try
            If Me.chkTax.Checked = True Then
                Me.txtTax.ReadOnly = False
                Me.txtTax.TabStop = True
                Me.txtTax.BackColor = Me.txtPrimaryInvoice.BackColor
                '2025 Core Tax Tahap 2.2
                'txtNoFakturPajakSederhana.Visible = False
                'txtNoSeriFakturPajak.Visible = True
                'Label8.Visible = True
                'txtNoFakturPajak.Visible = True
                'llbEditNoFakturPajak.Visible = True

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.TabStop = False
                Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Me.txtTax.Text = "0"
                '2025 Core Tax Tahap 2.2
                'txtNoFakturPajakSederhana.Visible = True
                'txtNoSeriFakturPajak.Visible = False
                'Label8.Visible = False
                'txtNoFakturPajak.Visible = False
                'llbEditNoFakturPajak.Visible = False
            End If

            If Me.TxtRemarks.Enabled = True Then
                Me.SetTaxRate()

                'hitung ulang nilai vat, vat2, dan sub total + vat2 untuk setiap baris dgv
                Me.HitungVATItem()
                Me.Hitung_Total()

                If Me.chkTax.Checked = True Then
                    '2025 Core Tax Tahap 2.2
                    'Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    'Me.llbEditNoFakturPajak.Enabled = True
                    Me.cboTaxtype.SelectedIndex = 0
                    Me.cboTaxtype.Enabled = True
                    Me.cboTaxtype.TabStop = True
                    Me.cboPPNDipungut.Enabled = True
                    Me.cboPPNDipungut.TabStop = True

                    '//#2732
                    Me.txtRevisedInvoice.Enabled = True
                    Me.txtNoFakturPajak.Enabled = True
                Else
                    Me.llbEditNoFakturPajak.Enabled = False
                    '2025 Core Tax Tahap 2.2
                    'Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    'If Me.txtNoFakturPajak.Text.Length >= 8 Then
                    '    Me.txtNoFakturPajak.Text = Me.txtNoFakturPajak.Text.Substring(0, 8) + "00000000"
                    'End If
                    Me.cboTaxtype.SelectedIndex = 0
                    Me.cboTaxtype.Enabled = False
                    Me.cboTaxtype.TabStop = False
                    Me.cboPPNDipungut.Enabled = False
                    Me.cboPPNDipungut.TabStop = False
                    '//#2732
                    Me.txtRevisedInvoice.Enabled = False
                    Me.txtNoFakturPajak.Enabled = False
                End If
                Me.cboPPNDipungut.SelectedIndex = 0
            End If

            If Me.chkTax.Checked = True Then
                Me.ColVAT2.Visible = True
                Me.ColAmount2.Visible = True
            Else
                Me.ColVAT2.Visible = False
                Me.ColAmount2.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungVATItem()
        Dim rowDGV As DataGridViewRow
        Dim DefaultPrecisionRow As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.HitungVATItem_MiddlePoint()
                Exit Sub
            End If
            If Me.dgvItem.Rows.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(0)
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvItem.Rows
                With rowDGV
                    .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((CDec(.Cells(Me.ColNettPrice.Index).Value) * CDec(Me.txtTax.Text)) / 100, DefaultPrecisionRow)
                    If rowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecisionRow)
                    Else
                        .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQty.Index).Value), DefaultPrecisionRow)
                    End If
                    .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), DefaultPrecisionRow)
                End With
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub HitungVATItem_MiddlePoint()
        Dim rowDGV As DataGridViewRow
        Dim DefaultPrecisionRow As Integer = 0

        Try
            If Me.dgvItem.Rows.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(0)
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
            End If
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvItem.Rows
                With rowDGV
                    .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping((CDec(.Cells(Me.ColAmount.Index).Value) * CDec(Me.txtTax.Text)) / 100, DefaultPrecisionRow)
                    If rowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT2.Index).Value / CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecisionRow)
                    Else
                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT2.Index).Value / CDec(.Cells(Me.ColTotalQty.Index).Value), DefaultPrecisionRow)
                    End If
                    .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), DefaultPrecisionRow)
                End With
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        If Me.txtCustomerCode.ReadOnly = True Then
            Exit Sub
        End If
        '2025 Core Tax Tahap 2.1
        Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    If Me.CboAddressShipTo.Items.Count > 0 Then
                        CboAddressShipTo.SelectedIndex = 0
                    End If
                    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                    Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                    '//TR : Tambahan baru Nik/Pas
                    '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                    End If
                    '2025 Core Tax Tahap 2.1
                    If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                    Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                    Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                        'Me.llbUpDateNPWP.Visible = False
                        Me.cboNIKPASS.Visible = True
                        Me.cboNIKPASS.TabStop = True
                        Me.txtNoNIKPass.Visible = True
                        Me.txtNoNPWP.Visible = False
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                        Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                        If Me.baru = True Then
                            Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNIKPass.Text = Me.ObjInvoice.GetNoNikPassbyInvoice(Me.txtPrimaryInvoice.Text)
                        End If
                    Else
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If
                    Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                    Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                    Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                    Me.cboCustomerAddress.SelectedIndex = 0
                    Me.TxtSalesPersonCode.Focus()
                    Me.txtPaymentTerm.Text = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                        If Me.CboAddressShipTo.Items.Count > 0 Then
                            CboAddressShipTo.SelectedIndex = 0
                        End If
                        Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                        Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                        '//TR : Tambahan baru Nik/Pas
                        '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                        If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                            NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                        End If
                        '2025 Core Tax Tahap 2.1
                        If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                        Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                        Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                            'Me.llbUpDateNPWP.Visible = False
                            Me.cboNIKPASS.Visible = True
                            Me.cboNIKPASS.TabStop = True
                            Me.txtNoNIKPass.Visible = True
                            Me.txtNoNPWP.Visible = False
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                            Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                            If Me.baru = True Then
                                Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                            Else
                                Me.txtNoNIKPass.Text = Me.ObjInvoice.GetNoNikPassbyInvoice(Me.txtPrimaryInvoice.Text)
                            End If
                        Else
                            'Me.llbUpDateNPWP.Visible = True
                            Me.cboNIKPASS.Visible = False
                            Me.cboNIKPASS.TabStop = False
                            Me.txtNoNIKPass.Visible = False
                            Me.txtNoNPWP.Visible = True
                        End If
                        Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                        Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                        Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.TxtSalesPersonCode.Focus()
                        Me.txtPaymentTerm.Text = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)
                        '//TR : Tambahan baru Nik/Pas
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.CboAddressShipTo.Items.Clear()
                        Me.TxtIdentifierAddressShipTo.Text = ""
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtNoSeriFakturPajak.Clear()
                        Me.txtCustomerCode.Focus()
                        Me.txtPaymentTerm.Text = "0"
                        Me.Txt_E_Faktur_Jalan.Text = ""
                        Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)

                        '//TR : Tambahan baru Nik/Pas
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        If Me.txtCustomerCode.ReadOnly = True Then
            Exit Sub
        End If
        '2025 Core Tax Tahap 2.1
        Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.GetJml(Me.txtCustomerCode.Text) = 1 Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    If Me.CboAddressShipTo.Items.Count > 0 Then
                        Me.CboAddressShipTo.SelectedIndex = 0
                    End If
                    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                    Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                    '//TR : Tambahan baru Nik/Pas
                    '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                    End If
                    '2025 Core Tax Tahap 2.1
                    If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                    Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                    Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                        'Me.llbUpDateNPWP.Visible = False
                        Me.cboNIKPASS.Visible = True
                        Me.cboNIKPASS.TabStop = True
                        Me.txtNoNIKPass.Visible = True
                        Me.txtNoNPWP.Visible = False
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                        Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                        If Me.baru = True Then
                            Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNIKPass.Text = Me.ObjInvoice.GetNoNikPassbyInvoice(Me.txtPrimaryInvoice.Text)
                        End If
                    Else
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If
                    Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                    Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                    Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                    Me.cboCustomerAddress.SelectedIndex = 0
                    Me.txtPaymentTerm.Text = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                        If Me.CboAddressShipTo.Items.Count > 0 Then
                            Me.CboAddressShipTo.SelectedIndex = 0
                        End If
                        Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                        Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                        '//TR : Tambahan baru Nik/Pas
                        '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                        If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                            NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                        End If
                        '2025 Core Tax Tahap 2.1
                        If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                        Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                        Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                            'Me.llbUpDateNPWP.Visible = False
                            Me.cboNIKPASS.Visible = True
                            Me.cboNIKPASS.TabStop = True
                            Me.txtNoNIKPass.Visible = True
                            Me.txtNoNPWP.Visible = False
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                            Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                            If Me.baru = True Then
                                Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                            Else
                                Me.txtNoNIKPass.Text = Me.ObjInvoice.GetNoNikPassbyInvoice(Me.txtPrimaryInvoice.Text)
                            End If
                        Else
                            'Me.llbUpDateNPWP.Visible = True
                            Me.cboNIKPASS.Visible = False
                            Me.cboNIKPASS.TabStop = False
                            Me.txtNoNIKPass.Visible = False
                            Me.txtNoNPWP.Visible = True
                        End If

                        Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                        Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                        Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.TxtSalesPersonCode.Focus()
                        Me.txtPaymentTerm.Text = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.CboAddressShipTo.Items.Clear()
                        Me.TxtIdentifierAddressShipTo.Text = ""
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtNoSeriFakturPajak.Clear()
                        Me.txtCustomerCode.Focus()
                        Me.txtPaymentTerm.Text = "0"
                        Me.Txt_E_Faktur_Jalan.Text = ""
                        Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                        '//TR : Tambahan baru Nik/Pas
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If
                End If
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.CboAddressShipTo.Items.Clear()
                Me.TxtIdentifierAddressShipTo.Text = ""
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtNoSeriFakturPajak.Clear()
                Me.txtPaymentTerm.Text = "0"
                Me.Txt_E_Faktur_Jalan.Text = ""
                Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                '//TR : Tambahan baru Nik/Pas
                'Me.llbUpDateNPWP.Visible = True
                Me.cboNIKPASS.Visible = False
                Me.cboNIKPASS.TabStop = False
                Me.txtNoNIKPass.Visible = False
                Me.txtNoNPWP.Visible = True
            End If
            Me.UpdateSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        '2025 Core Tax Tahap 2.1
        Dim NPWPType As Core_Module.ClsTrans_SettingNPWPNo.TipeNPWP = ClsTrans_SettingNPWPNo.TipeNPWP.Company
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.SelectedIndex = 0
                Me.CboAddressShipTo.SelectedIndex = 0
                Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                '//TR : Tambahan baru Nik/Pas
                '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    '2025 Core Tax Tahap 2.1
                    Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                    NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                Else
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                End If
                '2025 Core Tax Tahap 2.1
                If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" _
                Or Me.txtNoNPWP.Text = "0000 0000 0000 0000" Or Me.txtNoNPWP.Text = "0000000000000000" _
                Or NPWPType = ClsTrans_SettingNPWPNo.TipeNPWP.Individual Then
                    'Me.llbUpDateNPWP.Visible = False
                    Me.cboNIKPASS.Visible = True
                    Me.cboNIKPASS.TabStop = True
                    Me.txtNoNIKPass.Visible = True
                    Me.txtNoNPWP.Visible = False
                    '2025 Core Tax Tahap 2.1
                    Me.txtNoNPWP.Text = "0000 0000 0000 0000"

                    Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                    If Me.baru = True Then
                        Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNIKPass.Text = Me.ObjInvoice.GetNoNikPassbyInvoice(Me.txtPrimaryInvoice.Text)
                    End If
                Else
                    'Me.llbUpDateNPWP.Visible = True
                    Me.cboNIKPASS.Visible = False
                    Me.cboNIKPASS.TabStop = False
                    Me.txtNoNIKPass.Visible = False
                    Me.txtNoNPWP.Visible = True
                End If

                Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                Me.TxtSalesPersonCode.Focus()
                Me.txtPaymentTerm.Text = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.CboAddressShipTo.Items.Clear()
                Me.TxtIdentifierAddressShipTo.Text = ""
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtNoSeriFakturPajak.Clear()
                Me.txtCustomerCode.Focus()
                Me.txtPaymentTerm.Text = "0"
                Me.Txt_E_Faktur_Jalan.Text = ""
                Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                '//TR : Tambahan baru Nik/Pas
                Me.cboNIKPASS.Visible = False
                Me.cboNIKPASS.TabStop = False
                Me.txtNoNIKPass.Visible = False
                Me.txtNoNPWP.Visible = True
            End If
            Me.UpdateSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub cboCustomerAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustomerAddress.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.cboCustomerAddress.SelectedIndex <> -1 Then
                    Me.txtCustomerAddIdentifier.Text = Me.ObjCustomer.GetIdentifierAdd(Me.txtPrimaryCustomer.Text, _
                    Me.cboCustomerAddress.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyDown
        Try
            Me.buffer = Me.txtTax.Text
            If Me.txtTax.ReadOnly = False Then
                Me.ObjCD.Number_KeyDown(Me.txtTax, sender, e)
            End If
            If IsNumeric(Me.txtTax.Text) = True Then
                If CDbl(Me.txtTax.Text) > 100 Then
                    Me.txtTax.Text = Me.buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyUp
        Try
            If Me.txtTax.ReadOnly = False Then
                Me.ObjCD.Number_KeyUp(Me.txtTax, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerName.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If Me.txtPrimaryCustomer.Text <> "-1" And Me.txtPrimaryCustomer.Text <> "" Then
                If Me.TxtCustomerName.Text <> "" Then
                    Me.Text = "Quick Invoice: " + Me.txtCustomerCode.Text + " - " + Me.TxtCustomerName.Text
                Else
                    Me.Text = "Quick Invoice"
                End If
            Else
                Me.Text = "Quick Invoice"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
            If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso CDec(Me.txtPrimaryInvoice.Text) > 0) Then
                If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                    MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If

                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.txtCustomerAddIdentifier.Text
                ij.PrimaryCust = Me.txtPrimaryCustomer.Text

                If Me.TxtRemarks.Enabled = False Then
                    ij.cboCustAddCode.Enabled = False
                Else
                    ij.btnOK.BringToFront()
                End If

                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiPaymentTerm()
        Dim Minimum As Double = 0
        Dim Dummy As Double = 0
        Dim ListPrimarySO As New List(Of Double)
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If ListPrimarySO.Contains(rowDGV.Cells(Me.ColPrimarySO.Index).Value) = False Then
                    ListPrimarySO.Add(rowDGV.Cells(Me.ColPrimarySO.Index).Value)
                    Dummy = Me.ObjSO.GetPaymentTerm(rowDGV.Cells(Me.ColPrimarySO.Index).Value)
                    If rowDGV.Index = 0 And Dummy <> -1 Then
                        'untuk baris pertama langsung dicopy aja
                        Minimum = Dummy
                    ElseIf Minimum > Dummy And Dummy <> -1 Then
                        'untuk baris selanjutnya dicopy jika memang lebih kecil dari minimum
                        Minimum = Dummy
                    End If
                End If
            Next

            If CDbl(Me.txtPaymentTerm.Text) <> CDbl(Minimum) Then
                If MsgBox("Update payment term?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.txtPaymentTerm.Text = Me.ObjCD.DigitGrouping(Minimum)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimarySO = Nothing
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Me.buffer = e.Row.Cells(Me.ColPrimaryDO.Index).Value
    End Sub

    Public Sub SetCalculateBy()
        Try
            If Me.TxtRemarks.Enabled = True Then
                'reset data item exchange rate
                Me.dgvExchageRates.Rows.Clear()
                If Me.cboCalculatesBy.SelectedIndex = 0 Then
                    'hitung berdasarkan sales invoice
                    If Me.dgvItem.RowCount = 0 Then
                        'kalo ga ada anak, maka exc
                        Me.dgvExchageRates.Rows.Add("0", Me.txtInvoiceNo.Text, Me.txtCurrencyCode.Text, _
                        1, "", Me.txtPrimaryInvoice.Text, Me.txtPrimaryCurrency.Text)
                    Else
                        Dim rowDGV As DataGridViewRow
                        rowDGV = Me.dgvItem.Rows(0)
                        Me.dgvExchageRates.Rows.Add("0", Me.txtInvoiceNo.Text, rowDGV.Cells(Me.ColCurrency.Index).Value, _
                        1, "", Me.txtPrimaryInvoice.Text, rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value)
                    End If
                ElseIf Me.cboCalculatesBy.SelectedIndex = 1 Then
                    'hitung berdasarkan delivery order
                    Dim Ketemu As Boolean = False
                    Dim rowER, rowItem As DataGridViewRow

                    For Each rowItem In Me.dgvItem.Rows
                        Ketemu = False
                        For Each rowER In Me.dgvExchageRates.Rows
                            If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimaryDO.Index).Value Then
                                Ketemu = True
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan exchange rate baru 
                            Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(Me.ColDONO.Index).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
                            1, "", rowItem.Cells(Me.ColPrimaryDO.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
                        End If
                    Next
                ElseIf Me.cboCalculatesBy.SelectedIndex = 2 Then
                    'hitung berdasarkan sales order
                    Dim Ketemu As Boolean = False
                    Dim rowER, rowItem As DataGridViewRow

                    For Each rowItem In Me.dgvItem.Rows
                        Ketemu = False
                        For Each rowER In Me.dgvExchageRates.Rows
                            If rowER.Cells(5).Value = rowItem.Cells(Me.ColPrimarySO.Index).Value Then
                                Ketemu = True
                                Exit For
                            End If
                        Next
                        If Ketemu = False Then
                            'tambahkan exchange rate baru 
                            Me.dgvExchageRates.Rows.Add("0", rowItem.Cells(Me.ColSONO.Index).Value, rowItem.Cells(Me.ColCurrency.Index).Value, _
                            1, "", rowItem.Cells(Me.ColPrimarySO.Index).Value, rowItem.Cells(Me.ColPrimaryCurrency.Index).Value)
                        End If
                    Next
                End If
                If Me.txtCurrencyCode.Text <> "" Then
                    Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate to " + Me.txtCurrencyCode.Text
                Else
                    Me.dgvExchageRates.Columns(Me.ColExchangeRate.Index).HeaderText = "Exchange Rate"
                End If
                'hitung ulang
                Me.Hitung_Total()
                'set number
                Me.ObjCD.SetNumber(Me.dgvExchageRates)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCalculatesBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCalculatesBy.SelectedIndexChanged
        Try
            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDutyStamp.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtDutyStamp, sender, e)
    End Sub

    Private Sub txtDutyStamp_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDutyStamp.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtDutyStamp, sender, e, False)
    End Sub

    Private Sub txtTaxExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Enter
        Try
            If Me.txtTaxExchangeRate.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTaxExchangeRate, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtTaxExchangeRate, sender, e)
    End Sub

    Private Sub txtTaxExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxExchangeRate.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtTaxExchangeRate, sender, e)
    End Sub

    Private Sub dgvExchageRates_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvExchageRates.CellBeginEdit
        Try
            'yang bisa diedit hanya:
            '1. exchange rate
            '2. remarks
            If e.ColumnIndex = 3 Or e.ColumnIndex = 4 Then
                Me.buffer = Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value
                If e.ColumnIndex = Me.ColExchangeRate.Index Then
                    If Me.txtPrimaryCurrency.Text = Me.dgvExchageRates.Rows(e.RowIndex).Cells(Me.ColPrimaryCurrency2.Index).Value Then
                        'kalo primary currency = primary currency dgv exchange maka nilainya =1
                        Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = 1
                        e.Cancel = True
                    End If
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvExchageRates_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvExchageRates.CellEndEdit
        Try
            'pastikan diisi
            If IsNothing(Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex = 3 Then
                    'pastikan nilainya positif
                    If Me.ObjCD.CekBilanganPositif(Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                        "Information")
                        Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'digit grouping
                    Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value)
                    'hitung kembali total invoice
                    Me.Hitung_Total()
                End If
            Else
                If e.ColumnIndex = 4 Then
                    Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = ""
                Else
                    Me.dgvExchageRates.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddDP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddDP.LinkClicked
        Dim ij As New List(Of Double)
        Dim row As DataGridViewRow

        Try
            'tambahkan item DP
            Me.dgvDP.Rows.Clear()
            For Each row In Me.dgvItem.Rows
                If ij.Contains(Me.dgvItem.Item(Me.ColPrimarySO.Index, row.Index).Value) = False Then
                    ij.Add(Me.dgvItem.Item(Me.ColPrimarySO.Index, row.Index).Value)
                End If
            Next
            Me.ObjInvoice.IsiDgvDP(ij, Me.dgvDP, Me.txtPrimaryCurrency.Text, Me.txtPrimaryInvoice.Text)
            Me.ObjCD.SetNumber(Me.dgvDP)
            Me.FormatDGV()
            Me.Hitung_Total()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDP.CellBeginEdit
        Try
            'yang bisa diedit hanya remarks
            If Me.TxtRemarks.Enabled = True Then
                If e.ColumnIndex <> 4 Then
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetSR() As String
        Dim ListPrimary As New List(Of Double)
        Dim row As DataGridViewRow
        Dim rowSR As dseTransaksiSalesReturn.DataTable3Row
        Dim SR As String = ""

        Try
            'cari data SR, berdasarkan data do yang dituliskan
            For Each row In Me.dgvItem.Rows
                If ListPrimary.Contains(row.Cells(Me.ColPrimaryDO.Index).Value) = False Then
                    ListPrimary.Add(row.Cells(Me.ColPrimaryDO.Index).Value)
                End If
            Next
            Me.ObjSR.IsiTBLSRByPrimaryDO(ListPrimary, TBLSR)

            For Each rowSR In Me.TBLSR.Rows
                If Me.ApakahAdaDalamDGVReturn(rowSR.Sales_Return_No) = False Then
                    If SR = "" Then
                        SR = rowSR.Sales_Return_No
                    Else
                        SR = SR + rowSR.Sales_Return_No
                    End If
                End If
            Next

            Return SR
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ListPrimary = Nothing
        End Try
    End Function

    Private Function ApakahAdaDalamDGVReturn(ByVal vSRNumber As String) As Boolean
        Dim rowdgv As DataGridViewRow

        Try
            For Each rowdgv In Me.dgvItemSR.Rows
                If rowdgv.Cells(1).Value = vSRNumber Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub FormatDGVSR()
        Dim row As DataGridViewRow

        Try
            Dim AdaSec As Boolean = False
            Dim AdaExpiry As Boolean = False
            Dim AdaBatchNo As Boolean = False

            For Each row In Me.dgvSR.Rows
                row.Cells(Me.COlTotalQtySR.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.COlTotalQtySR.Index).Value)
                row.Cells(Me.ColQtyMaster_SR.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyMaster_SR.Index).Value)
                row.Cells(Me.ColSecQty_SR.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColSecQty_SR.Index).Value)
                row.Cells(Me.ColQtyPerMaster_SR.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQtyPerMaster_SR.Index).Value)
                row.Cells(Me.COlTotalQtySR.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.COlTotalQtySR.Index).Value)
                If IsNumeric(row.Cells(Me.ColSecQty_SR.Index).Value) = True Then
                    If row.Cells(Me.ColSecQty_SR.Index).Value <> -1 Then
                        AdaSec = True
                    End If
                End If
                If IsDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = True Then
                    If row.Cells(Me.ColExpiryDate_SR.Index).Value <> #1/1/1950# Then
                        AdaExpiry = True
                        row.Cells(Me.ColExpiryDate_SR.Index).Value = Format(CDate(row.Cells(Me.ColExpiryDate_SR.Index).Value), "dd MMM yyyy")
                    Else
                        row.Cells(Me.ColExpiryDate_SR.Index).Value = "-"
                    End If
                End If
                If row.Cells(Me.ColBatchNo_SR.Index).Value <> "" Then
                    AdaBatchNo = True
                End If
            Next

            If AdaSec = True Then
                Me.ColSecQty_SR.Visible = True
                Me.ColSecUnit_SR.Visible = True
            Else
                Me.ColSecQty_SR.Visible = False
                Me.ColSecUnit_SR.Visible = False
            End If
            If AdaExpiry = True Then
                Me.ColExpiryDate_SR.Visible = True
            Else
                Me.ColExpiryDate_SR.Visible = False
            End If
            If AdaBatchNo = True Then
                Me.ColBatchNo_SR.Visible = True
            Else
                Me.ColBatchNo_SR.Visible = False
            End If
            Me.ObjCD.SetNumber(Me.dgvSR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            'tandai
            If Me.vFromQuick = False Then
                If Me.TxtRemarks.Enabled = True Then
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.Enter
        Try
            If Me.txtDutyStamp.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtDutyStamp, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub dgvDP_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvDP.UserDeletedRow
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.SetNumber(Me.dgvDP)
                Me.FormatDGV()
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboFakturPajak_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboFakturPajak.SelectedIndexChanged
        Try
            If Me.CboFakturPajak.SelectedIndex = 0 Then
                'faktur pajak sederhana
                Me.txtNoFakturPajakSederhana.Visible = True
                Me.txtNoFakturPajakSederhana.TabStop = True
                Me.txtNoFakturPajakSederhana.BringToFront()
                Me.txtNoFakturPajakSederhana.BackColor = Me.txtPrimaryInvoice.BackColor
                Me.txtNoSeriFakturPajak.TabStop = False
                'Me.txtNoFakturPajak.TabStop = False '//#2732
            Else
                'faktur pajak standar
                Me.txtNoFakturPajakSederhana.Visible = False
                Me.txtNoFakturPajakSederhana.TabStop = False
                Me.txtNoFakturPajakSederhana.SendToBack()
                Me.txtNoSeriFakturPajak.TabStop = True
                'Me.txtNoFakturPajak.TabStop = False '//#2732

                '2025 Core Tax Tahap 2.2
                ' If Me.baru = False And Me.txtNoFakturPajak.Text = "" Then
                '    If Me.chkTax.Checked = True Then
                '        Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '        Me.llbEditNoFakturPajak.Enabled = True
                '    Else
                '        Me.llbEditNoFakturPajak.Enabled = False
                '        Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '        If Me.txtNoFakturPajak.Text <> "" Then
                '            Me.txtNoFakturPajak.Text = Me.txtNoFakturPajak.Text.Substring(0, 8) + "00000000"
                '        End If
                '    End If
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtDutyStamp.ReadOnly = True Then
                    Me.txtDutyStamp.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtDutyStamp.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDutyStamp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtDutyStamp.ReadOnly = True Then
                    Me.txtDutyStamp.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtDutyStamp.BackColor = Me.txtPrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajakSederhana_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajakSederhana.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoFakturPajakSederhana.ReadOnly = True Then
                    Me.txtNoFakturPajakSederhana.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtNoFakturPajakSederhana.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajakSederhana_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajakSederhana.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoFakturPajakSederhana.ReadOnly = False Then
                    Me.txtNoFakturPajakSederhana.BackColor = Me.txtPrimaryCurrency.BackColor
                Else
                    Me.txtNoFakturPajakSederhana.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNamaSigned_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNamaSigned.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtNamaSigned.ReadOnly = True Then
                    Me.TxtNamaSigned.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtNamaSigned.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNamaSigned_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNamaSigned.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtNamaSigned.ReadOnly = True Then
                    Me.TxtNamaSigned.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtNamaSigned.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJabatanSigned_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJabatanSigned.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJabatanSigned.ReadOnly = True Then
                    Me.TxtJabatanSigned.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtJabatanSigned.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJabatanSigned_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJabatanSigned.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtJabatanSigned.ReadOnly = True Then
                    Me.TxtJabatanSigned.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtJabatanSigned.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtTax.ReadOnly = True Then
                    Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTax.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtTax.ReadOnly = True Then
                    Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTax.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = True Then
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTaxExchangeRate.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtTaxExchangeRate.ReadOnly = True Then
                    Me.txtTaxExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtTaxExchangeRate.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbEditNoFakturPajak_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditNoFakturPajak.LinkClicked
        Try
            If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso CDec(Me.txtPrimaryInvoice.Text) > 0) Then
                Me.ShowDetailFakturPajak()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowDetailFakturPajak()
        Dim i As New FrmTransaksiSalesInvoice_EditNo
        Dim j As New FrmTransaksiSalesInvoice_EditNo2
        Dim Dummy As String = ""
        Dim PrimaryInvoice As Double = -1
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim StartingNumber As Double = 0
        Dim EndingNumber As Double = 0

        Try
            If Me.dtpDate.Value < Obj.BatasStartDate Then
                i.mtb_Number.Mask = "00000000"
                i.mtb_Number.Text = Me.txtNoFakturPajak.Text.Substring(3)
                i.MinLength = 8
                i.Text = "Edit Tax Invoice Number"
                If Me.TxtRemarks.Enabled = False Then
                    i.mtb_Number.Enabled = False
                    i.TxtRevision.Enabled = False
                End If
                If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dummy = Me.txtNoFakturPajak.Text.Substring(0, 3) + i.mtb_Number.Text
                    PrimaryInvoice = Me.ObjInvoice.GetPrimaryInvoiceByNoFakturPajak(Dummy)
                    If PrimaryInvoice = -1 Then
                        Me.txtNoFakturPajak.Text = Dummy
                    Else
                        If CDbl(Me.txtPrimaryInvoice.Text) = PrimaryInvoice Then
                            Me.txtNoFakturPajak.Text = Dummy
                        Else
                            MsgBox("The tax invoice number " + Dummy + " is used by " + _
                            Me.ObjInvoice.GetNoInvoice(PrimaryInvoice) + ".", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            Else
                j.mtb_Number.Mask = "00000000"
                If Me.txtNoFakturPajak.Text.Length > 7 Then
                    j.mtb_Number.Text = Me.txtNoFakturPajak.Text.Substring(8)
                Else
                    j.mtb_Number.Text = ""
                End If
                j.PrimaryInvoice = Me.txtPrimaryInvoice.Text
                j.InvDate = Me.dtpDate.Value
                j.primaryTax = Me.ObjTaxInvoiceSetting.GetPrimary(Me.dtpDate.Value, Me.txtNoFakturPajak.Text)
                j.MinLength = 8
                j.Text = "Edit Tax Invoice Number"

                If Me.TxtRemarks.Enabled = False Then
                    j.cboJenisFaktur.Enabled = False
                    j.mtb_Number.Enabled = False
                    j.Txt_E_Faktur_Blok.Enabled = False
                    j.Txt_E_Faktur_Jalan.Enabled = False
                    j.Txt_E_Faktur_Kabupaten.Enabled = False
                    j.Txt_E_Faktur_Kecamatan.Enabled = False
                    j.Txt_E_Faktur_Kelurahan.Enabled = False
                    j.Txt_E_Faktur_KodePos.Enabled = False
                    j.Txt_E_Faktur_No.Enabled = False
                    j.Txt_E_Faktur_NoTelepon.Enabled = False
                    j.Txt_E_Faktur_Propinsi.Enabled = False
                    j.Txt_E_Faktur_RT.Enabled = False
                    j.Txt_E_Faktur_RW.Enabled = False
                End If

                j.Txt_E_Faktur_Blok.Text = Me.Txt_E_Faktur_Blok.Text
                j.Txt_E_Faktur_Jalan.Text = Me.Txt_E_Faktur_Jalan.Text
                j.Txt_E_Faktur_Kabupaten.Text = Me.Txt_E_Faktur_Kabupaten.Text
                j.Txt_E_Faktur_Kecamatan.Text = Me.Txt_E_Faktur_Kecamatan.Text
                j.Txt_E_Faktur_Kelurahan.Text = Me.Txt_E_Faktur_Kelurahan.Text
                j.Txt_E_Faktur_KodePos.Text = Me.Txt_E_Faktur_KodePos.Text
                j.Txt_E_Faktur_No.Text = Me.Txt_E_Faktur_No.Text
                j.Txt_E_Faktur_NoTelepon.Text = Me.Txt_E_Faktur_NoTelepon.Text
                j.Txt_E_Faktur_Propinsi.Text = Me.Txt_E_Faktur_Propinsi.Text
                j.Txt_E_Faktur_RT.Text = Me.Txt_E_Faktur_RT.Text
                j.Txt_E_Faktur_RW.Text = Me.Txt_E_Faktur_RW.Text

                If j.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dummy = j.Result
                    Me.txtNoFakturPajak.Text = Dummy

                    Me.Txt_E_Faktur_Blok.Text = j.Txt_E_Faktur_Blok.Text
                    Me.Txt_E_Faktur_Jalan.Text = j.Txt_E_Faktur_Jalan.Text
                    Me.Txt_E_Faktur_Kabupaten.Text = j.Txt_E_Faktur_Kabupaten.Text
                    Me.Txt_E_Faktur_Kecamatan.Text = j.Txt_E_Faktur_Kecamatan.Text
                    Me.Txt_E_Faktur_Kelurahan.Text = j.Txt_E_Faktur_Kelurahan.Text
                    Me.Txt_E_Faktur_KodePos.Text = j.Txt_E_Faktur_KodePos.Text
                    Me.Txt_E_Faktur_No.Text = j.Txt_E_Faktur_No.Text
                    Me.Txt_E_Faktur_NoTelepon.Text = j.Txt_E_Faktur_NoTelepon.Text
                    Me.Txt_E_Faktur_Propinsi.Text = j.Txt_E_Faktur_Propinsi.Text
                    Me.Txt_E_Faktur_RT.Text = j.Txt_E_Faktur_RT.Text
                    Me.Txt_E_Faktur_RW.Text = j.Txt_E_Faktur_RW.Text
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    'Private Sub llbEditNoInvoice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditNoInvoice.LinkClicked
    '    Dim i As New FrmTransaksiSalesInvoice_EditNo
    '    Dim Dummy As String = ""
    '    Dim PrimaryInvoice As Double = -1
    '    Dim JmlRujukan As Integer = 0
    '    Dim JmlRujukan1 As Integer = 0
    '    Dim JmlRujukan2 As Integer = 0
    '    Dim Dummy1 As String = ""

    '    Try
    '        If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso CDec(Me.txtPrimaryInvoice.Text) > 0) Then

    '            i.BackColorAll = Me.BackColorAll
    '            JmlRujukan = Me.ObjNumbering.GetJmlRunningNumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice)
    '            While Dummy.Length < JmlRujukan
    '                Dummy = Dummy + "0"
    '            End While
    '            i.mtb_Number.Mask = Dummy
    '            i.MinLength = JmlRujukan
    '            JmlRujukan1 = Me.ObjNumbering.GetUrutan(Me.txtInvoiceNo.Text, ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, Me.dtpDate.Value)

    '            i.mtb_Number.Text = JmlRujukan1
    '            While i.mtb_Number.Text.Length < JmlRujukan
    '                i.mtb_Number.Text = "0" + i.mtb_Number.Text
    '            End While
    '            Dummy1 = i.mtb_Number.Text

    '            i.Label1.Text = "Invoice Number"
    '            i.Text = "Edit Invoice Number"

    '            'cek revisi
    '            JmlRujukan2 = Me.ObjInvoice.GetNoBaru1(Me.dtpDate.Value).Length
    '            If JmlRujukan2 <= Me.txtInvoiceNo.Text.Length Then
    '                i.TxtRevision.Text = Me.txtInvoiceNo.Text.Substring(JmlRujukan2)
    '            Else
    '                i.TxtRevision.Text = ""
    '            End If

    '            If Me.TxtRemarks.Enabled = False Then
    '                i.mtb_Number.Enabled = False
    '                i.TxtRevision.Enabled = False
    '            End If

    '            If i.ShowDialog = Windows.Forms.DialogResult.OK Then
    '                Dummy = Me.txtInvoiceNo.Text.Replace(Dummy1, i.mtb_Number.Text)
    '                Dummy = Dummy.Substring(0, JmlRujukan2) + i.TxtRevision.Text

    '                PrimaryInvoice = Me.ObjInvoice.GetPrimaryInvoice(Dummy)
    '                If PrimaryInvoice = -1 Then
    '                    Me.txtInvoiceNo.Text = Dummy
    '                Else
    '                    If CDbl(Me.txtPrimaryInvoice.Text) = PrimaryInvoice Then
    '                        Me.txtInvoiceNo.Text = Dummy
    '                    Else
    '                        MsgBox("The invoice number " + Dummy + " is used. Please enter another number.", _
    '                        MsgBoxStyle.Information, "Information")
    '                    End If
    '                End If
    '            End If
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        i.Dispose()
    '    End Try
    'End Sub

    Private Sub llbEditNoInvoice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbEditNoInvoice.LinkClicked
        Dim i As New FrmTransaksiSalesInvoice_EditNo
        Dim Dummy As String = ""
        Dim PrimaryInvoice As Double = -1
        Dim JmlRujukan As Integer = 0
        Dim JmlRujukan1 As Integer = 0
        Dim JmlRujukan2 As Integer = 0
        Dim Dummy1 As String = ""

        Try
            If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso CDec(Me.txtPrimaryInvoice.Text) > 0) Then

                i.BackColorAll = Me.BackColorAll
                JmlRujukan = Me.ObjNumbering.GetJmlRunningNumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice)
                While Dummy.Length < JmlRujukan
                    Dummy = Dummy + "0"
                End While
                i.mtb_Number.Mask = Dummy
                i.MinLength = JmlRujukan
                JmlRujukan1 = Me.ObjNumbering.GetUrutan(Me.txtInvoiceNo.Text, ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, Me.dtpDate.Value)

                i.mtb_Number.Text = JmlRujukan1
                While i.mtb_Number.Text.Length < JmlRujukan
                    i.mtb_Number.Text = "0" + i.mtb_Number.Text
                End While
                Dummy1 = i.mtb_Number.Text

                i.Label1.Text = "Invoice Number"
                i.Text = "Edit Invoice Number"

                'cek revisi
                JmlRujukan2 = Me.ObjInvoice.GetNoBaru1(Me.dtpDate.Value).Length
                If JmlRujukan2 <= Me.txtInvoiceNo.Text.Length Then
                    i.TxtRevision.Text = Me.txtInvoiceNo.Text.Substring(JmlRujukan2)
                Else
                    i.TxtRevision.Text = ""
                End If

                If Me.TxtRemarks.Enabled = False Then
                    i.mtb_Number.Enabled = False
                    i.TxtRevision.Enabled = False
                End If

                'update ory
                i.TanggalTransaksi = dtpDate.Value
                i.exPrimary = CDbl(txtPrimaryInvoice.Text)

                If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dummy = Me.txtInvoiceNo.Text.Replace(Dummy1, i.mtb_Number.Text)
                    Dummy = Dummy.Substring(0, JmlRujukan2) + i.TxtRevision.Text

                    PrimaryInvoice = Me.ObjInvoice.GetPrimaryInvoice(Dummy)
                    If PrimaryInvoice = -1 Then
                        Me.txtInvoiceNo.Text = Dummy
                    Else
                        If CDbl(Me.txtPrimaryInvoice.Text) = PrimaryInvoice Then
                            Me.txtInvoiceNo.Text = Dummy
                        Else
                            MsgBox("The invoice number " + Dummy + " is used. Please enter another number.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            If Me.TxtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotalBeforeDP.Text) Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2)
                Me.HitungVATItem()
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTaxExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Leave
        Try
            If CDec(Me.txtTaxExchangeRate.Text) = 0 Then
                Me.txtTaxExchangeRate.Text = 1
            End If
            Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.txtTaxExchangeRate.Text, 2)

            If Me.TxtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotalBeforeDP.Text) Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajak_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajak.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoFakturPajak.ReadOnly = True Then
                    Me.txtNoFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtNoFakturPajak.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoFakturPajak_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoFakturPajak.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoFakturPajak.ReadOnly = True Then
                    Me.txtNoFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtNoFakturPajak.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUpDateNPWP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUpDateNPWP.LinkClicked
        Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
        Try
            If Me.txtPrimaryCustomer.Text <> "" And Me.txtPrimaryCustomer.Text <> "-1" Then
                If Me.cboNIKPASS.Visible = True Then
                    Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass2(Me.txtPrimaryCustomer.Text, Me.OptionNIKPASS)
                Else
                    If MsgBox("Update NPWP number from customer data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                        If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        Else
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                        End If
                        Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, Me.ObjInvoice.AlamatPajak1, Me.ObjInvoice.AlamatPajak2)
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0
        Dim UnitPrice, Diskon As Decimal

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                    'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                    'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                    'Me.TxtDiscFromPerct.Text = "1"

                    If Me.cboBatchNo.Visible = True Then
                        Me.cboBatchNo.Focus()
                    ElseIf Me.cboLocation.Visible = True Then
                        Me.TxtQtyDefault.Focus()
                    Else
                        If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                            Me.TxtUnitPrice.Focus()
                        Else
                            Me.TxtCurrencyCode1.Focus()
                        End If
                    End If
                    Me.Hitung_Amount2()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.cboBatchNo.Visible = True Then
                            Me.cboBatchNo.Focus()
                        ElseIf Me.cboLocation.Visible = True Then
                            Me.cboLocation.Focus()
                        Else
                            If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                Me.TxtQtyDefault.Focus()
                            Else
                                Me.TxtCurrencyCode1.Focus()
                            End If
                        End If
                        'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                        'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                        'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                        'Me.TxtDiscFromPerct.Text = "1"
                        Me.Hitung_Amount2()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct
        Dim UnitPrice, Diskon As Decimal

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.SetGboItem()
                    If Me.cboBatchNo.Visible = True Then
                        Me.cboBatchNo.Focus()
                    ElseIf Me.cboLocation.Visible = True Then
                        Me.cboLocation.Focus()
                    Else
                        If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                            Me.TxtQtyDefault.Focus()
                        Else
                            Me.TxtCurrencyCode1.Focus()
                        End If
                    End If
                    'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                    'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                    'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                    'Me.TxtDiscFromPerct.Text = "1"
                    Me.Hitung_Amount2()
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.SetGboItem()
                        If Me.cboBatchNo.Visible = True Then
                            Me.cboBatchNo.Focus()
                        ElseIf Me.cboLocation.Visible = True Then
                            Me.cboLocation.Focus()
                        Else
                            If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                                Me.TxtQtyDefault.Focus()
                            Else
                                Me.TxtCurrencyCode1.Focus()
                            End If
                        End If
                        'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                        'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                        'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                        'Me.TxtDiscFromPerct.Text = "1"
                        Me.Hitung_Amount2()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct
        'Dim UnitPrice, Diskon As Decimal

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.BackColorAll = Me.BackColorAll
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.SetGboItem()
                If Me.cboBatchNo.Visible = True Then
                    Me.cboBatchNo.Focus()
                ElseIf Me.cboLocation.Visible = True Then
                    Me.cboLocation.Focus()
                Else
                    If Me.TxtPrimaryCurrency1.Text <> "-1" Then
                        Me.TxtQtyDefault.Focus()
                    Else
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If

                'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text))
                'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                'Me.TxtDiscFromPerct.Text = "1"
                Me.Hitung_Amount2()
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetGboItem()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNo.Text, Me.chkLocationControlled, _
            Me.chkBatchNoControlled, Me.chkExpiryDateControlled, Me.TxtPrimaryUnitsDefault)

            Me.cboBatchNo.DataSource = Nothing
            Me.cboLocation.DataSource = Nothing

            If Me.chkBatchNoControlled.Checked = True Then
                Me.cboBatchNo.TabStop = True
                Me.lblBatchNo.Visible = False
                Me.cboBatchNo.Visible = True
                'isikan data batch no dari basis data
                Me.IsiCboItem(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryInvoice.Text)
            Else
                Me.cboBatchNo.TabStop = False
                Me.lblBatchNo.Visible = True
                Me.cboBatchNo.Visible = False
                Me.cboBatchNo.Items.Clear()
            End If

            If Me.chkLocationControlled.Checked = False Then
                Me.txtPrimaryLocation.Text = "-1"
                Me.cboLocation.Items.Clear()
                Me.lblLocation.Visible = True
                Me.cboLocation.Visible = False
            Else
                Me.lblLocation.Visible = False
                Me.cboLocation.Visible = True
                If Me.chkLocationControlled.Checked = True Then
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryInvoice.Text, Me.cboBatchNo.Text)
                End If
            End If

            If Me.chkExpiryDateControlled.Checked = True Then
                Me.lblExpiryDate.Visible = False
                Me.DtpExpired.Visible = True
            Else
                Me.lblExpiryDate.Visible = True
                Me.DtpExpired.Visible = False
                Me.DtpExpired.Value = #1/1/1950#
            End If

            If Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Me.lblUnitsDefault.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
            Else
                Me.lblUnitsDefault.Text = "Units"
            End If

            Me.TxtQtyDefault.Text = "0"
            Me.TxtQtyMulti.Text = "0"

            Me.SetCboPriceByMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCboPriceByMultiUnit()
        Dim UnitPrice, Discount As Decimal
        Dim UnitCodeDummy As String = ""

        Try
            Me.cboPriceByMultiUnit.Items.Clear()

            If Me.CboUnit.Items.Count = 1 Then
                Me.LblUnitPrice.Text = "Unit Price / " + Me.CboUnit.Text
                Me.cboPriceByMultiUnit.Visible = False
                Me.cboPriceByMultiUnit.TabStop = False
                Me.cboPriceByMultiUnit.Items.Add(Me.lblUnitsDefault.Text)
                Me.cboPriceByMultiUnit.SelectedIndex = 0
                Me.txtPriceByMultiUnit.Text = "1"
            Else
                Me.LblUnitPrice.Text = "Unit Price"
                Me.cboPriceByMultiUnit.Visible = True
                Me.cboPriceByMultiUnit.TabStop = True
                Me.cboPriceByMultiUnit.Items.Add(Me.lblUnitsDefault.Text)
                UnitCodeDummy = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.CboUnit.Text <> UnitCodeDummy Then
                    Me.cboPriceByMultiUnit.Items.Add(Me.CboUnit.Text)
                    Me.cboPriceByMultiUnit.SelectedIndex = Me.ObjProduct.GetPriceByMultiUnit(Me.TxtPrimaryCodeNo.Text)
                Else
                    Me.cboPriceByMultiUnit.SelectedIndex = 0
                End If
                If Me.cboPriceByMultiUnit.Text <> UnitCodeDummy Then
                    Me.txtPriceByMultiUnit.Text = "1"
                Else
                    Me.txtPriceByMultiUnit.Text = "0"
                End If
                Me.cboPriceByMultiUnit.Left = Me.LblUnitPrice.Left + Me.LblUnitPrice.Width
            End If

            If Me.txtPrimaryCurrency.Text.Trim = "" Then
                Me.txtPrimaryCurrency.Text = "-1"
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                  Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Discount)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                  Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsDefault.Text, UnitPrice, Discount)
            End If
            If UnitPrice = 0 And Discount = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                Else
                    Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
            Me.Hitung_Amount2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboItem(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Try
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNo, Me.TBL_Item)

            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item2Row
            Dim row2 As Core_Module.dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean

            If Me.baru = False Then
                Me.ObjInvoiceItem.IsikanTBLBatchNo(Me.TBL4_Item, vPrimaryProduct, vPrimaryTransaksi)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4_Item.Rows.Count > 0
                    row1 = Me.TBL4_Item.Rows(0)
                    For Each row2 In Me.TBL_Item.Rows
                        If row1.Lot_No = row2.Lot_No Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL_Item.Rows.Add(row1.Lot_No, row1.Expiry_Date, row1.Qty)
                    End If
                    Me.TBL4_Item.Rows.Remove(row1)
                End While
            End If

            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLBatchNo(vPrimaryProduct)
            While Me.TBL4_Item.Rows.Count > 0
                row1 = Me.TBL4_Item.Rows(0)
                For Each row2 In Me.TBL_Item.Rows
                    If row1.Lot_No = row2.Lot_No Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL_Item.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL_Item.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4_Item.Rows.Remove(row1)
            End While

            If Me.chkExpiryDateControlled.Checked = False Then
                Me.TBL_Item.DefaultView.Sort = "[Lot No] ASC"
            Else
                Me.TBL_Item.DefaultView.Sort = "[Expiry Date] ASC"
            End If

            Me.cboBatchNo.DataSource = Me.TBL_Item.DefaultView
            Me.cboBatchNo.ValueMember = "Lot No"
            Me.cboBatchNo.DisplayMember = "Lot No"
            Me.cboBatchNo.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Double)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item2Row

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4_Item.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColLotNumber.Index).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next

                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(Me.ColExpiry.Index).Value) = True Then
                            ED = row.Cells(Me.ColExpiry.Index).Value
                        End If
                        Me.TBL4_Item.Rows.Add(CStr(row.Cells(Me.ColLotNumber.Index).Value), _
                        CDbl(row.Cells(Me.ColTotalQty.Index).Value), ED)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboLokasi(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboLocation(vPrimaryProduct, Me.cboLocation, _
            Me.TBL2_Item, vBatchNo)
            Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item3Row
            Dim row2 As Core_Module.dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.ObjInvoiceItem.IsikanTBLLocation(Me.TBL5_Item, vPrimaryProduct, vPrimaryTransaksi, vBatchNo)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL5_Item.Rows.Count > 0
                    row1 = Me.TBL5_Item.Rows(0)
                    For Each row2 In Me.TBL2_Item.Rows
                        If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL2_Item.Rows.Add(row1.Lot_No, row1.Location, row1.Qty)
                    End If
                    Me.TBL5_Item.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLLocation(vPrimaryProduct, vBatchNo)
            While Me.TBL5_Item.Rows.Count > 0
                row1 = Me.TBL5_Item.Rows(0)
                For Each row2 In Me.TBL2_Item.Rows
                    If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL2_Item.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL2_Item.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL5_Item.Rows.Remove(row1)
            End While

            Me.TBL2_Item.DefaultView.Sort = "Location ASC"
            Me.cboLocation.DataSource = Me.TBL2_Item.DefaultView
            Me.cboLocation.ValueMember = "Location"
            Me.cboLocation.DisplayMember = "Location"
            Me.cboLocation.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLLocation(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item3Row
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                CStr(row.Cells(Me.ColLotNumber.Index).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL5_Item.Rows
                        If row1.Lot_No = CStr(row.Cells(Me.ColLotNumber.Index).Value) _
                        And row1.Location = row.Cells(Me.ColLocationCode.Index).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(Me.ColTotalQty.Index).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next

                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL5_Item.Rows.Add(CStr(row.Cells(Me.ColLotNumber.Index).Value), _
                        row.Cells(Me.ColLocationCode.Index).Value, _
                        CDbl(row.Cells(Me.ColTotalQty.Index).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan index berdasarkan unit price, discount, batch number, primary location, primary product
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetIndex(ByVal vUnitPrice As Decimal, ByVal vBatchNumber As String, ByVal vDiscount As Decimal, _
    ByVal vPrimaryLocation As Double, ByVal vPrimaryProduct As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColUnitPrice.Index).Value = vUnitPrice And vBatchNumber = row.Cells(Me.ColLotNumber.Index).Value _
                            And vDiscount = row.Cells(Me.ColDiscount.Index).Value And vPrimaryLocation = row.Cells(Me.ColPrimaryLocation.Index).Value And _
                            vPrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim rowDGV As DataGridViewRow
        Dim QtyGudang, QtyDGV, QtyDatabaseItem, JmlMaks, QtyPending As Decimal
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim AngkaDecimal As Integer = 0
        Dim ObjInvCustom As New CustomSMESS.ClsTransaksiSalesInvoice

        Try
            'data salah jika:
            '1. tidak diisikan, kode product, currency
            'cek data yang harus diisi
            KoleksiNamaData.Add("code no")
            KoleksiTextBox.Add(Me.TxtCodeNo)
            KoleksiNamaData.Add("currency code")
            KoleksiTextBox.Add(Me.TxtCurrencyCode1)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '2. jika sudah ada item, maka kode currencynya harus sama
            If Me.dgvItem.Rows.Count > 0 Then
                rowDGV = Me.dgvItem.Rows(0)
                If Me.TxtPrimaryCurrency1.Text <> rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value Then
                    MsgBox("Currency must be identical for each item", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '3. berkaitan dengan data angka:
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            '   a. unit price, qty harus positif
            KoleksiNamaData.Add("Unit Price")
            KoleksiTextBox.Add(Me.TxtUnitPrice)
            KoleksiNamaData.Add("Qty")
            KoleksiTextBox.Add(Me.TxtQtyMulti)
            If ModCekData.CekBilanganPositif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            '   3.b. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyDefault.Text) = False Then
                MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyMulti.Focus()
                Return False
            End If
            '   b. discount paling besar 100% --> otomatis + non negatif
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()
            KoleksiNamaData.Add("discount")
            KoleksiTextBox.Add(Me.TxtDiscount)
            If ModCekData.CekNonNegatif(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            '4. Remarks maksimal 150 karakter
            If Me.TxtRemarkItem.Text.Length > 150 Then
                MsgBox("Maximum remarks data length is 150 character", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan data stock tidak minus
            'lakukan pengecekan
            If Me.ObjProduct.ApakahGoods(Me.TxtPrimaryCodeNo.Text) = True Then
                QtyGudang = Me.ObjProduct.GetJml2(Me.TxtPrimaryCodeNo.Text, DTA1, -1, _
                            Me.txtPrimaryLocation.Text, Me.cboBatchNo.Text)
                QtyDGV = Me.GetJmlDGV_Default(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text, _
                Me.txtPrimaryLocation.Text)
                QtyDatabaseItem = Me.ObjInvoiceItem.GetJml_Default(Me.txtPrimaryInvoice.Text, Me.TxtPrimaryCodeNo.Text, _
                Me.cboBatchNo.Text, Me.txtPrimaryLocation.Text)
                AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                QtyPending = Me.objDO.GetJmlPending_Default(Me.TxtPrimaryCodeNo.Text)

                JmlMaks = QtyGudang + QtyDatabaseItem - QtyDGV - QtyPending
                If JmlMaks < 0 Then
                    JmlMaks = 0
                End If

                If CDec(Me.TxtQtyDefault.Text) > JmlMaks Then
                    'Dim Max_ As String = ObjCD.DigitGrouping(JmlMaks * Me.ObjSO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text) / Me.ObjSO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text), ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                    Dim Max_ As String = JmlMaks * Me.ObjSO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text) / Me.ObjSO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
                    Max_ = Me.ObjCD.RoundDown1(Max_, ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))

                    MsgBox("Maximum qty data is " & Max_ & " " + Me.ObjUnit.GetUnitCode(Me.TxtPrimaryUnitsMulti.Text), _
                    MsgBoxStyle.Information, "Information")
                    Me.TxtQtyMulti.Focus()
                    Return False
                End If
            End If

            'cek batasan custom
            If ObjInvCustom.ApakahDataBenarItem(Me.TxtPrimaryCodeNo.Text, Me.txtPriceByMultiUnit.Text, _
            Me.TxtPrimaryUnitsMulti.Text, Me.TxtPrimaryUnitsDefault.Text) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            DTA1.Dispose()
            ObjInvCustom = Nothing
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vPrimaryProduct As Double, ByVal vBatchNumber As String, ByVal vPrimaryLocation As Double) As Decimal
        Dim row As DataGridViewRow
        Dim retVal As Double = 0

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                row.Cells(Me.ColLotNumber.Index).Value = vBatchNumber And _
                row.Cells(Me.ColPrimaryLocation.Index).Value = vPrimaryLocation Then
                    retVal += CDec(row.Cells(Me.ColTotalQty.Index).Value)
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'mengembalikan nilai index dari baris yang nilai batch number, code no, unit price, discount,
    'location
    Private Function GetIndex() As Integer
        Try
            Return Me.GetIndex(Me.TxtUnitPrice.Text, Me.cboBatchNo.Text, Me.TxtDiscount.Text, _
            Me.txtPrimaryLocation.Text, Me.TxtPrimaryCodeNo.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1

        Try
            If Me.ApakahDataItemBenar = True Then
                indexSama = Me.GetIndex
                Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
                If indexSama = -1 Then
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)
                    AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(Me.TxtPrimaryCurrency1.Text)

                    With row
                        .Cells(Me.ColNo0.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColCustCodeNo.Index).Value = ""
                        .Cells(Me.ColCustDesc.Index).Value = ""
                        .Cells(Me.ColSONO.Index).Value = ""
                        .Cells(Me.ColCustPONo.Index).Value = ""
                        .Cells(Me.ColDONO.Index).Value = ""
                        .Cells(Me.ColStockOutNo.Index).Value = ""
                        .Cells(Me.ColLotNumber.Index).Value = Me.cboBatchNo.Text
                        If Me.DtpExpired.Visible = False Then
                            .Cells(Me.ColExpiry.Index).Value = #1/1/1950#
                        Else
                            .Cells(Me.ColExpiry.Index).Value = Me.DtpExpired.Value
                        End If
                        .Cells(Me.ColLocationCode.Index).Value = Me.cboLocation.Text
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyDefault.Text), AngkaDecimal)
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColSecQty.Index).Value = -1
                        .Cells(Me.ColSecUnit.Index).Value = ""
                        .Cells(Me.ColCurrency.Index).Value = Me.TxtCurrencyCode1.Text
                        .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscount.Index).Value = CDec(Me.TxtDiscount.Text)
                        .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtNettPrice.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(Me.txtAmount.Text), AngkaDecimalCurrency)
                        .Cells(Me.ColRemarksItem.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryLocation.Index).Value = Me.txtPrimaryLocation.Text
                        .Cells(Me.ColPrimaryDO.Index).Value = -1
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimarySO.Index).Value = -1
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(Me.ColPrimaryUnit.Index).Value = Me.TxtPrimaryUnitsDefault.Text
                        .Cells(Me.ColPrimarySecUnit.Index).Value = -1
                        .Cells(Me.ColPrimaryCurrency.Index).Value = Me.TxtPrimaryCurrency1.Text
                        .Cells(Me.ColAliasDesc.Index).Value = Me.ObjInvoiceItem.GetAliasDesc(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCustomer.Text)
                        .Cells(Me.ColAliasCode.Index).Value = Me.ObjInvoiceItem.GetAliasCode(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCustomer.Text)

                        .Cells(Me.ColDiscAmount.Index).Value = Me.TxtDiscAmount1.Text
                        .Cells(Me.ColDiscFromPrect.Index).Value = Me.TxtDiscFromPerct.Text

                        .Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, DefPrecMulti)
                        .Cells(Me.ColUnitsMulti.Index).Value = Me.CboUnit.Text
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.TxtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjSO.GetRatioDefault(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), AngkaDecimal)
                        .Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjSO.GetRatioReference(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecMulti)
                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((CDec(Me.txtNettPrice.Text) * CDec(Me.txtTax.Text)) / 100, AngkaDecimalCurrency)
                        If AngkaDecimalCurrency = 0 Then
                            If Me.txtNettPrice.Text.EndsWith("55") Then
                                .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT.Index).Value) - 1, AngkaDecimalCurrency)
                            End If
                        End If
                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), AngkaDecimalCurrency)
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnitsMulti.Index).Value
                        Else
                            .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(.Cells(Me.ColVAT.Index).Value * CDec(.Cells(Me.ColTotalQty.Index).Value), AngkaDecimalCurrency)
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnit.Index).Value
                        End If
                        .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), AngkaDecimalCurrency)
                        Me.SetRatioOnDgv(row)

                        .Cells(Me.ColPriceByMultiUnit.Index).Value = Me.txtPriceByMultiUnit.Text
                    End With
                Else
                    row = Me.dgvItem.Rows(indexSama)
                    'tambahkan qty multi, 
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(Me.TxtQtyMulti.Text), DefPrecMulti)
                    'hitung ulang amount
                    row = Me.dgvItem.Rows(indexSama)

                    Me.SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, row.Index).Value))

                    Me.Hitung_Amount(row)
                End If

                PrimaryCurrenct = Me.TxtPrimaryCurrency1.Text
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()

                Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
                Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                Me.FormatDGV2()

                Me.cboCalculatesBy.SelectedIndex = -1
                Me.cboCalculatesBy.SelectedIndex = 0

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                If Me.ApakahAdaAlias = True Then
                    Me.chkAlias.Checked = True
                Else
                    Me.chkAlias.Checked = False
                End If

                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV2()
        Dim row As DataGridViewRow
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True
        Dim IsMultiUnit As Boolean = False
        Dim PriceByUnitDefault As Boolean = False

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(Me.ColExpiry.Index).Value) = False Then
                    row.Cells(Me.ColExpiry.Index).Value = "-"
                Else
                    If CDate(row.Cells(Me.ColExpiry.Index).Value).Date = #1/1/1950# Then
                        row.Cells(Me.ColExpiry.Index).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(Me.ColExpiry.Index).Value = Format(CDate(row.Cells(Me.ColExpiry.Index).Value), "dd MMM yyyy")
                    End If
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(Me.ColPrimaryLocation.Index).Value = "-1" Then
                    row.Cells(Me.ColLocationCode.Index).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(Me.ColLotNumber.Index).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    PriceByUnitDefault = True
                    Exit For
                End If
            Next
            Me.ColPriceByMultiUnit_Str.Visible = PriceByUnitDefault

            '2. tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
            Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False

            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                Me.dgvItem.Columns(Me.ColLocationCode.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLocationCode.Index).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                Me.dgvItem.Columns(Me.ColLotNumber.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColLotNumber.Index).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                Me.dgvItem.Columns(Me.ColExpiry.Index).Visible = False
            Else
                Me.dgvItem.Columns(Me.ColExpiry.Index).Visible = True
            End If

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryUnitsMulti.Index).Value <> row.Cells(Me.ColPrimaryUnit.Index).Value Then
                    IsMultiUnit = True
                    Exit For
                End If
            Next
            If IsMultiUnit = False Then
                Me.ColTotalQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColTotalQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True
            End If

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.Buffer2 = Me.TxtUnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e, 9999999999)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            Me.Buffer2 = Me.TxtDiscount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        If Me.TxtRemarks.Enabled = True Then
            If Me.TxtCodeNo.ReadOnly = False And Me.TxtCodeNo.Enabled = True Then
                TxtCodeNo.BackColor = BackColorAll
            Else
                TxtCodeNo.BackColor = Me.txtPrimaryInvoice.BackColor
            End If
        End If
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        If Me.TxtRemarks.Enabled = True Then
            If Me.TxtCodeNo.ReadOnly = False And Me.TxtCodeNo.Enabled = True Then
                TxtCodeNo.BackColor = Me.txtPrimaryInvoice.BackColor
            End If
        End If
    End Sub

    Private Sub TxtCurrencyCode1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.GotFocus
        If Me.TxtCurrencyCode1.ReadOnly = False And Me.TxtCurrencyCode1.Enabled = True Then
            TxtCurrencyCode1.BackColor = BackColorAll
        Else
            TxtCurrencyCode1.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtCurrencyCode1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.LostFocus
        If Me.TxtCurrencyCode1.ReadOnly = False And Me.TxtCurrencyCode1.Enabled = True Then
            TxtCurrencyCode1.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
            TxtUnitPrice.BackColor = BackColorAll
        Else
            TxtUnitPrice.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        If Me.TxtUnitPrice.ReadOnly = False And Me.TxtUnitPrice.Enabled = True Then
            TxtUnitPrice.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.GotFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = BackColorAll
        Else
            TxtDiscount.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtQty_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.GotFocus
        If Me.TxtQtyDefault.ReadOnly = False And Me.TxtQtyDefault.Enabled = True Then
            TxtQtyDefault.BackColor = BackColorAll
        Else
            TxtQtyDefault.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyDefault.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyDefault.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyDefault, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.Leave
        Dim UnitPrice, Diskon As Decimal

        Try
            If Me.TxtRemarkItem.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                End If

                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                      Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                Else
                    Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                      Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsDefault.Text, UnitPrice, Diskon)
                End If 
                If UnitPrice = 0 And Diskon = 0 Then
                    If Me.txtPriceByMultiUnit.Text = "1" Then
                        Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                        UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                    Else
                        Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                        UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                    End If
                End If
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
                Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                Me.TxtDiscFromPerct.Text = "1"
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.LostFocus
        If Me.TxtQtyDefault.ReadOnly = False And Me.TxtQtyDefault.Enabled = True Then
            TxtQtyDefault.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        If Me.TxtRemarkItem.ReadOnly = False And Me.TxtRemarkItem.Enabled = True Then
            TxtRemarkItem.BackColor = BackColorAll
        Else
            TxtRemarkItem.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        If Me.TxtRemarkItem.ReadOnly = False And Me.TxtRemarkItem.Enabled = True Then
            TxtRemarkItem.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, Me.DefaultPrecision)
                If CDec(Me.TxtDiscAmount1.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtUnitPrice.Text = Me.Buffer2
                    Exit Sub
                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)
                If CDec(Me.TxtUnitPrice.Text) = 0 Then
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                If CDec(Me.TxtDiscount.Text) <> CDec(Me.Buffer2) Then
                    Me.TxtDiscFromPerct.Text = "1"
                End If
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNo.SelectedIndexChanged
        Try
            If Me.chkLocationControlled.Checked = True Then
                If Me.cboBatchNo.Text <> "" Then
                    'isikan data lokasi
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryInvoice.Text, Me.cboBatchNo.Text)
                End If
            Else
                'kosongkan data lokasi
                Me.cboLocation.DataSource = Nothing
                Me.cboLocation.Items.Clear()
            End If
            'isikan data expiry date 
            If Me.chkExpiryDateControlled.Checked = True Then
                Me.DtpExpired.Value = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNo.Text, Me.cboBatchNo.Text)
            Else
                If Me.DtpExpired.Visible = True Then
                    Me.DtpExpired.Value = Today
                Else
                    Me.DtpExpired.Value = #1/1/1950#
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCurrencyCode1.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
                    Me.TxtUnitPrice.Focus()
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
                Me.SetCalculateBy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCurrencyCode1.Leave
        Dim i As New FrmCariMasterCurrency

        Try
            If Me.TxtCurrencyCode1.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                    Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
                Else
                    If Me.TxtCurrencyCode1.Text <> "" Then
                        i.txtCari.Text = Me.TxtCurrencyCode1.Text
                    End If

                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                        Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                        Me.TxtUnitPrice.Focus()
                    Else
                        Me.TxtCurrencyCode1.Clear()
                        Me.TxtPrimaryCurrency1.Text = "-1"
                        Me.TxtCurrencyCode1.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCurrency1.Text = "-1"
            End If

            Hitung_Amount2()
            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try

        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCurrency2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCurrency2.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.TxtCurrencyCode1.Text) = True Then
                Me.TxtUnitPrice.Focus()
                Me.TxtPrimaryCurrency1.Text = ObjCurrency.GetPrimary(Me.TxtCurrencyCode1.Text)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.TxtCurrencyCode1.Text <> "" Then
                    i.txtCari.Text = Me.TxtCurrencyCode1.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.TxtPrimaryCurrency1.Text = i.PrimaryCurrency
                    Me.TxtCurrencyCode1.Text = ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                    Me.TxtUnitPrice.Focus()
                Else
                    Me.TxtCurrencyCode1.Clear()
                    Me.TxtCurrencyCode1.Focus()
                End If
            End If

            If Me.TxtRemarkItem.Enabled = True Then
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang boleh diedit hanya qty, unit price, remarks item, dan discount
            If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or _
            e.ColumnIndex = Me.ColRemarksItem.Index Or e.ColumnIndex = Me.ColDiscount.Index Or _
            e.ColumnIndex = Me.ColAliasDesc.Index Or e.ColumnIndex = Me.ColAliasCode.Index Or _
            e.ColumnIndex = Me.ColDiscAmount.Index Then
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                'buang nol di belakang decimal jika edit kolom qty
                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Or _
                e.ColumnIndex = Me.ColDiscAmount.Index Or _
                e.ColumnIndex = Me.ColDiscount.Index Or _
                e.ColumnIndex = Me.ColUnitPrice.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSah(ByVal rowIndex As Integer, Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim row, rowEdit As DataGridViewRow
        Dim JmlDGV, JmlGudang, JmlDatabase, JmlPending As Decimal
        Dim PrimaryCodeNo, PrimaryLocation As Double
        Dim BatchNumber As String = ""
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim AngkaDecimal As Integer = 0

        Try
            rowEdit = Me.dgvItem.Rows(rowIndex)
            'cek qty
            If Me.ObjProduct.ApakahGoods(rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value) = True Then
                BatchNumber = rowEdit.Cells(Me.ColLotNumber.Index).Value
                PrimaryCodeNo = rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value
                PrimaryLocation = rowEdit.Cells(Me.ColPrimaryLocation.Index).Value
                AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(PrimaryCodeNo)

                For Each row In Me.dgvItem.Rows
                    If row.Cells(Me.ColPrimaryLocation.Index).Value = PrimaryLocation And _
                    row.Cells(Me.ColPrimaryCodeNo.Index).Value = PrimaryCodeNo And _
                    row.Cells(Me.ColLotNumber.Index).Value = BatchNumber Then
                        JmlDGV += CDec(row.Cells(Me.ColTotalQty.Index).Value)
                    End If
                Next
                JmlGudang = Me.ObjProduct.GetJml2(PrimaryCodeNo, DTA1, -1, PrimaryLocation, BatchNumber)
                JmlDatabase = Me.ObjInvoiceItem.GetJml_Default(Me.txtPrimaryInvoice.Text, PrimaryCodeNo, BatchNumber, PrimaryLocation)
                JmlPending = Me.objDO.GetJmlPending_Default(PrimaryCodeNo)



                If JmlGudang + JmlDatabase - JmlDGV - JmlPending < 0 Then
                    If Me.ObjCD.DigitGrouping((CDec(rowEdit.Cells(Me.ColTotalQty.Index).Value) + _
                    (JmlGudang + JmlDatabase - JmlDGV - JmlPending)), AngkaDecimal) < 0 Then
                        If vShowMsg = True Then
                            MsgBox("Maximum qty you can enter is: " + Me.buffer + " " + rowEdit.Cells(Me.ColUnitsMulti.Index).Value.ToString)
                        End If
                    Else
                        Dim MaxDefault As Decimal = CDec(rowEdit.Cells(Me.ColTotalQty.Index).Value) + (JmlGudang + JmlDatabase - JmlDGV - JmlPending)
                        'Dim MaxMulti As String = Me.ObjCD.DigitGrouping(MaxDefault * CDec(rowEdit.Cells(Me.ColRatioReference.Index).Value) / CDec(rowEdit.Cells(Me.ColRatioDefault.Index).Value), ObjProduct.GetAngkaDecimal(rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value, rowEdit.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
                        Dim MaxMulti As String = MaxDefault * CDec(rowEdit.Cells(Me.ColRatioReference.Index).Value) / CDec(rowEdit.Cells(Me.ColRatioDefault.Index).Value)
                        MaxMulti = ObjCD.RoundDown1(MaxMulti, ObjProduct.GetAngkaDecimal(rowEdit.Cells(Me.ColPrimaryCodeNo.Index).Value, rowEdit.Cells(Me.ColPrimaryUnitsMulti.Index).Value))

                        If vShowMsg = True Then
                            MsgBox("Maximum qty you can enter is: " + MaxMulti + " " + rowEdit.Cells(Me.ColUnitsMulti.Index).Value.ToString)
                        End If
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
        End Try
    End Function

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim row As DataGridViewRow
        Dim AngkaDecimal, IndexTujuan As Integer

        Try
            row = Me.dgvItem.Rows(e.RowIndex)
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'bila tidak ada perubahan maka tidak perlu lakukan apa2
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                        Exit Sub
                    End If

                    'bila ada perubahan:
                    If Me.ColTotalQtyMulti.Index = e.ColumnIndex Or Me.ColUnitPrice.Index = e.ColumnIndex Or _
                    Me.ColDiscount.Index = e.ColumnIndex Or Me.ColDiscAmount.Index = e.ColumnIndex Then
                        If e.ColumnIndex = Me.ColDiscount.Index Then
                            'pastikan nilai discount non negatif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                            'pastikan nilai diskon maksimum = 100
                            If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) >= 100 Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            If CDec(Me.buffer) = CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            row.Cells(Me.ColDiscFromPrect.Index).Value = "1"
                        ElseIf e.ColumnIndex = Me.ColDiscAmount.Index Then
                            'untuk nilai disc amount, pastikan nilainya positif
                            If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for discount amount.", MsgBoxStyle.Information, _
                                "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            'digit grouping
                            AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                            Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                            Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)

                            If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                                'jika unit price < nett price maka kembalikan ke nilai buffer
                                'hanya berlaku klo menentukan diskon berdasarkan amount
                                If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                                'tandai diskon dihitung berdasarkan amount
                                row.Cells(Me.ColDiscFromPrect.Index).Value = "0"
                                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                                               ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)), 2))
                            Else
                                Exit Sub
                            End If
                        ElseIf e.ColumnIndex = Me.ColTotalQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Then
                            If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 15 Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            End If
                            If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            Else
                                'pastikan nilai angka tidak lebih dari 9 M
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 9999999999 Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                            'pastikan nilai unit price dan qty positif
                            If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Exit Sub
                            Else
                                If e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
                                Else
                                    AngkaDecimal = Me.ObjCurrency.GetDefaultPrecission(Me.dgvItem.Item(Me.ColPrimaryCurrency.Index, e.RowIndex).Value)
                                End If
                                'digit grouping
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                                Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, AngkaDecimal)
                            End If
                        End If

                        'jika qty maka cek berkaitan dengan stock
                        If e.ColumnIndex = Me.ColTotalQtyMulti.Index Then
                            Me.SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                            If CDec(row.Cells(Me.ColTotalQty.Index).Value) <= 0 Then
                                MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                                Exit Sub
                            End If
                            If Me.ApakahEditSah(e.RowIndex) = False Then
                                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                Me.SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                                Exit Sub
                            End If
                        End If

                        'jika unit price, pastikan disc amount < unit price
                        If e.ColumnIndex = Me.ColUnitPrice.Index Then
                            If Me.dgvItem.Item(Me.ColDiscFromPrect.Index, e.RowIndex).Value = "0" Then
                                If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > _
                                CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If

                        IndexTujuan = Me.GetIndexSama(e.RowIndex)
                        If IndexTujuan <> -1 Then
                            'gabungkan baris
                            Me.GabungkanBaris(e.RowIndex, IndexTujuan)
                        Else
                            'jika lewat semua maka hitung amount
                            Me.Hitung_Amount(Me.dgvItem.Rows(e.RowIndex))
                        End If
                        'hitung ulang total
                        Me.Hitung_Total()
                    ElseIf e.ColumnIndex = Me.ColRemarksItem.Index Or e.ColumnIndex = Me.ColAliasDesc.Index Or _
                    e.ColumnIndex = Me.ColAliasCode.Index Then
                        'jika kolom remarks maka cek, length maksimum = 150
                        If CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Length > 150 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = _
                            CStr(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Substring(0, 150)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemarksItem.Index Or _
                    e.ColumnIndex = Me.ColAliasDesc.Index Or _
                    e.ColumnIndex = Me.ColAliasCode.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' menggabungkan 2 baris yang kemudian mengembalikan nilai baris hasil penggabungan.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GabungkanBaris(ByVal vIndexAsal As Integer, ByVal vIndexTujuan As Integer) As Integer
        Dim rowAsal, rowTujuan As DataGridViewRow

        Try
            rowAsal = Me.dgvItem.Rows(vIndexAsal)
            rowTujuan = Me.dgvItem.Rows(vIndexTujuan)
            'tambahakan qty pada rowtujuan
            rowTujuan.Cells(Me.ColTotalQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowAsal.Cells(Me.ColTotalQtyMulti.Index).Value) + _
            CDec(rowTujuan.Cells(Me.ColTotalQtyMulti.Index).Value), Me.ObjProduct.GetAngkaDecimal(rowTujuan.Cells(Me.ColPrimaryCodeNo.Index).Value, rowTujuan.Cells(Me.ColPrimaryUnitsMulti.Index).Value))
            SetQtyDefaultOnDgv1(rowTujuan, Me.ObjProduct.GetAngkaDecimal(rowTujuan.Cells(Me.ColPrimaryCodeNo.Index).Value))
            'hitung ulang amount pada row tujuan
            Me.Hitung_Amount(rowTujuan)
            'buang baris row asal
            Me.dgvItem.Rows.Remove(rowAsal)
            'set ulang nomor
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.ObjCD.SetNumber(Me.dgvItem)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'kembalikan index rowtujuan
            Return rowTujuan.Index
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' memberikan nilai true jika baris yang diberikan indexnya harus melakukan penggabungan baris.
    ''' 2 baris digabungkan bila field berikut ini sama:
    ''' a. primary code number
    ''' b. unit price
    ''' c. discount
    ''' d. primary location
    ''' e. batch number
    ''' f. primary unit
    ''' g. primary currency
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndexSama(ByVal rowIndexAcuan As Integer) As Integer
        Dim rowDGV, rowAcuan As DataGridViewRow

        Try
            rowAcuan = Me.dgvItem.Rows(rowIndexAcuan)
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index <> rowIndexAcuan Then
                    With rowAcuan
                        If .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value And _
                        .Cells(Me.ColUnitPrice.Index).Value = rowDGV.Cells(Me.ColUnitPrice.Index).Value And _
                        .Cells(Me.ColDiscount.Index).Value = rowDGV.Cells(Me.ColDiscount.Index).Value And _
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value And _
                        .Cells(Me.ColLotNumber.Index).Value = rowDGV.Cells(Me.ColLotNumber.Index).Value And _
                        .Cells(Me.ColPrimaryUnit.Index).Value = rowDGV.Cells(Me.ColPrimaryUnit.Index).Value And _
                        .Cells(Me.ColPrimaryCurrency.Index).Value = rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value Then
                            Return rowDGV.Index
                        End If
                    End With
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtDiscount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If Me.TxtRemarkItem.Enabled = True Then
                If IsNumeric(Me.TxtDiscount.Text) = True And IsNumeric(Me.TxtUnitPrice.Text) = True And IsNumeric(Me.TxtQtyDefault) = True Then
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If Me.TxtRemarkItem.Enabled = True Then
                If IsNumeric(Me.TxtDiscount.Text) = True And IsNumeric(Me.TxtUnitPrice.Text) = True And IsNumeric(Me.TxtQtyDefault) = True Then
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyDefault.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If Me.TxtRemarkItem.Enabled = True Then
                If IsNumeric(Me.TxtDiscount.Text) = True And IsNumeric(Me.TxtUnitPrice.Text) = True And IsNumeric(Me.TxtQtyDefault) = True Then
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLocation.SelectedIndexChanged
        Dim rowLocation As Core_Module.dseMasterProduct.stock_location_viewRow

        Try
            If Me.cboLocation.SelectedIndex <> -1 Then
                rowLocation = TryCast(Me.TBL2_Item.DefaultView(Me.cboLocation.SelectedIndex).Row, _
                Core_Module.dseMasterProduct.stock_location_viewRow)
                Me.txtPrimaryLocation.Text = Me.ObjLocation.GetPrimary(rowLocation.Location)
            Else
                Me.txtPrimaryLocation.Text = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Try
            Me.dgvItem.TabStop = True
            If Me.TxtRemarks.Enabled = True Then
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.TxtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            'berkaitan denga import e-faktur binaya 
            If vFromQuick = True Then
                Exit Sub
            End If
            '

            If Me.TxtRemarks.Enabled = True Then
                If Me.SedangRevisi = False Then
                    If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                        If Me.baru = False Then
                            If Me.ObjInvoice.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                            Me.ObjInvoice.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                                'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                                Me.txtInvoiceNo.Text = Me.ObjInvoice.NoTransaksi
                            Else
                                'jika pindah bulan maka akan ganti no
                                Me.txtInvoiceNo.Text = Me.ObjInvoice.GetNoBaru1(Me.dtpDate.Value)
                            End If
                        Else
                            Me.txtInvoiceNo.Text = Me.ObjInvoice.GetNoBaru1(Me.dtpDate.Value)
                        End If
                    End If
                End If

                '2025 Core Tax Tahap 2.2
                'If Me.dtpDate.Value.Date < CDate(#4/1/2013#) Then
                '    'update nomor tahun pada faktur pajak standar
                '    If Me.CboFakturPajak.SelectedIndex = 1 Then
                '        Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '    End If
                'Else
                '    If Me.chkTax.Checked = True Then
                '        If Me.CboFakturPajak.SelectedIndex = 1 Then
                '            If Me.baru = True Then
                '                Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value.Date)
                '            Else
                '                If Me.ObjInvoice.TanggalTransaksi >= #4/1/2013# Then
                '                    Me.txtNoFakturPajak.Text = Me.ObjInvoice.NoFakturPajak
                '                Else
                '                    Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value.Date)
                '                End If
                '            End If
                '        End If
                '        Me.llbEditNoFakturPajak.Enabled = True
                '    Else
                '        Me.llbEditNoFakturPajak.Enabled = False
                '        Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                '        If Me.txtNoFakturPajak.Text <> "" Then
                '            Me.txtNoFakturPajak.Text = Me.txtNoFakturPajak.Text.Substring(0, 8) + "00000000"
                '        End If
                '    End If
                'End If

                Me.SetTaxRate()
                Me.SetCalculateBy()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel1.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer6.Panel2.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.txtPrimaryInvoice.BackColor, False)
            End If
            If Me.TabControl1.SelectedIndex = 1 Then
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShipment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipment.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.TxtShipment, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShipment_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipment.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtShipment.Enabled = True Then
                    If Me.TxtShipment.ReadOnly = True Then
                        Me.TxtShipment.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.TxtShipment.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShipment_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtShipment.KeyDown
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyDown(Me.TxtShipment, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShipment_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtShipment.KeyUp
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyUp(Me.TxtShipment, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShipment_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipment.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(Me.TxtShipment.Text, Me.DefaultPrecision)
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtShipment_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipment.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtShipment.Enabled = True Then
                    If Me.TxtShipment.ReadOnly = True Then
                        Me.TxtShipment.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.TxtShipment.BackColor = Me.txtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboShipmentTax_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboShipmentTax.SelectedValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPPNDipungut_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPPNDipungut.SelectedValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemarks.Enabled = False Then
                If Me.ObjInvoice.CariDetail(vPrimary, _
                Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me, Me.TblCustomDummy, Me.TblCustomInOutDummy) _
                = True Then

                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPersonCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True And Me.TxtSalesPersonCode.Enabled = True Then
                If Me.TxtSalesPersonCode.ReadOnly = True Then
                    Me.TxtSalesPersonCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSalesPersonCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPersonCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True And Me.TxtSalesPersonCode.Enabled = True Then
                If Me.TxtSalesPersonCode.ReadOnly = True Then
                    Me.TxtSalesPersonCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtSalesPersonCode.BackColor = Me.TxtPrimaryCodeNo.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True And Me.TxtDiscString.Enabled = True Then
                If Me.TxtDiscString.ReadOnly = True Then
                    Me.TxtDiscString.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDiscString.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.Enter
        Try
            Me.buffer = Me.TxtDiscString.Text
            Me.ObjCD.Disc_Enter(Me.TxtDiscString, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscString.KeyDown
        Try
            Me.ObjCD.Disc_KeyDown(Me.TxtDiscString, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscString.KeyUp
        Try
            Me.ObjCD.Disc_KeyUp(Me.TxtDiscString, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.Leave
        Try
            Me.ObjCD.Disc_Leave(Me.TxtDiscString, Me.TxtDiscString.Text, sender, e, Me.DefaultPrecision, Me.ObjCD, Me.buffer)
            Me.Hitung_Total()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscString_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscString.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True And Me.TxtDiscString.Enabled = True Then
                If Me.TxtDiscString.ReadOnly = True Then
                    Me.TxtDiscString.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDiscString.BackColor = Me.TxtPrimaryCodeNo.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoSeriFakturPajak_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True And Me.txtNoSeriFakturPajak.Enabled = True Then
                If Me.txtNoSeriFakturPajak.ReadOnly = True Then
                    Me.txtNoSeriFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtNoSeriFakturPajak.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoSeriFakturPajak_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True And Me.txtNoSeriFakturPajak.Enabled = True Then
                If Me.txtNoSeriFakturPajak.ReadOnly = True Then
                    Me.txtNoSeriFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.txtNoSeriFakturPajak.BackColor = Me.TxtPrimaryCodeNo.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkAlias_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAlias.CheckedChanged
        Try
            If Me.chkAlias.Checked = True Then
                Me.ColAliasCode.Visible = True
                Me.ColAliasDesc.Visible = True
            Else
                Me.ColAliasCode.Visible = False
                Me.ColAliasDesc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSalesPersonCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesPersonCode.KeyDown
        If Me.TxtSalesPersonCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmMasterSalesPerson_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSalesPerson.ApakahSudahAda(Me.TxtSalesPersonCode.Text) = True Then
                    Me.TxtPrimarySalesPerson.Text = Me.ObjSalesPerson.GetPrimary(Me.TxtSalesPersonCode.Text)
                    Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
                    Me.cboCustomerAddress.Focus()
                Else
                    'ada 2 SalesPerson dengan kode yang sama atau tidak ada SalesPerson yang dimaksud
                    If Me.TxtSalesPersonCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesPersonCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySalesPerson > 0 Then
                        Me.TxtPrimarySalesPerson.Text = i.PrimarySalesPerson
                        Me.TxtSalesPersonCode.Text = Me.ObjSalesPerson.GetCode(Me.TxtPrimarySalesPerson.Text)
                        Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
                        Me.cboCustomerAddress.Focus()
                    Else
                        Me.TxtSalesPersonCode.Clear()
                        Me.TxtSalesPersonName.Clear()
                        Me.TxtPrimarySalesPerson.Text = "-1"
                        Me.TxtSalesPersonCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSalesPersonCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPersonCode.Leave
        If Me.TxtSalesPersonCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmMasterSalesPerson_Cari
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtSalesPersonCode.Text <> "" Then
                If Me.ObjSalesPerson.GetJml(Me.TxtSalesPersonCode.Text) = 1 Then
                    Me.TxtPrimarySalesPerson.Text = Me.ObjSalesPerson.GetPrimary(Me.TxtSalesPersonCode.Text)
                    Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
                Else
                    If Me.TxtSalesPersonCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSalesPersonCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySalesPerson > 0 Then
                        Me.TxtPrimarySalesPerson.Text = i.PrimarySalesPerson
                        Me.TxtSalesPersonCode.Text = Me.ObjSalesPerson.GetCode(Me.TxtPrimarySalesPerson.Text)
                        Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
                    Else
                        Me.TxtSalesPersonCode.Clear()
                        Me.TxtSalesPersonName.Clear()
                        Me.TxtPrimarySalesPerson.Text = "-1"
                    End If
                End If
            Else
                Me.TxtSalesPersonCode.Clear()
                Me.TxtSalesPersonName.Clear()
                Me.TxtPrimarySalesPerson.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSalesPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSalesPerson.Click
        Dim i As New FrmMasterSalesPerson_Cari
        i.BackColorAll = Me.BackColorAll

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtSalesPersonCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSalesPersonCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySalesPerson > 0 Then
                Me.TxtPrimarySalesPerson.Text = i.PrimarySalesPerson
                Me.TxtSalesPersonCode.Text = Me.ObjSalesPerson.GetCode(Me.TxtPrimarySalesPerson.Text)
                Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
                Me.cboCustomerAddress.Focus()
            Else
                Me.TxtSalesPersonCode.Clear()
                Me.TxtSalesPersonName.Clear()
                Me.TxtPrimarySalesPerson.Text = "-1"
                Me.TxtSalesPersonCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub UpdateSalesPerson()
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtPrimaryCustomer.Text <> "-1" Then
                    If Me.TxtPrimarySalesPerson.Text = "-1" Or Me.TxtPrimarySalesPerson.Text = "" Then
                        Me.TxtPrimarySalesPerson.Text = Me.ObjCustomer.GetPrimary_SalesPerson(Me.txtPrimaryCustomer.Text)
                        Me.TxtSalesPersonCode.Text = Me.ObjSalesPerson.GetCode(Me.TxtPrimarySalesPerson.Text)
                        Me.TxtSalesPersonName.Text = Me.ObjSalesPerson.GetName(Me.TxtPrimarySalesPerson.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.Enter
        Try
            Me.Buffer2 = Me.TxtDiscAmount1.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount1, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount1.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscAmount1, sender, e, Me.TxtUnitPrice.Text)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount1.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscAmount1, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.GotFocus
        If Me.TxtDiscAmount1.ReadOnly = False And Me.TxtDiscAmount1.Enabled = True Then
            TxtDiscAmount1.BackColor = BackColorAll
        Else
            TxtDiscAmount1.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtDiscAmount1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.LostFocus
        If Me.TxtDiscAmount1.ReadOnly = False And Me.TxtDiscAmount1.Enabled = True Then
            TxtDiscAmount1.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtDiscAmount1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount1.Leave
        Try
            If Me.TxtRemarkItem.Enabled = True Then
                Me.TxtDiscAmount1.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscAmount1.Text, Me.DefaultPrecision)
                If CDec(Me.TxtDiscAmount1.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtDiscAmount1.Text = Me.Buffer2
                    Exit Sub
                End If
                If CDec(Me.TxtDiscAmount1.Text) <> CDec(Me.Buffer2) Then
                    Me.TxtDiscFromPerct.Text = "0"
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbImport_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbImport.LinkClicked
        Dim frmTransferDate As New FrmTransferBeginingBalance_Date

        Try
            If MsgBox("Start data import?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If frmTransferDate.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.LblProgressImport.Text = ""
                    Me.LblProgressImport.Visible = True
                    Me.btnFind.Visible = False
                    Me.btnNew.Visible = False
                    Me.btnEdit.Visible = False
                    Me.btnDelete.Visible = False
                    Me.btnPrint.Visible = False
                    Me.btnExit.Visible = False
                    Me.BtnFaktur.Visible = False
                    Me.vStartDate = frmTransferDate.dtpFrom.Value
                    Me.vEndDate = frmTransferDate.dtpTo.Value

                    Me.BackgroundWorker2.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frmTransferDate = Nothing
        End Try
    End Sub

    '-balikin lg
    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            TransferQuickInvoice(Me.vStartDate, Me.vEndDate)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '-balikin lagi
    Private Sub TransferQuickInvoice(ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim frmTransfer As New CustomSMESS.FrmTransferBeginingBalance
        Dim objDo As New ClsTransaksiDO

        Dim dtaSalesInvoiceCopy As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_salesinvoice_copyTableAdapter
        Dim tblSalesInvoiceCopy As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_copyDataTable
        Dim rowSalesInvoiceCopy As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_copyRow

        Dim dtaOldSalesInvoiceItem As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_salesinvoice_item_copyTableAdapter
        Dim tblOldSalesInvoiceItem As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_copyDataTable

        Dim dtaOldSalesInvoiceItemExchangeRate As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_salesinvoice_item_exchangerate_copyTableAdapter
        Dim tblOldSalesInvoiceItemExchangeRate As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_exchangerate_copyDataTable

        Dim dtaOldSalesInvoiceItemSalesReturn As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_salesinvoice_item_salesreturn_copyTableAdapter
        Dim tblOldSalesInvoiceItemSalesReturn As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_salesreturn_copyDataTable

        Dim dtaOldSalesInvoiceItemInvoiceDp As New CustomSMESS.DseTransferBeginingBalanceTableAdapters.tbltrans_salesinvoice_itemdp_copyTableAdapter
        Dim tblOldSalesInvoiceItemInvoiceDp As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_itemdp_copyDataTable

        Dim tbl As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewDataTable
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewRow

        Dim tblOldSalesInvoice1 As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice2DataTable
        Dim tblOldSalesInvoiceItem1 As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_itemDataTable
        Dim tblOldSalesInvoiceItem2 As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_itemDataTable
        Dim tblOldSalesInvoiceItemInvoiceDp1 As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_itemdpDataTable
        Dim tblOldSalesInvoiceItemExchangeRate1 As New CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_exchangerateDataTable

        Try
            Lanjut = True
            listOfTransfered.Clear()
            listOfSkiped.Clear()
            listOfFailed.Clear()
            Msg = ""
            vErrorMessage = ""

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ControlBox = False
            Me.btnFind.Enabled = False
            Me.btnNew.Enabled = False
            Me.btnEdit.Enabled = False
            Me.BtnFaktur.Enabled = False
            Me.btnDelete.Enabled = False
            Me.btnPrint.Enabled = False
            Me.btnExit.Enabled = False
            Me.llbImport.Enabled = False

            Me.objCustom.TransferInvoice_CekData(Me.ProgressBar1, vStartDate, vEndDate, _
            Me.LblProgressImport, tblOldSalesInvoice1, tblOldSalesInvoiceItem1, tblOldSalesInvoiceItem2, _
            tblOldSalesInvoiceItemInvoiceDp1, tblOldSalesInvoiceItemExchangeRate1, vErrorMessage)

            If vErrorMessage <> "" Then
                SudahSelesaiCek = False
                Me.BackgroundWorker2.ReportProgress(0)

                While SudahSelesaiCek = False
                    Threading.Thread.Sleep(1000)
                End While

                If Me.Lanjut = False Then
                    Exit Sub
                End If
            End If
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.objCustom.TransferInvoice(Me.ProgressBar1, vStartDate, vEndDate, _
            Me.LblProgressImport, tblOldSalesInvoice1, tblOldSalesInvoiceItem1, _
            tblOldSalesInvoiceItem2, tblOldSalesInvoiceItemInvoiceDp1, tblOldSalesInvoiceItemExchangeRate1)

            dtaSalesInvoiceCopy.Fill(tblSalesInvoiceCopy, vStartDate, vEndDate)
            dtaOldSalesInvoiceItem.Fill(tblOldSalesInvoiceItem)
            dtaOldSalesInvoiceItemExchangeRate.Fill(tblOldSalesInvoiceItemExchangeRate)
            dtaOldSalesInvoiceItemSalesReturn.Fill(tblOldSalesInvoiceItemSalesReturn)
            dtaOldSalesInvoiceItemInvoiceDp.Fill(tblOldSalesInvoiceItemInvoiceDp)
            Me.ProgressBar1.Maximum = tblSalesInvoiceCopy.Rows.Count

            For Each rowSalesInvoiceCopy In tblSalesInvoiceCopy
                Me.ProgressBar1.Value = tblSalesInvoiceCopy.Rows.IndexOf(rowSalesInvoiceCopy) + 1
                Me.LblProgressImport.Text = "Saving to database, data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + rowSalesInvoiceCopy.No_Transaksi

                tbl.Clear()
                row = tbl.Newtbltrans_salesinvoice_NewRow
                With rowSalesInvoiceCopy
                    row.Primary = .Primary
                    row.No_Transaksi = .No_Transaksi
                    row.Urutan = .Urutan
                    row.Urutan_Faktur_Pajak = .Urutan_Faktur_Pajak
                    row.Tipe_Invoice = .Tipe_Invoice
                    row.Tanggal_Transaksi = .Tanggal_Transaksi
                    row.Primary_Customer = .Primary_Customer
                    row.Primary_Currency = .Primary_Currency
                    row.Calculates_By = .Calculates_By
                    row.Tax_Exclude = .Tax_Exclude
                    row.Tax = .Tax
                    row.No_Seri_Faktur_Pajak = .No_Seri_Faktur_Pajak
                    row.No_Faktur_Pajak = .No_Faktur_Pajak
                    row.No_NPWP = .No_NPWP
                    row.Address_Identifier = .Address_Identifier
                    row.Remarks = .Remarks
                    row.Created_By = .Created_By
                    row.Created_Date = .Created_Date
                    row.Modified_By = .Modified_By
                    row.Modified_Date = .Modified_Date
                    row.Timestamp = .Timestamp
                    row.Jenis_Faktur_Pajak = .Jenis_Faktur_Pajak
                    row.Payment_Term = .Payment_Term
                    row.JmlCetak = .JmlCetak
                    row.Revisi = .Revisi
                    row.Primary_SO_4_DP = .Primary_SO_4_DP
                    row.No_Faktur_Pajak_Sederhana = .No_Faktur_Pajak_Sederhana
                    row.Nama_Signed = .Nama_Signed
                    row.Jabatan_Signed = .Jabatan_Signed
                    row.Alamat_Pajak_1 = .Alamat_Pajak_1
                    row.Alamat_Pajak_2 = .Alamat_Pajak_2
                    row.Sub_Total_Before_DP = .Sub_Total_Before_DP
                    row.Down_Payment = .Down_Payment
                    row.Sub_Total = .Sub_Total
                    row.Duty_Stamp = .Duty_Stamp
                    row.Tax_Exchange_Rate = .Tax_Exchange_Rate
                    row.Tax_Amount = .Tax_Amount
                    row.Total = .Total
                    row.Exchange_Rate_To_USD = .Exchange_Rate_To_USD
                    row.Exchange_Rate_To_IDR = .Exchange_Rate_To_IDR
                    row.exchange_rate = .exchange_rate
                    row.ongkos_kirim = .ongkos_kirim
                    row.VAT_Dipungut = .VAT_Dipungut
                    row.Shipment_With_VAT = .Shipment_With_VAT
                    row.primary_sales_person = .primary_sales_person
                    row.disc_string = .disc_string
                    row.disc_amount = .disc_amount
                    row.Nama_Jalan = .Nama_Jalan
                    row.Blok = .Blok
                    row.Nomor = .Nomor
                    row.RT = .RT
                    row.RW = .RW
                    row.Kecamatan = .Kecamatan
                    row.Kelurahan = .Kelurahan
                    row.Kabupaten = .Kabupaten
                    row.Propinsi = .Propinsi
                    row.Kode_Pos = .kode_jenis_transaksi_pajak
                    row.Telepon = .Telepon
                    row.kode_jenis_transaksi_pajak = .kode_jenis_transaksi_pajak
                End With
                tbl.Rows.Add(row)
                Me.ObjInvoice.CariDetail(tbl, tblOldSalesInvoiceItem, tblOldSalesInvoiceItemExchangeRate, _
                        tblOldSalesInvoiceItemSalesReturn, tblOldSalesInvoiceItemInvoiceDp, _
                        Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me, Me.TblCustomDummy, Me.TblCustomInOutDummy)
                If Me.ObjInvoice.SimpanBaru(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                        Me.dgvExchageRates, Me.UserCode, Me, objDo.GetNoBaru_Obs(row.Tanggal_Transaksi), Me.TblCustomDummy, Me.TblCustomInOutDummy) = True Then
                    listOfTransfered.Add(rowSalesInvoiceCopy.No_Transaksi)
                Else
                    listOfFailed.Add(rowSalesInvoiceCopy.No_Transaksi)
                    'If MsgBox("Stop process?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    '    Exit For
                    'End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblOldSalesInvoice1.Dispose()
            tblOldSalesInvoiceItem1.Dispose()
            tblOldSalesInvoiceItem2.Dispose()
            tblOldSalesInvoiceItemInvoiceDp1.Dispose()
            tblOldSalesInvoiceItemExchangeRate1.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        Try
            FrmErr.txtError.Text = vErrorMessage
            Me.LblProgressImport.Text = "Waiting form confirmation..."
            If FrmErr.ShowDialog <> Windows.Forms.DialogResult.OK Then
                Me.Lanjut = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SudahSelesaiCek = True
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If listOfTransfered.Count > 0 Then
                Msg = "Invoice transfered: " & listOfTransfered.Count
            End If
            If listOfSkiped.Count > 0 Then
                If String.IsNullOrEmpty(Msg) = False Then
                    Msg += vbNewLine
                End If
                Msg += "Invoice skiped: " & listOfSkiped.Count
            End If
            If listOfFailed.Count > 0 Then
                If String.IsNullOrEmpty(Msg) = False Then
                    Msg += vbNewLine
                End If
                Msg += "Invoice failed: " & listOfFailed.Count
            End If
            If String.IsNullOrEmpty(Msg) = False Then
                MsgBox(Msg, MsgBoxStyle.Information, "Information")
            Else
                MsgBox("No invoice transfered.", MsgBoxStyle.Information, "Information")
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()
            Me.KosongkanField()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = 0
            Me.ControlBox = True
            Me.btnFind.Enabled = True
            Me.btnNew.Enabled = True
            Me.btnEdit.Enabled = True
            Me.BtnFaktur.Enabled = True
            Me.btnDelete.Enabled = True
            Me.btnPrint.Enabled = True
            Me.btnExit.Enabled = True
            Me.llbImport.Enabled = True
            Me.Cursor = Cursors.Default

            Me.btnFind.Visible = True
            Me.btnNew.Visible = True
            Me.btnEdit.Visible = True
            Me.btnDelete.Visible = True
            Me.btnPrint.Visible = True
            Me.btnExit.Visible = True
            Me.LblProgressImport.Visible = False
            Me.BtnFaktur.Visible = True
        End Try
    End Sub

    Private Sub TxtQtyMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.GotFocus
        If Me.TxtQtyMulti.ReadOnly = False And Me.TxtQtyMulti.Enabled = True Then
            TxtQtyMulti.BackColor = BackColorAll
        Else
            TxtQtyMulti.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtQtyMulti_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQtyMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyMulti.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQtyMulti, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.Leave
        Dim UnitPrice, Diskon As Decimal

        Try
            If Me.TxtRemarkItem.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    Me.TxtQtyMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                End If

                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                                                   Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                Else
                    Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                    Me.TxtPrimaryCurrency1.Text, Me.TxtPrimaryUnitsDefault.Text, UnitPrice, Diskon)
                End If
                If UnitPrice = 0 And Diskon = 0 Then
                    If Me.txtPriceByMultiUnit.Text = "1" Then
                        Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                        UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                    Else
                        Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                        UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                    End If
                End If
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
                Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                Me.TxtDiscFromPerct.Text = "1"
                Me.SetQtyDefault()
                Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Dim RatioDefault As Decimal = 0
        Dim RatioReference As Decimal = 0
        Try
            RatioDefault = Me.ObjSO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
            RatioReference = Me.ObjSO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text)
            Me.TxtQtyDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyMulti.Text) * RatioDefault / RatioReference, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            Me.PrepareMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrepareMultiUnit()
        Dim Dta As New DseTransaksiSOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Try
            Me.CboUnit.DataSource = Nothing
            If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                Dta.Connection.ConnectionString = Core_Module.GetConstring
                Dta.FillBy_PrimaryProduct(Me.mTblMultiUnit, Me.TxtPrimaryCodeNo.Text)
                Me.IsiCboUnit(Me.TxtPrimaryCodeNo.Text, Me.CboUnit, Me.tblCboUnit, Me.mTblMultiUnit)
                Me.CboUnit.SelectedIndex = -1
                Me.CboUnit.SelectedIndex = 0
            Else
                TxtPrimaryUnitsDefault.Text = "-1"
                TxtPrimaryUnitsMulti.Text = "-1"
                lblUnitsDefault.Text = "Units"
                'lblUnitsMulti.Text = "Units"
                TxtQtyDefault.Text = "0"
                TxtQtyMulti.Text = "0"
                LblUnitPrice.Text = "Unit Price"
                CboUnit.DataSource = Nothing
                CboUnit.Refresh()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsiCboUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
            ByVal vTbl As DseTransaksiSO.CboUnitDataTable, _
            ByVal vTblMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitDataTable) As Boolean
        Dim rowTbl As DseTransaksiSO.CboUnitRow
        Dim rowMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitRow
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objUnit As New Core_Module.ClsMasterUnit
        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vTbl.Rows.Clear()
            vTblMultiUnit.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
            rowTbl = vTbl.NewCboUnitRow
            rowTbl.Primary = objProduct.GetPrimaryUnit(vPrimaryProduct)
            rowTbl.Code = objUnit.GetUnitCode(rowTbl.Primary)
            vTbl.AddCboUnitRow(rowTbl)
            If vTblMultiUnit.DefaultView.Count > 0 Then
                For i As Integer = 0 To vTblMultiUnit.DefaultView.Count - 1
                    rowMultiUnit = vTblMultiUnit.DefaultView(i).Row
                    rowTbl = vTbl.NewCboUnitRow
                    rowTbl.Primary = rowMultiUnit.Primary_Unit
                    rowTbl.Code = objUnit.GetUnitCode(rowMultiUnit.Primary_Unit)
                    vTbl.AddCboUnitRow(rowTbl)
                Next
            End If
            vCbo.DataSource = vTbl.DefaultView
            vTbl.DefaultView.Sort = "[Code] ASC"
            vCbo.DisplayMember = "Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub CboUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnit.Leave
        Dim UnitPrice, Diskon As Decimal

        Try
            UnitPrice = 0
            Diskon = 0

            If Me.CboUnit.Text <> Me.cboPriceByMultiUnit.Text Then
                Me.txtPriceByMultiUnit.Text = 0
            Else
                Me.txtPriceByMultiUnit.Text = 1
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsDefault.Text, UnitPrice, Diskon)
            End If
            If UnitPrice = 0 And Diskon = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                Else
                    Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
            Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) 

            Me.Hitung_Amount2()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnit.SelectedIndexChanged
        'Dim UnitPrice, Discount As Decimal
        Dim row As DseTransaksiSO.CboUnitRow
        Try
            If Me.CboUnit.SelectedIndex <> -1 Then
                row = Me.tblCboUnit.DefaultView(Me.CboUnit.SelectedIndex).Row
                Me.TxtPrimaryUnitsMulti.Text = row.Primary
                'Me.lblUnitsMulti.Text = Me.CboUnit.Text
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnit.Text
                Me.TxtQtyMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.TxtPrimaryUnitsMulti.Text))
                'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                'Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Discount)
                'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.SetQtyDefault()
                Me.SetCboPriceByMultiUnit()
                'Me.Hitung_Amount2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyMulti_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.LostFocus
        If Me.TxtQtyMulti.ReadOnly = False And Me.TxtQtyMulti.Enabled = True Then
            TxtQtyMulti.BackColor = Me.txtPrimaryInvoice.BackColor
        End If
    End Sub

    Private Sub TxtQtyMulti_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyMulti.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            If IsNumeric(Me.TxtQtyMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                SetQtyDefault()
                If CDec(TxtUnitPrice.Text) > 0 Then
                    Me.Hitung_Amount2()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefaultOnDgv1(ByVal vRow As DataGridViewRow, ByVal vDefPrecDefault As Integer)
        Try
            vRow.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(vRow.Cells(Me.ColTotalQtyMulti.Index).Value) * CDec(vRow.Cells(Me.ColRatioDefault.Index).Value) / CDec(vRow.Cells(Me.ColRatioReference.Index).Value), vDefPrecDefault)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetRatioOnDgv(ByVal vDgvRow As DataGridViewRow)
        Try
            With vDgvRow
                If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnit.Index).Value Then
                    .Cells(Me.ColRatioDesc.Index).Value = .Cells(Me.ColRatioReference.Index).Value & " " & .Cells(Me.ColUnitsMulti.Index).Value & " = " & _
                    .Cells(Me.ColRatioDefault.Index).Value & " " & .Cells(Me.ColUnit.Index).Value
                Else
                    .Cells(Me.ColRatioDesc.Index).Value = "-"
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub lblOtherCharge_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblOtherCharge.LinkClicked
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim I As New FrmTransaksiSalesInvoice_AddData
        Try
            If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso Me.txtPrimaryInvoice.Text > 0) Then
                'KoleksiNamaData.Add("customer code")
                'KoleksiTextBox.Add(Me.txtCustomerCode)
                KoleksiNamaData.Add("currency")
                KoleksiTextBox.Add(Me.txtCurrencyCode)

                If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                    Me.Hitung_Total()
                    I.TBLHeaderCost.Rows.Clear()
                    I.TBLitemCost.Rows.Clear()
                    Exit Sub
                End If

                If Me.txtPrimaryInvoice.Text = "" Then
                    I.PrimaryInvoice = -1
                Else
                    I.PrimaryInvoice = Me.txtPrimaryInvoice.Text
                End If
                I.PrimaryCustomerShipTo = Me.PrimaryCustomerShipTo
                I.IdentifierAddressCustomerShipTo = Me.IdentifierAddressCustomerShipTo
                I.PrimaryCurrency = Me.txtPrimaryCurrency.Text
                I.PrimaryCustomer = Me.txtPrimaryCustomer.Text
                I.TBLHeaderCost = Me.TBLHeaderCost
                I.rowHeaderCost = Me.rowHeaderCost
                I.TBLitemCost = Me.TBLitemCost
                I.rowitemCost = Me.rowitemCost
                I.OtherCharge = Me.txtOtherCharge.Text
                I.OtherChargeDummy = Me.txtOtherCharge.Text
                I.Text = "Additional data (all amount in " + Me.txtCurrencyCode.Text + ")"
                If Me.TxtRemarks.Enabled = True Then
                    I.Aktif = True
                Else
                    I.Aktif = False
                End If
                I.ShowDialog()

                If IsNumeric(I.OtherCharge) Then
                    Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping(I.OtherCharge, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                Else
                    Me.txtOtherCharge.Text = 0
                End If
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            I.Dispose()
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Sub

    Private Sub txtPrimaryCustomer_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCustomer.TextChanged

        Try
            If Me.Visible = False Then Exit Sub
            If Me.TxtRemarks.Enabled = True Then
                Me.txtOtherCharge.Text = 0
                Me.Hitung_Total()
                Me.TBLHeaderCost.Rows.Clear()
                Me.TBLitemCost.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FromCutOff_SalesInvoiceQuick_Create(ByVal vPrimaryTransaksi As Double, ByVal vNoSalesInv As String)

        Try
            Me.baru = False
            Me.ObjInvoice.ResetObject(Me.dgvItem, Me.dgvDP, Me.dgvExchageRates, Me.dgvItemSR)
            Me.ObjInvoice.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice

            Me.ObjInvoice.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me, Me.TblCustomDummy, Me.TblCustomInOutDummy)
            Me.ObjInvoice.NoTransaksi = vNoSalesInv
            Me.UserCode = Me.ObjInvoice.CreatedBy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_SalesInvoiceQuick_Save(ByVal vNoDO As String) As Double
        Try
            Me.baru = True
            Me.ObjInvoice.SimpanBaru1(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me.UserCode, Me, False, vNoDO, Me.TblCustomDummy, Me.TblCustomInOutDummy, False, False, True)
            Return Me.ObjInvoice.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub cboNIKPASS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboNIKPASS.SelectedIndexChanged
        '//TR : Tambahan baru Nik/Pas

        Try
            If Me.txtPrimaryCustomer.Text <> "-1" And Me.txtPrimaryCustomer.Text <> "" Then
                Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass2(Me.txtPrimaryCustomer.Text, Me.cboNIKPASS.SelectedIndex)
            End If
            If Me.cboNIKPASS.SelectedIndex = 0 Then
                Me.OptionNIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByNIK
            ElseIf Me.cboNIKPASS.SelectedIndex = 1 Then
                Me.OptionNIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByPASSPORT
            Else '2025 Core Tax Tahap 2.1
                Me.OptionNIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByOtherID
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboAddressShipTo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboAddressShipTo.CheckedChanged
        Try
            If RboAddressShipTo.Checked Then
                DefaultAddressShipToCondition()
            Else
                OtherCustomerAddressShipToCondition()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboCustomerShipTo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboCustomerShipTo.CheckedChanged
        Try
            If RboCustomerShipTo.Checked Then
                OtherCustomerAddressShipToCondition()
            Else
                DefaultAddressShipToCondition()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub DefaultAddressShipToCondition()
        Try
            If Me.txtRemarks.Enabled = True Then
                CboAddressShipTo.Enabled = True
            End If
            LlbViewAddressShipTo.Enabled = True
            TxtCustomerCodeShipTo.Enabled = False
            BtnFindCustomerShipTo.Enabled = False
            CboAddressCustomerShipTo.Enabled = False
            LlbViewAddressCustomerShipTo.Enabled = False
            TxtPrimaryCustomerShipTo.Text = "-1"
            TxtCustomerCodeShipTo.Text = ""
            TxtCustomerDescShipTo.Text = ""
            TxtIdentifierAddressCustomerShipTo.Text = ""
            CboAddressCustomerShipTo.Items.Clear()
            If Me.txtPrimaryCustomer.Text = "" Then Me.txtPrimaryCustomer.Text = "-1"
            ObjCustomer.IsiCboAddress(CboAddressShipTo, Me.txtPrimaryCustomer.Text)
            If CboAddressShipTo.Items.Count > 0 Then
                CboAddressShipTo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub OtherCustomerAddressShipToCondition()
        Try
            CboAddressShipTo.Enabled = False
            LlbViewAddressShipTo.Enabled = False
            If Me.txtRemarks.Enabled = True Then
                TxtCustomerCodeShipTo.Enabled = True
                BtnFindCustomerShipTo.Enabled = True
                CboAddressCustomerShipTo.Enabled = True
            End If
            LlbViewAddressCustomerShipTo.Enabled = True
            TxtIdentifierAddressShipTo.Text = ""
            CboAddressShipTo.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboAddressShipTo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboAddressShipTo.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.CboAddressShipTo.SelectedIndex <> -1 Then
                    Me.TxtIdentifierAddressShipTo.Text = Me.ObjCustomer.GetIdentifierAdd(Me.txtPrimaryCustomer.Text, Me.CboAddressShipTo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboAddressCustomerShipTo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboAddressCustomerShipTo.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.CboAddressCustomerShipTo.SelectedIndex <> -1 Then
                    Me.TxtIdentifierAddressCustomerShipTo.Text = Me.ObjCustomer.GetIdentifierAdd(Me.TxtPrimaryCustomerShipTo.Text, Me.CboAddressCustomerShipTo.Text)

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbViewAddressShipTo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressShipTo.LinkClicked
        Try
            If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                Me.txtCustomerCode.Focus()
                Exit Sub
            End If
            Dim ij As New Core_Module.FrmDetailAddress
            ij.vAddIdentifier = Me.TxtIdentifierAddressShipTo.Text
            ij.PrimaryCust = Me.txtPrimaryCustomer.Text
            ij.btnOK.BringToFront()
            ij.AcceptButton = ij.btnOK
            ij.btnExit.TabStop = False
            If Me.TxtRemarks.Enabled = False Then
                ij.cboCustAddCode.Enabled = False
            End If
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbViewAddressCustomerShipTo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressCustomerShipTo.LinkClicked
        Try
            If Me.TxtPrimaryCustomerShipTo.Text = "" Or Me.TxtPrimaryCustomerShipTo.Text = "-1" Then
                MsgBox("Please specify ship to customer first.", MsgBoxStyle.Information, "Information")
                Me.TxtCustomerCodeShipTo.Focus()
                Exit Sub
            End If
            Dim ij As New Core_Module.FrmDetailAddress
            ij.vAddIdentifier = Me.TxtIdentifierAddressCustomerShipTo.Text
            ij.PrimaryCust = Me.TxtPrimaryCustomerShipTo.Text
            ij.btnOK.BringToFront()
            ij.AcceptButton = ij.btnOK
            ij.btnExit.TabStop = False
            If Me.TxtRemarks.Enabled = False Then
                ij.cboCustAddCode.Enabled = False
            End If
            ij.ShowDialog()
            If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
            End If
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindCustomerShipTo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindCustomerShipTo.Click
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustomerCodeShipTo.Text <> "" Then
                i.txtCari.Text = Me.TxtCustomerCodeShipTo.Text
            End If
            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCustomerShipTo.Text = i.PrimaryCustomer
                Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomerShipTo.Text)
                Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                Me.CboAddressCustomerShipTo.SelectedIndex = 0
                Me.CboAddressCustomerShipTo.Focus()
            Else
                Me.TxtCustomerCodeShipTo.Clear()
                Me.TxtCustomerDescShipTo.Clear()
                Me.CboAddressCustomerShipTo.Items.Clear()
                Me.TxtIdentifierAddressCustomerShipTo.Text = ""
                Me.TxtPrimaryCustomerShipTo.Text = "-1"
                Me.TxtCustomerCodeShipTo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCustomerCodeShipTo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.GotFocus
        TxtCustomerCodeShipTo.BackColor = BackColorAll
    End Sub

    Private Sub TxtCustomerCodeShipTo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerCodeShipTo.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtCustomerCodeShipTo.Text) = True Then
                    Me.TxtPrimaryCustomerShipTo.Text = Me.ObjCustomer.GetPrimary(Me.TxtCustomerCodeShipTo.Text)
                    Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    Me.CboAddressCustomerShipTo.SelectedIndex = 0
                Else
                    Me.BtnFindCustomerShipTo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCodeShipTo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.Leave
        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtCustomerCodeShipTo.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtCustomerCodeShipTo.Text) = True Then
                    Me.TxtPrimaryCustomerShipTo.Text = Me.ObjCustomer.GetPrimary(Me.TxtCustomerCodeShipTo.Text)
                    Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    Me.CboAddressCustomerShipTo.SelectedIndex = 0
                Else
                    If Me.TxtCustomerCodeShipTo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerCodeShipTo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCustomerShipTo.Text = i.PrimaryCustomer
                        Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomerShipTo.Text)
                        Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                        Me.CboAddressCustomerShipTo.SelectedIndex = 0
                    Else
                        Me.TxtCustomerCodeShipTo.Clear()
                        Me.TxtCustomerDescShipTo.Clear()
                        Me.CboAddressCustomerShipTo.Items.Clear()
                        Me.TxtIdentifierAddressCustomerShipTo.Text = ""
                        Me.TxtPrimaryCustomerShipTo.Text = "-1"
                    End If
                End If
            Else
                Me.TxtCustomerCodeShipTo.Clear()
                Me.TxtCustomerDescShipTo.Clear()
                Me.CboAddressCustomerShipTo.Items.Clear()
                Me.TxtIdentifierAddressCustomerShipTo.Text = ""
                Me.TxtPrimaryCustomerShipTo.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCustomerCodeShipTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.LostFocus
        TxtCustomerCodeShipTo.BackColor = Me.txtPrimaryCustomer.BackColor
    End Sub 

    Private Sub TxtIdentifierAddressShipTo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtIdentifierAddressShipTo.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            Me.UpdateShippingData
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtIdentifierAddressCustomerShipTo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtIdentifierAddressCustomerShipTo.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            Me.UpdateShippingData
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryCustomerShipTo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCustomerShipTo.TextChanged
        Try
            If Me.Visible = False Then Exit Sub
            Me.UpdateShippingData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub UpdateShippingData()
        Try
            If Me.txtPrimaryCustomer.Text = "" Then Me.txtPrimaryCustomer.Text = "-1"
            If Me.TxtPrimaryCustomerShipTo.Text = "" Then Me.TxtPrimaryCustomerShipTo.Text = "-1"
            If Me.RboAddressShipTo.Checked Then
                Me.PrimaryCustomerShipTo = Me.txtPrimaryCustomer.Text
                Me.IdentifierAddressCustomerShipTo = Me.TxtIdentifierAddressShipTo.Text
            Else
                Me.PrimaryCustomerShipTo = Me.TxtPrimaryCustomerShipTo.Text
                Me.IdentifierAddressCustomerShipTo = Me.TxtIdentifierAddressCustomerShipTo.Text
            End If
            For Each rowHeaderCost In TBLHeaderCost
                rowHeaderCost.primary_cust_ship_to = Me.PrimaryCustomerShipTo
                rowHeaderCost.identifier_cust_ship_to = Me.IdentifierAddressCustomerShipTo
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub Atur1()
        Dim AksesAddCustomProduct As Boolean = False 'binaya + admiral
        'Dim AksesAddCustomItem As Boolean = False 'admiral

        Try
            AksesAddCustomProduct = Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoiceQuick_AddByCustomProduct")
            'AksesAddCustomItem = Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoiceQuick_AddItembyCustom")
            Me.ToolStripMenuItem1.Visible = False
            Me.AddItemByCustom.Visible = False
            If AksesAddCustomProduct = False Then 'And AksesAddCustomItem = False Then
                'Me.GroupBox2.Enabled = True
                Me.SplitContainer5.Enabled = True
                Me.dgvItem.ReadOnly = False
                Me.dgvItem.AllowUserToDeleteRows = True
            Else
                If AksesAddCustomProduct = True Then
                    Me.ToolStripMenuItem1.Visible = True
                End If
                'If AksesAddCustomItem = True Then
                '    Me.AddItemByCustom.Visible = True
                'End If
                Me.SplitContainer5.Enabled = False
                Me.dgvItem.ReadOnly = True
                'Me.GroupBox2.Enabled = False
                Me.dgvItem.AllowUserToDeleteRows = False
            End If
            Me.LinkLabel1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsikanDgvFromTblDummy() As Boolean
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim rowDummy As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1Row
        Dim indexSama As Integer = -1
        Dim row As DataGridViewRow
        Dim AngkaDecimal As Double = 0
        Dim AngkaDecimalCurrency As Double = 0
        Dim PrimaryCurrenct As Double = -1
        Try
            Me.dgvItem.Rows.Clear()
            For Each rowDummy In Me.TblCustomDummy.Rows
                indexSama = Me.GetIndex(rowDummy.Unit_Price, rowDummy.Batch_Number, rowDummy.Discount, rowDummy.Primary_Location, rowDummy.Primary_Product)
                Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product, rowDummy.Primary_Units_Multi)
                If indexSama = -1 Then
                    'tambahkan baris baru
                    Me.dgvItem.Rows.Add()
                    row = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    AngkaDecimal = Me.ObjProduct.GetAngkaDecimal(rowDummy.Primary_Product)
                    AngkaDecimalCurrency = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)

                    With row
                        .Cells(Me.ColNo0.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColCodeNo.Index).Value = rowDummy.Product_Code
                        .Cells(Me.ColDescription.Index).Value = rowDummy.Description
                        .Cells(Me.ColCustCodeNo.Index).Value = rowDummy.Customer_s_Code_No
                        .Cells(Me.ColCustDesc.Index).Value = rowDummy.Customer_s_Desc
                        .Cells(Me.ColSONO.Index).Value = rowDummy.SO_Number
                        .Cells(Me.ColCustPONo.Index).Value = ""
                        .Cells(Me.ColDONO.Index).Value = ""
                        .Cells(Me.ColStockOutNo.Index).Value = ""
                        .Cells(Me.ColLotNumber.Index).Value = rowDummy.Batch_Number
                        .Cells(Me.ColExpiry.Index).Value = rowDummy.Expiry
                        .Cells(Me.ColLocationCode.Index).Value = rowDummy.Location
                        .Cells(Me.ColQtyPerMaster.Index).Value = 0
                        .Cells(Me.ColQtyMaster.Index).Value = 0
                        .Cells(Me.ColMasterUnits.Index).Value = ""
                        .Cells(Me.ColQty.Index).Value = 0
                        .Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Total_Qty), AngkaDecimal)
                        .Cells(Me.ColUnit.Index).Value = rowDummy.Unit_Code
                        .Cells(Me.ColSecQty.Index).Value = -1
                        .Cells(Me.ColSecUnit.Index).Value = ""
                        .Cells(Me.ColCurrency.Index).Value = rowDummy.currency_Code
                        .Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Unit_Price), AngkaDecimalCurrency)
                        .Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Discount), 2)
                        .Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Nett_Price), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.Amount), AngkaDecimalCurrency)
                        .Cells(Me.ColRemarksItem.Index).Value = rowDummy.Remarks
                        .Cells(Me.ColPrimaryLocation.Index).Value = rowDummy.Primary_Location
                        .Cells(Me.ColPrimaryDO.Index).Value = rowDummy.Primary_DO
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = rowDummy.Primary_Product
                        .Cells(Me.ColPrimarySO.Index).Value = rowDummy.Primary_SO
                        .Cells(Me.ColPrimaryMasterUnits.Index).Value = rowDummy.Primary_Master_Unit
                        .Cells(Me.ColPrimaryUnit.Index).Value = rowDummy.Primary_Units
                        .Cells(Me.ColPrimarySecUnit.Index).Value = rowDummy.Primary_Secondary_Units
                        .Cells(Me.ColPrimaryCurrency.Index).Value = rowDummy.Primary_Currency
                        .Cells(Me.ColAliasDesc.Index).Value = Me.ObjInvoiceItem.GetAliasDesc(rowDummy.Primary_Product, Me.txtPrimaryCustomer.Text)
                        .Cells(Me.ColAliasCode.Index).Value = Me.ObjInvoiceItem.GetAliasCode(rowDummy.Primary_Product, Me.txtPrimaryCustomer.Text)

                        .Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Disc_Amount, AngkaDecimalCurrency)
                        .Cells(Me.ColDiscFromPrect.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.DiscFromPerc, AngkaDecimalCurrency)

                        .Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(rowDummy.Total_Qty_Multi, DefPrecMulti)
                        .Cells(Me.ColUnitsMulti.Index).Value = rowDummy.Unit_Code_Multi
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = rowDummy.Primary_Units_Multi
                        .Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Ratio_Default, AngkaDecimal)
                        .Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.Ratio_Reference, DefPrecMulti)
                        '.Cells(Me.ColRatioDesc.Index).Value = rowDummy.Ratio

                        .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((CDec(rowDummy.Nett_Price) * CDec(Me.txtTax.Text)) / 100, AngkaDecimalCurrency)
                        If AngkaDecimalCurrency = 0 Then
                            If Me.txtNettPrice.Text.EndsWith("55") Then
                                .Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDummy.VAT) - 1, AngkaDecimalCurrency)
                            End If
                        End If
                        .Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(rowDummy.VAT * CDec(.Cells(Me.ColTotalQtyMulti.Index).Value), AngkaDecimalCurrency)
                        .Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(.Cells(Me.ColVAT2.Index).Value) + CDec(.Cells(Me.ColAmount.Index).Value), AngkaDecimalCurrency)
                        Me.SetRatioOnDgv(row)
                    End With
                Else
                    row = Me.dgvItem.Rows(indexSama)
                    'tambahkan qty multi, 
                    row.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value) + CDec(rowDummy.Total_Qty_Multi), DefPrecMulti)
                    'hitung ulang amount
                    row = Me.dgvItem.Rows(indexSama)

                    Me.SetQtyDefaultOnDgv1(row, Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, row.Index).Value))

                    Me.Hitung_Amount(row)
                End If

                PrimaryCurrenct = Me.TxtPrimaryCurrency1.Text
                'Me.KosongkanItem()
                'Me.TxtCodeNo.Focus()

                Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
                Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                Me.FormatDGV2()

                Me.cboCalculatesBy.SelectedIndex = -1
                Me.cboCalculatesBy.SelectedIndex = 0

                'fokuskan ke baris terakhir
            Next

            If Me.ApakahAdaAlias = True Then
                Me.chkAlias.Checked = True
            Else
                Me.chkAlias.Checked = False
            End If

            Me.Hitung_Total()
            If Me.vFromQuick = True Then
                Me.HitungVATItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Function

    Public Sub AddByCustomBinaya()
        Dim Frm As New CustomSMESS.FrmTransaksiSalesInvoiceQuick_AddItembyCustom
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            Frm.PrimaryInvoice = Me.txtPrimaryInvoice.Text
            Frm.PrimaryCustomer = Me.txtPrimaryCustomer.Text
            Frm.PrimaryCurrency = Me.txtPrimaryCurrency.Text
            Frm.Tax = Me.txtTax.Text
            Frm.DefaultPrecision = Me.DefaultPrecision
            Frm.DefaultPrecision_DefCur = Me.DefaultPrecision_DefCur
            If Me.chkTax.Checked = True Then
                Frm.IsTax = True
            End If
            If Me.btnOK.Enabled = True Then
                Frm.AktifForm = True
            Else
                Frm.AktifForm = False
            End If
            HitungVATItem2()
            Frm.TblDummy = Me.TblCustomDummy

            If Frm.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblCustomDummy = Frm.TblDummy
                Me.IsikanDgvFromTblDummy()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Dim i As New CustomSMESS.FrmTransaksiSalesInvoiceQuick_AddByCustomProduct
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ApakahCustomAdmiral As Boolean = False
        Dim ObjCustIdentifier As New CustomSMESS.ClsIdentifier
        Dim DummyIdentifier As String = ""

        Try
            DummyIdentifier = ObjCustIdentifier.GetIdentifier


            If Me.txtPrimaryInvoice.Text = "" Then
                Exit Sub
            End If

            If Me.TxtRemarks.Enabled = True Then
                If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
                    MsgBox("To start entering item data, you must specify customer data first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
            End If

            If DummyIdentifier = "smess_binaya" Then
                Me.AddByCustomBinaya()
                Exit Sub
            End If


            If DummyIdentifier = "smess_admiral2" Or DummyIdentifier = "smess_linda" Then
                ApakahCustomAdmiral = True
                If Me.ObjInvoice.Primary <> -1 Then
                    If Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = False Then
                        MsgBox("This invoice is not quick invoice. You can't open form advance.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            End If

            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll

            i.PrimaryInvoice = Me.txtPrimaryInvoice.Text
            i.PrimaryCustomer = Me.txtPrimaryCustomer.Text
            i.PrimaryCurrency = Me.txtPrimaryCurrency.Text
            i.TglTransaksi = Me.dtpDate.Value
            i.Tax = CDec(Me.txtTax.Text)
            i.IsTax = Me.chkTax.Checked
            i.Baru = Me.baru
            If Me.TxtRemarks.Enabled = False Then
                i.AktifForm = False
            Else
                i.AktifForm = True
            End If
            i.DefaultPrecision = Me.DefaultPrecision
            i.DefaultPrecision_DefCur = Me.DefaultPrecision_DefCur

            If ApakahCustomAdmiral = False Then
                HitungVATItem2() 'dipanggil custom binaya
            End If

            i.TblItemInvCustom = Me.TblCustomDummy.Copy
            i.TblItemCustom = Me.TblCustomInOutDummy.Copy
            If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.TblCustomDummy.Clear()
                Me.TblCustomInOutDummy.Clear()
                Me.TblCustomDummy = i.TblItemInvCustom.Copy
                Me.TblCustomInOutDummy = i.TblItemCustom.Copy
                If ApakahCustomAdmiral = False Then
                    Me.IsikanDgvFromTblDummy() 'dipanggil custom binaya
                Else
                    Me.IsiDgvByTblCustom(Me.TblCustomDummy)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.ContextMenuStrip1.Show(LinkLabel1, New Point(0, 15))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungVATItem2()
        Dim DefaultPrecisionRow As Integer = 0
        Dim rowDummy As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1Row

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.HitungVATItem2_MiddlePoint()
                Exit Sub
            End If
            If Me.TblCustomDummy.Rows.Count > 0 Then
                rowDummy = TblCustomDummy.DefaultView(0).Row
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)
            End If

            For Each rowDummy In TblCustomDummy.Rows
                With rowDummy
                    .VAT = Me.ObjCD.DigitGrouping((CDec(.Nett_Price) * CDec(Me.txtTax.Text)) / 100, DefaultPrecisionRow)
                    If .price_by_multi_unit = 1 Then
                        .VAT2 = Me.ObjCD.DigitGrouping(.VAT * CDec(.Total_Qty_Multi), DefaultPrecisionRow)
                    Else
                        .VAT2 = Me.ObjCD.DigitGrouping(.VAT * CDec(.Total_Qty), DefaultPrecisionRow)
                    End If
                    .Amount2 = Me.ObjCD.DigitGrouping(CDec(.VAT2) + CDec(.Amount), DefaultPrecisionRow)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub HitungVATItem2_MiddlePoint()
        Dim DefaultPrecisionRow As Integer = 0
        Dim rowDummy As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1Row

        Try
            If Me.TblCustomDummy.Rows.Count > 0 Then
                rowDummy = TblCustomDummy.DefaultView(0).Row
                DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(rowDummy.Primary_Currency)
            End If

            For Each rowDummy In TblCustomDummy.Rows
                With rowDummy
                    .VAT2 = Me.ObjCD.DigitGrouping((CDec(.Amount) * CDec(Me.txtTax.Text)) / 100, DefaultPrecisionRow)
                    If .price_by_multi_unit = 1 Then
                        .VAT = Me.ObjCD.DigitGrouping(.VAT2 / CDec(.Total_Qty_Multi), DefaultPrecisionRow)
                    Else
                        .VAT = Me.ObjCD.DigitGrouping(.VAT2 / CDec(.Total_Qty), DefaultPrecisionRow)
                    End If
                    .Amount2 = Me.ObjCD.DigitGrouping(CDec(.VAT2) + CDec(.Amount), DefaultPrecisionRow)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddItemByCustom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByCustom.Click
        'toolstrip ini ditutup karena yang asalnya custom admiral digabung dengan custom binaya terkait toolstrip pemanggil form nya.
        'Dim i As New CustomSMESS.FrmTransaksiSalesInvoiceQuick_AddItembyCustom

        'Try
        '    If Me.txtPrimaryInvoice.Text = "" Then
        '        Exit Sub
        '    End If

        '    If Me.TxtRemarkItem.Enabled = True Then
        '        If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
        '            MsgBox("To start entering item data, you must specify customer data first.", MsgBoxStyle.Information, "Information")
        '            Me.txtCustomerCode.Focus()
        '            Exit Sub
        '        End If
        '    End If

        '    i.UserCode = Me.UserCode
        '    i.BackColorAll = Me.BackColorAll
        '    i.TblItemInvCustom = Me.TblCustomDummy.Copy
        '    i.TblItemCustom = Me.TblCustomInOutDummy.Copy
        '    i.PrimaryTransaksi = Me.txtPrimaryInvoice.Text
        '    i.TglTransaksi = Me.dtpDate.Value
        '    i.TaxFromMain = CDec(Me.txtTax.Text)
        '    i.WithTax = Me.chkTax.Checked
        '    i.PrimaryCustomer = Me.txtPrimaryCustomer.Text
        '    i.Baru = Me.baru
        '    If Me.TxtRemarks.Enabled = False Then
        '        i.CanEditForm = False
        '    End If

        '    If i.ShowDialog() = Windows.Forms.DialogResult.OK Then
        '        Me.TblCustomDummy.Clear()
        '        Me.TblCustomInOutDummy.Clear()
        '        Me.TblCustomDummy = i.TblItemInvCustom.Copy
        '        Me.TblCustomInOutDummy = i.TblItemCustom.Copy
        '        Me.IsiDgvByTblCustom(Me.TblCustomDummy)
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub cboPriceByMultiUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPriceByMultiUnit.SelectedIndexChanged
        Dim UnitPrice, Discount As Decimal

        Try
            UnitPrice = 0
            Discount = 0

            If Me.cboPriceByMultiUnit.SelectedIndex <> -1 Then
                If Me.TxtRemarkItem.Enabled = True Then
                    If Me.cboPriceByMultiUnit.Text <> Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) Then
                        Me.txtPriceByMultiUnit.Text = "1"

                        Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyMulti.Text, _
                        Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsMulti.Text, UnitPrice, Discount)
                        If UnitPrice = 0 And Discount = 0 Then
                            Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                            UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                        End If

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                    Else
                        Me.txtPriceByMultiUnit.Text = "0"

                        Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyDefault.Text, _
                        Me.txtPrimaryCurrency.Text, Me.TxtPrimaryUnitsDefault.Text, UnitPrice, Discount)
                        If UnitPrice = 0 And Discount = 0 Then
                            Me.ObjSO.GetHistoricalPrice(Me.TxtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                                UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                        End If

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                    End If


                    If IsNumeric(Me.TxtUnitPrice.Text) = True Then
                        If CDec(Me.TxtUnitPrice.Text) > 0 Then
                            Me.Hitung_Amount2()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiDgvByTblCustom(ByVal vTBLInvoice As DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable)
        Dim rowItemInvCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1Row
        Dim rd As DataGridViewRow
        Dim indexSama As Integer = -1
        Dim PrimaryCurrenct As Double = -1
        Dim DecCurr As Integer = 0

        Try
            Me.dgvItem.Rows.Clear()
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vTBLInvoice.DefaultView.RowFilter = Nothing
            vTBLInvoice.DefaultView.Sort = "[Product Code], [Expiry], [Batch Number], [Qty At]"
            For i As Integer = 0 To vTBLInvoice.DefaultView.Count - 1
                rowItemInvCustom = vTBLInvoice.DefaultView(i).Row
                With rowItemInvCustom
                    If i = 0 Then
                        DecCurr = Me.ObjCurrency.GetDefaultPrecission(.Primary_Currency)
                    End If
                    .Default_Prec_Currency = DecCurr
                    indexSama = Me.GetIndexCustom(.Primary_Product, .Batch_Number)
                    If indexSama = -1 Then
                        'tambahkan baris baru
                        rd = Me.dgvItem.Rows(Me.dgvItem.Rows.Add())
                        rd.Cells(Me.ColNo0.Index).Value = Me.dgvItem.Rows.Count
                        rd.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rd.Cells(Me.ColDescription.Index).Value = .Description
                        rd.Cells(Me.ColCustCodeNo.Index).Value = ""
                        rd.Cells(Me.ColCustDesc.Index).Value = ""
                        rd.Cells(Me.ColSONO.Index).Value = ""
                        rd.Cells(Me.ColCustPONo.Index).Value = ""
                        rd.Cells(Me.ColDONO.Index).Value = ""
                        rd.Cells(Me.ColStockOutNo.Index).Value = ""
                        rd.Cells(Me.ColLotNumber.Index).Value = .Batch_Number
                        If IsDate(rd.Cells(Me.ColExpiry.Index).Value) = False Then
                            rd.Cells(Me.ColExpiry.Index).Value = #1/1/1950#
                        Else
                            rd.Cells(Me.ColExpiry.Index).Value = .Expiry
                        End If
                        rd.Cells(Me.ColLocationCode.Index).Value = .Location
                        rd.Cells(Me.ColQtyPerMaster.Index).Value = 0
                        rd.Cells(Me.ColQtyMaster.Index).Value = 0
                        rd.Cells(Me.ColMasterUnits.Index).Value = ""
                        rd.Cells(Me.ColQty.Index).Value = 0
                        rd.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty, .AngkaDecimalProduct)
                        rd.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rd.Cells(Me.ColSecQty.Index).Value = -1
                        rd.Cells(Me.ColSecUnit.Index).Value = ""
                        rd.Cells(Me.ColCurrency.Index).Value = .Currency_Code
                        rd.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(.Unit_Price, .Default_Prec_Currency)
                        rd.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(.Discount, 2)
                        rd.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(.Nett_Price, .Default_Prec_Currency)
                        rd.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(.Amount, .Default_Prec_Currency)
                        rd.Cells(Me.ColRemarksItem.Index).Value = .Remarks
                        rd.Cells(Me.ColPrimaryLocation.Index).Value = .Primary_Location
                        rd.Cells(Me.ColPrimaryDO.Index).Value = -1
                        rd.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rd.Cells(Me.ColPrimarySO.Index).Value = -1
                        rd.Cells(Me.ColPrimaryMasterUnits.Index).Value = -1
                        rd.Cells(Me.ColPrimaryUnit.Index).Value = .Primary_Units
                        rd.Cells(Me.ColPrimarySecUnit.Index).Value = -1
                        rd.Cells(Me.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rd.Cells(Me.ColAliasDesc.Index).Value = ""
                        rd.Cells(Me.ColAliasCode.Index).Value = ""

                        rd.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(.Disc_Amount, .Default_Prec_Currency)
                        rd.Cells(Me.ColDiscFromPrect.Index).Value = .DiscFromPerc

                        rd.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, .dec_multi)
                        rd.Cells(Me.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        If .is_kain = True Then
                            rd.Cells(Me.ColRatioDefault.Index).Value = rd.Cells(Me.ColTotalQty.Index).Value
                            rd.Cells(Me.ColRatioReference.Index).Value = rd.Cells(Me.ColTotalQtyMulti.Index).Value
                        Else
                            rd.Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Default, .AngkaDecimalProduct)
                            rd.Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, .dec_multi)
                        End If
                        rd.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi

                        rd.Cells(Me.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If rd.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                            rd.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rd.Cells(Me.ColUnitsMulti.Index).Value
                        Else
                            rd.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rd.Cells(Me.ColUnit.Index).Value
                        End If

                        rd.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((.Nett_Price * CDec(Me.txtTax.Text)) / 100, .Default_Prec_Currency)
                        If .Default_Prec_Currency = 0 Then
                            If Me.txtNettPrice.Text.EndsWith("55") Then
                                rd.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColVAT.Index).Value) - 1, .Default_Prec_Currency)
                            End If
                        End If
                        If rd.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                            rd.Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(rd.Cells(Me.ColVAT.Index).Value * CDec(rd.Cells(Me.ColTotalQtyMulti.Index).Value), .Default_Prec_Currency)
                        Else
                            rd.Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(rd.Cells(Me.ColVAT.Index).Value * CDec(rd.Cells(Me.ColTotalQty.Index).Value), .Default_Prec_Currency)
                        End If
                        rd.Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColVAT2.Index).Value) + CDec(rd.Cells(Me.ColAmount.Index).Value), .Default_Prec_Currency)
                        Me.SetRatioOnDgv(rd)
                    Else
                        rd = Me.dgvItem.Rows(indexSama)
                        'tambahkan qty multi, 
                        rd.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColTotalQtyMulti.Index).Value) + .Total_Qty_Multi, .dec_multi)
                        rd.Cells(Me.ColTotalQty.Index).Value = ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColTotalQty.Index).Value) + .Total_Qty, .AngkaDecimalProduct)
                        If .is_kain = True Then
                            rd.Cells(Me.ColRatioDefault.Index).Value = rd.Cells(Me.ColTotalQty.Index).Value
                            rd.Cells(Me.ColRatioReference.Index).Value = rd.Cells(Me.ColTotalQtyMulti.Index).Value
                        End If

                        If String.IsNullOrEmpty(.Remarks) = False Then
                            If String.IsNullOrEmpty(rd.Cells(Me.ColRemarksItem.Index).Value) = False Then
                                rd.Cells(Me.ColRemarksItem.Index).Value = rd.Cells(Me.ColRemarksItem.Index).Value & ", " & .Remarks
                            Else
                                rd.Cells(Me.ColRemarksItem.Index).Value = .Remarks
                            End If
                        End If

                        rd.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping((.Nett_Price * CDec(Me.txtTax.Text)) / 100, .Default_Prec_Currency)
                        If .Default_Prec_Currency = 0 Then
                            If Me.txtNettPrice.Text.EndsWith("55") Then
                                rd.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColVAT.Index).Value) - 1, .Default_Prec_Currency)
                            End If
                        End If
                        If rd.Cells(Me.ColPriceByMultiUnit.Index).Value = "1" Then
                            rd.Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(rd.Cells(Me.ColVAT.Index).Value * CDec(rd.Cells(Me.ColTotalQtyMulti.Index).Value), .Default_Prec_Currency)
                        Else
                            rd.Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(rd.Cells(Me.ColVAT.Index).Value * CDec(rd.Cells(Me.ColTotalQty.Index).Value), .Default_Prec_Currency)
                        End If
                        rd.Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(rd.Cells(Me.ColVAT2.Index).Value) + CDec(rd.Cells(Me.ColAmount.Index).Value), .Default_Prec_Currency)

                        'hitung ulang amount
                        rd = Me.dgvItem.Rows(indexSama)
                        Me.SetRatioOnDgv(rd)
                        'Me.SetQtyDefaultOnDgv1(rd, .AngkaDecimalProduct)

                        Me.Hitung_Amount(rd)
                    End If

                    PrimaryCurrenct = Me.TxtPrimaryCurrency1.Text
                    Me.KosongkanItem()
                    Me.TxtCodeNo.Focus()

                    Me.TxtPrimaryCurrency1.Text = PrimaryCurrenct
                    Me.TxtCurrencyCode1.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtPrimaryCurrency1.Text)
                    Me.FormatDGV2()

                    Me.cboCalculatesBy.SelectedIndex = -1
                    Me.cboCalculatesBy.SelectedIndex = 0

                    'fokuskan ke baris terakhir
                    While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                    Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                        Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While

                    If Me.ApakahAdaAlias = True Then
                        Me.chkAlias.Checked = True
                    Else
                        Me.chkAlias.Checked = False
                    End If

                    Me.Hitung_Total()
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.FormatDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetIndexCustom(ByVal vPrimaryProduct As Double, ByVal vBatchNumber As String) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If vBatchNumber = row.Cells(Me.ColLotNumber.Index).Value And vPrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'aul
    Private Sub txtNoSeriFakturPajak_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.ObjInvoice.SetDontCollect(Me.txtNoSeriFakturPajak.Text, Me.chkTax.Checked, Me.cboPPNDipungut.SelectedIndex) = True Then
                    Me.cboPPNDipungut.SelectedIndex = 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#Region "New ab"
    Private Sub SimpanLangsung()
        Dim FrmDO As New FrmTransaksiSalesInvoiceQuick_DetailNumber
        Try
            If Me.vApakahDatabenar = True Then
                'Me.IsiCBOArAkunByTipePerkiraan1()
                'Me.txtPrimaryCurrency.Text = 1
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.ObjInvoice.SimpanBaru(Me.dgvItem, Me.dgvDP, _
                            Me.dgvSR, Me.dgvExchageRates, Me.UserCode, Me, FrmDO.TxtDONumber.Text, Me.TblCustomDummy, Me.TblCustomInOutDummy, False) = True Then
                        Me.KosongkanItem()
                        Me.SalinKeField()
                        Me.NonAktifkanDetail()
                        Me.btnFind.Focus()
                    End If
                Else
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Me.NonAktifkanDetail()
                    'Me.vStopCancellation = True
                End If
            End If
            Me.NonAktifkanDetail()
            Me.btnFind.Focus()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Public Sub SetBaru()
        Dim TaxPercentSetting, UseTaxSetting As String
        Try
            Me.baru = True
            Me.SedangRevisi = False
            Me.txtPrimaryCurrency.Text = "1"
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TabControl1.SelectedIndex = 0
            Me.dtpDate.Focus()
            Me.ObjInvoice.ResetObject(Me.dgvItem, Me.dgvDP, Me.dgvExchageRates, Me.dgvItemSR)
            Me.txtInvoiceNo.Text = Me.ObjInvoice.GetNoBaru1(Me.dtpDate.Value)
            Me.CboFakturPajak.SelectedIndex = 1
            Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
            Me.ObjInvoice.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice


            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            Me.SetDefaultPrecision()
            Me.IsiCBOArAkunByTipePerkiraan1()
            Me.IsiCBOCurrencyTax()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
            Else
                If UseTaxSetting = "0" Then
                    Me.chkTax.Checked = False
                    Me.CboFakturPajak.SelectedIndex = 0
                Else
                    Me.chkTax.Checked = True
                    TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(TaxPercentSetting) = True Then
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping(TaxPercentSetting, 2)
                    End If
                End If
            End If

            Me.TxtNamaSigned.Text = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.SignedByName)
            Me.TxtJabatanSigned.Text = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.SignedByJabatan)

            Me.SetTaxRate()
            Me.SetCalculateBy()

            If Me.txtInvoiceNo.Text = "" And Me.txtInvoiceNo.Enabled = True Then
                Me.txtInvoiceNo.Focus()
            End If

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.ShipmentTax)
            If UseTaxSetting = "" Or UseTaxSetting = "0" Then
                Me.cboShipmentTax.SelectedIndex = 1
            Else
                Me.cboShipmentTax.SelectedIndex = 0
            End If
            Me.cboPPNDipungut.SelectedIndex = 0
            Me.Hitung_Total()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub GetCustomer_Leave()
        If Me.txtCustomerCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New Core_Module.FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.GetJml(Me.txtCustomerCode.Text) = 1 Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    If Me.CboAddressShipTo.Items.Count > 0 Then
                        Me.CboAddressShipTo.SelectedIndex = 0
                    End If
                    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                    '//TR : Tambahan baru Nik/Pas
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    If txtNoNPWP.Text = "00.000.000.0-000.000" Or Me.txtNoNPWP.Text = "" Then
                        'Me.llbUpDateNPWP.Visible = False
                        Me.cboNIKPASS.Visible = True
                        Me.cboNIKPASS.TabStop = True
                        Me.txtNoNIKPass.Visible = True
                        Me.txtNoNPWP.Visible = False
                        Me.cboNIKPASS.SelectedIndex = Me.ObjCustomer.GetPilihNikPass(Me.txtPrimaryCustomer.Text)
                        If Me.baru = True Then
                            Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNIKPass.Text = Me.ObjInvoice.GetNoNikPassbyInvoice(Me.txtPrimaryInvoice.Text)
                        End If
                    Else
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If
                    Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                    Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                    Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                    Me.cboCustomerAddress.SelectedIndex = 0
                    Me.txtPaymentTerm.Text = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                        If Me.CboAddressShipTo.Items.Count > 0 Then
                            Me.CboAddressShipTo.SelectedIndex = 0
                        End If
                        Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.GetData_EFaktur(Me.txtPrimaryCustomer.Text, Me.Txt_E_Faktur_Jalan, Me.Txt_E_Faktur_Blok, _
                        Me.Txt_E_Faktur_No, Me.Txt_E_Faktur_RT, Me.Txt_E_Faktur_RW, Me.Txt_E_Faktur_Kelurahan, Me.Txt_E_Faktur_Kecamatan, _
                        Me.Txt_E_Faktur_Kabupaten, Me.Txt_E_Faktur_Propinsi, Me.Txt_E_Faktur_KodePos, Me.Txt_E_Faktur_NoTelepon)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.TxtSalesPersonCode.Focus()
                        Me.txtPaymentTerm.Text = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.CboAddressShipTo.Items.Clear()
                        Me.TxtIdentifierAddressShipTo.Text = ""
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtNoSeriFakturPajak.Clear()
                        Me.txtCustomerCode.Focus()
                        Me.txtPaymentTerm.Text = "0"
                        Me.Txt_E_Faktur_Jalan.Text = ""
                        Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                        '//TR : Tambahan baru Nik/Pas
                        'Me.llbUpDateNPWP.Visible = True
                        Me.cboNIKPASS.Visible = False
                        Me.cboNIKPASS.TabStop = False
                        Me.txtNoNIKPass.Visible = False
                        Me.txtNoNPWP.Visible = True
                    End If
                End If
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.CboAddressShipTo.Items.Clear()
                Me.TxtIdentifierAddressShipTo.Text = ""
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtNoSeriFakturPajak.Clear()
                Me.txtPaymentTerm.Text = "0"
                Me.Txt_E_Faktur_Jalan.Text = ""
                Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                '//TR : Tambahan baru Nik/Pas
                'Me.llbUpDateNPWP.Visible = True
                Me.cboNIKPASS.Visible = False
                Me.cboNIKPASS.TabStop = False
                Me.txtNoNIKPass.Visible = False
                Me.txtNoNPWP.Visible = True
            End If
            Me.UpdateSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Public Sub CurrencyCode_Leave()
        Dim i As New FrmCariMasterCurrency

        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision()
                        Me.cboCalculatesBy.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
            End If

            Me.IsiCBOCurrencyTax()
            Me.IsiCBOArAkunByTipePerkiraan1()

            Me.SetTaxRate()
            Me.SetCalculateBy()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Public Sub Tax_Leave()
        Try
            If Me.TxtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotalBeforeDP.Text) Then
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, 2)
                Me.HitungVATItem()
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    '2025 Core Tax Tahap 2.2
    Private Sub UbahDPPNilaiLain(ByVal DecDasarPerhitunganPajak As Decimal)
        Try
            'Update SMESS, 2026-10-12
            Me.txtDPPNilaiLain.ReadOnly = True
            Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(DecDasarPerhitunganPajak * 11 / 12, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))

            ''DPP Nilai Lain untuk 04, sementara 11/12, tapi bisa diedit
            'If Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax04 Then
            '    Me.txtDPPNilaiLain.ReadOnly = True
            '    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotalAfterDP.Text) * 11 / 12, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'ElseIf Me.cboTaxtype.SelectedIndex = ClsTransaksiSalesInvoice.EnumTaxType.Tax05 Then
            '    'DPP Nilai Lain untuk 05, sementara 10% dari total DPP, bisa diedit
            '    Me.txtDPPNilaiLain.ReadOnly = True
            '    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotalAfterDP.Text) * 100 / 100, Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'Else
            '    Me.txtDPPNilaiLain.ReadOnly = True
            '    Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(CDec(Me.txtSubTotalAfterDP.Text), Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.2
    Private Sub LlbViewDetailTax_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewDetailTax.LinkClicked
        Try
            Me.ShowInfoPajak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ShowInfoPajak() As Boolean
        Dim ij As New FrmTransaksiSalesInvoice_InfoFakturPajak
        Try
            ij.InvoiceDate = Me.dtpDate.Value
            ij.PrimarySalesInvoice = Me.txtPrimaryInvoice.Text
            ij.UserCode = Me.UserCode
            ij.BackColorAll = Me.BackColorAll
            ij.InvoiceNo = Me.txtInvoiceNo.Text
            '2025 Core Tax Tahap 2.2, update ketika semua type bisa dipilih, tanpa percabangan
            'If Me.cboTaxType.SelectedIndex = 0 Then
            '    ij.VATType = ClsTransaksiSalesInvoice.EnumTaxType.Tax04
            'ElseIf Me.cboTaxType.SelectedIndex = 1 Then
            '    ij.VATType = ClsTransaksiSalesInvoice.EnumTaxType.Tax08
            'End If
            ij.VATType = Me.cboTaxType.SelectedIndex
            If Me.AddInfo <> -1 Then
                ij.PrimaryAddInfo = Me.AddInfo
            End If
            If Me.FacilityStamp <> -1 Then
                ij.PrimaryFacilityStamp = Me.FacilityStamp
            End If
            ij.ShowDialog()
            If ij.Simpan Then
                Me.SPPBDate = ij.dtpSPPBDate.Value
                Me.AddInfo = ij.cboTaxAddInfo.SelectedValue
                Me.FacilityStamp = ij.cboTaxFacilityStamp.SelectedValue
                'Me.TxtSPPBNo.Text = ij.TxtSPPBNo.Text
            End If
            Return ij.Simpan
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cboTaxType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTaxtype.SelectedIndexChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                If cboTaxType.SelectedIndex <> ClsTransaksiSalesInvoice.EnumTaxType.Tax07 And cboTaxType.SelectedIndex <> ClsTransaksiSalesInvoice.EnumTaxType.Tax08 Then
                    Me.LlbViewDetailTax.Enabled = False
                Else
                    Me.LlbViewDetailTax.Enabled = True
                    If Me.ObjInvoice.SetDontCollect1(Me.cboTaxType.SelectedIndex, Me.chkTax.Checked, Me.cboPPNDipungut.SelectedIndex) = True Then
                        Me.cboPPNDipungut.SelectedIndex = 1
                    End If
                End If
                '2025 Core Tax Tahap 2.2, tambahkan hitung total
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtDPPNilaiLain, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtDPPNilaiLain.Enabled = True Then
                    If Me.txtDPPNilaiLain.ReadOnly = True Then
                        Me.txtDPPNilaiLain.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtDPPNilaiLain.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDPPNilaiLain.KeyDown
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyDown(Me.txtDPPNilaiLain, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDPPNilaiLain.KeyUp
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_KeyUp(Me.txtDPPNilaiLain, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(Me.txtDPPNilaiLain.Text, Me.DefaultPrecision)
                'Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtDPPNilaiLain_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDPPNilaiLain.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtDPPNilaiLain.Enabled = True Then
                    If Me.txtDPPNilaiLain.ReadOnly = True Then
                        Me.txtDPPNilaiLain.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtDPPNilaiLain.BackColor = Me.txtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TaxCodeDefaultDariCustomer(ByVal vCode As String)
        Try
            If vCode.Length > 3 Then
                vCode = vCode.Substring(0, 3)
            End If

            Select Case vCode
                Case "01", "010"
                    Me.cboTaxtype.SelectedIndex = 0
                Case "02", "020"
                    Me.cboTaxtype.SelectedIndex = 1
                Case "03", "030"
                    Me.cboTaxtype.SelectedIndex = 2
                Case "04", "040"
                    Me.cboTaxtype.SelectedIndex = 3
                Case "05", "050"
                    Me.cboTaxtype.SelectedIndex = 4
                Case "06", "060"
                    Me.cboTaxtype.SelectedIndex = 5
                Case "07", "070"
                    Me.cboTaxtype.SelectedIndex = 6
                Case "08", "080"
                    Me.cboTaxtype.SelectedIndex = 7
                Case "09", "090"
                    Me.cboTaxtype.SelectedIndex = 8
                Case "10", "100"
                    Me.cboTaxtype.SelectedIndex = 9
                Case Else
                    Me.cboTaxtype.SelectedIndex = 0
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '2025 Core Tax Tahap 2.2 end

    Private Sub txtRevisedInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.Click
        Try '//#2732
            Me.ObjCD.Number_Click(Me.txtRevisedInvoice)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.Leave
        Try '//#2732
            If Me.txtRevisedInvoice.Text = "" Or _
            Me.txtRevisedInvoice.Text = "0" Then
                Me.txtRevisedInvoice.Text = "00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.Enter
        Try '//#2732
            Me.ObjCD.Number_Enter(txtRevisedInvoice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.GotFocus
        Try '//#2732
            Me.txtRevisedInvoice.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevisedInvoice.LostFocus
        Try '//#2732
            Me.txtRevisedInvoice.BackColor = Me.txtPrimaryInvoice.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtRevisedInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRevisedInvoice.KeyPress
        Dim txt As TextBox = CType(sender, TextBox)
        Dim hasil As String = txt.Text & e.KeyChar
        Try '//#2732
            ' Izinkan backspace
            If Char.IsControl(e.KeyChar) Then Exit Sub

            ' Hanya angka
            If Not Char.IsDigit(e.KeyChar) Then
                e.Handled = True
                Exit Sub
            End If
            ' Maksimal 2 digit dan hanya boleh 00 atau 01
            If hasil <> "0" AndAlso hasil <> "00" AndAlso hasil <> "01" Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class