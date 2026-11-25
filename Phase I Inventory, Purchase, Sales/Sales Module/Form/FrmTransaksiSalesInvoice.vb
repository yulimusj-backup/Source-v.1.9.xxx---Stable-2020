Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiSalesInvoice
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    'for quick input
    Public QuickInput As Boolean = False
    Public StopQuickInput As Boolean = False
    Public TblQuickInput As New DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceDataTable
    Public DirectlySave As Boolean = False
    Public DirectlyPrint As Boolean = False
    Public PathPrintCustom As String = ""
    Public SetDtpDate As Date = #1/1/1950#

    Dim baru As Boolean
    Dim ObjInvoice As New ClsTransaksiSalesInvoice
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCustomer_sProductCode As New clsTransaksiSettingcustomerProductCodeNo
    Dim objDO As New ClsTransaksiDO
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjInvoiceItem As New ClsTransaksiSalesInvoiceItem
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjSR As New ClsTransaksiSalesReturn
    Dim ObjSRItem As New ClsTransaksiSalesReturnItem
    Dim ObjStockOutSales As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjLicense As New ClsLicense
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjTaxInvoiceSetting As New ClsTransaksiSalesInvoice_SettingTaxInvoice
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjSalesPerson As New Core_Module.ClsMasterSalesPerson
    Dim BolehEditHarga As Boolean = False

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
    Public MetodeHitungInclude As EnumHitungInclude = EnumHitungInclude.NettPriceInclude

    'komponen di bawah ini digunakan untuk mengisikan data perhitungan do-so yang lambat
    Dim DTADO2 As New DseTransaksiDOTableAdapters.tbltrans_do_item8TableAdapter
    Dim TBLDO2 As New DseTransaksiDO.tbltrans_do_item8DataTable
    Dim DTADO3 As New DseTransaksiDOTableAdapters.tbltrans_so_item1TableAdapter
    Dim TBLDO3 As New DseTransaksiDO.tbltrans_so_item1DataTable
    Dim ListPrimarySODummy As New List(Of Double)

    'komponen ini digunakan untuk mengisikan data perhitungan item do pada btn Input
    Dim DTA_DoItem1 As New DseTransaksiDOTableAdapters.DOTableAdapter
    Dim TBL_DoItem1 As New DseTransaksiDO.DODataTable
    Dim DTA_SoItem1 As New DseTransaksiDOTableAdapters.SOTableAdapter
    Dim TBL_SoItem1 As New DseTransaksiDO.SODataTable

    'Perhatikan!!! KALO PAKE METODE UNIT PRICE INCLUDE ATAU SO DILAPORAN SALES ANALISIS 
    ' DAN SALES STATISTIK AKAN NGACO PERHITUNGANNYA


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

    Dim ObjAkun As New Core_Module.ClsMasterAkun1

    Dim OptionNIKPASS As clsMasterCustomer.EnumOptionNikPass = clsMasterCustomer.EnumOptionNikPass.ByNIK
    Dim bufferNikPass As String = ""

    '2025 Core Tax Tahap 2.1
    Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
    '2025 Core Tax Tahap 2.2
    Dim SPPBDate As Date = #1/1/1950#
    Dim AddInfo As Double = -1
    Dim FacilityStamp As Double = -1

    Private Sub IsiCBOArAkunByTipePerkiraan1()
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

    Private Sub txtPrimaryCurrency_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrency.TextChanged
        Try
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
            'Me.txtPrimaryAccount.Visible = True
            If IsNumeric(Me.cboPrimaryCurrTax.Text) AndAlso Me.cboPrimaryCurrTax.Text > -1 Then
                Me.txtPrimaryCurrTax.Text = Me.cboPrimaryCurrTax.Text
                DeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)

                If IsNumeric(Me.TxtDutyStamp.Text) Then
                    Me.TxtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.TxtDutyStamp.Text, DeffPrec)
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

    Private Sub txtPrimaryCurrTax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrimaryCurrTax.TextChanged
        Dim PrimaryCurrTax As Double = -1

        Try
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

    Private Sub Hitung_Total()
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0
        Dim exchangRate As Decimal = 0
        Dim TBL1 As New DseTransaksiDO.DataTable10DataTable
        Dim DTA1 As New DseTransaksiDOTableAdapters.DataTable10TableAdapter

        Dim DummyPrimaryDO As Double = -1
        Dim PrimaryDO As Double = -1
        Dim PrimarySO, PrimaryProduct As Double
        Dim PriceExclude As Boolean = True
        Dim Unit_Price_SO As Decimal = 0
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ED As Date = #1/1/1950#
        Dim DefPrec As Integer = 0
        Dim dummy_index As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Total_MiddlePoint()
                Exit Sub
            End If

            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            dummy_index = 0

            If Me.dgvItem.Rows.Count > 0 Then
                'hitung sub total
                For Each row In Me.dgvItem.Rows
                    dummy_index += 1
                    'PrimaryDO = row.Cells(Me.ColPrimaryDO.Index).Value
                    'PrimarySO = row.Cells(Me.ColPrimarySO.Index).Value
                    'PrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                    'If DummyPrimaryDO <> PrimaryDO Then
                    '    'DTA1.Fill_ByPrimaryDO(TBL1, PrimaryDO)
                    '    Me.IsiItemDO1(TBL1, PrimaryDO)
                    '    DummyPrimaryDO = PrimaryDO
                    'End If
                    'PriceExclude = Me.Get_Price_Exclude(PrimarySO, PrimaryProduct, TBL1)
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
                    Dim SubTotal1 As Decimal = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    Dim SecQty As Decimal
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
                    Dim JmlReturn As Decimal = 0

                    If Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                        If SecQty = -1 Then
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                JmlReturn = Me.GetJmlDGVSR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                row.Cells(Me.ColPrimaryDO.Index).Value, _
                                SecQty, row.Cells(Me.ColPrimarySO.Index).Value, _
                                row.Cells(Me.ColLotNumber.Index).Value, ED)
                            Else
                                JmlReturn = Me.GetJmlDGVSR_Default(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                row.Cells(Me.ColPrimaryDO.Index).Value, _
                                SecQty, row.Cells(Me.ColPrimarySO.Index).Value, _
                                row.Cells(Me.ColLotNumber.Index).Value, ED)
                            End If
                            If JmlReturn = 0 Then
                                SubTotal1 = CDec(row.Cells(Me.ColAmount.Index).Value)
                            Else
                                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                    SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                    row.Cells(Me.ColNettPrice.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                                Else
                                    SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                                    row.Cells(Me.ColNettPrice.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                                End If
                            End If
                        Else
                            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                                JmlReturn = Me.GetJmlDGVSR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                row.Cells(Me.ColPrimaryDO.Index).Value, _
                                SecQty, row.Cells(Me.ColPrimarySO.Index).Value, _
                                row.Cells(Me.ColLotNumber.Index).Value, ED)
                                SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            Else
                                JmlReturn = Me.GetJmlDGVSR_Default(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                                row.Cells(Me.ColPrimaryDO.Index).Value, _
                                SecQty, row.Cells(Me.ColPrimarySO.Index).Value, _
                                row.Cells(Me.ColLotNumber.Index).Value, ED)
                                SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    End If
                    SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
                Next

                If IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso Me.txtPrimaryCurrency.Text = 1 Then
                    'kalo IDR rounddown ke nol rupiah, sesuai dengan ketentuan perpajakan yang baru
                    Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(SubTotal, DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                Else
                    Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
                End If

                Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(Me.TxtShipment.Text, DefaultPrecision)
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.dgvDP.Rows
                    Me.txtDownPayment.Text = CDec(Me.txtDownPayment.Text) + CDec(row.Cells(3).Value)
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
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                'Else
                '    'shipment without tax
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                'End If

                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                    DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Else
                    DefPrecTax = DefaultPrecision
                End If


                If Me.chkTax.Checked = True Then
                    'If Me.cboPPNDipungut.SelectedIndex = 0 Then
                    'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                    If Me.txtPrimaryCurrency.Text = "1" Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                        DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                    Else
                        If Me.txtPrimaryCurrTax.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                                  DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                        Else

                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                                   DecDasarPerhitunganPajak / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Down), DefPrecTax)
                        End If
                    End If

                    Me.txtTaxAmount1.Text = Me.txtTaxAmount.Text 'new tax amount1
                    If Me.cboPPNDipungut.SelectedIndex = 1 Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                    End If
                    'Else
                    '    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                    'End If
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
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text)) 'new tax amount1
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
            End If


            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If IsNumeric(Me.txtPrimaryCurrency.Text) Then
                If Me.txtPrimaryCurrency.Text = 1 Then
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
            End If


        Catch ex As Exception
            MsgBox(ex.Message + dummy_index.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub Hitung_Total_MiddlePoint()
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0
        Dim exchangRate As Decimal = 0
        Dim TBL1 As New DseTransaksiDO.DataTable10DataTable
        Dim DTA1 As New DseTransaksiDOTableAdapters.DataTable10TableAdapter

        Dim DummyPrimaryDO As Double = -1
        Dim PrimaryDO As Double = -1
        Dim PrimarySO, PrimaryProduct As Double
        Dim PriceExclude As Boolean = True
        Dim Unit_Price_SO As Decimal = 0
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ED As Date = #1/1/1950#
        Dim DefPrec As Integer = 0
        Dim dummy_index As Integer = 0
        Dim DecDasarPerhitunganPajak As Decimal = 0
        Try
            DefPrec = Me.ObjCurrency.GetDecimal_DefaultCurrency
            dummy_index = 0

            If Me.dgvItem.Rows.Count > 0 Then
                'hitung sub total
                For Each row In Me.dgvItem.Rows
                    dummy_index += 1
                    'PrimaryDO = row.Cells(Me.ColPrimaryDO.Index).Value
                    'PrimarySO = row.Cells(Me.ColPrimarySO.Index).Value
                    'PrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value
                    'If DummyPrimaryDO <> PrimaryDO Then
                    '    'DTA1.Fill_ByPrimaryDO(TBL1, PrimaryDO)
                    '    Me.IsiItemDO1(TBL1, PrimaryDO)
                    '    DummyPrimaryDO = PrimaryDO
                    'End If
                    'PriceExclude = Me.Get_Price_Exclude(PrimarySO, PrimaryProduct, TBL1)
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
                    Dim SubTotal1 As Decimal = 0
                    DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                    Dim SecQty As Decimal
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
                    Dim JmlReturn As Decimal = Me.GetJmlDGVSR(row.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                    row.Cells(Me.ColPrimaryDO.Index).Value, _
                    SecQty, row.Cells(Me.ColPrimarySO.Index).Value, _
                    row.Cells(Me.ColLotNumber.Index).Value, ED)

                    'If Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                    '    If SecQty = -1 Then
                    '        SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                    '        row.Cells(Me.ColNettPrice.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '    Else
                    '        SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                    '        row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    '    End If
                    'End If
                    If Me.dtpDate.Value >= #10/9/2025# And Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = False Then
                        Me.HitungVATItemWithReturn(row, exchangRate, JmlReturn)
                        If row.Cells(Me.ColPriceCond.Index).Value = "Include" Then
                            Dim nettprice As Decimal = (row.Cells(Me.ColAmountInclude.Index).Value * (100 / (CDbl(Me.txtTax.Text) + 100)) / row.Cells(Me.ColTotalQtyMulti.Index).Value)
                            If SecQty = -1 Then
                                SubTotal1 = (row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                nettprice * exchangRate
                            Else
                                SubTotal1 = (row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                nettprice * row.Cells(Me.ColSecQty.Index).Value * exchangRate
                            End If
                        Else
                            If SecQty = -1 Then
                                SubTotal1 = (row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                row.Cells(Me.ColNettPrice.Index).Value * exchangRate
                            Else
                                SubTotal1 = (row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate
                            End If
                        End If
                    Else
                        If Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                            If SecQty = -1 Then
                                If JmlReturn = 0 Then
                                    SubTotal1 = CDec(row.Cells(Me.ColAmount.Index).Value)
                                Else
                                    SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                    row.Cells(Me.ColNettPrice.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                                End If

                            Else
                                SubTotal1 = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                                row.Cells(Me.ColNettPrice.Index).Value * row.Cells(Me.ColSecQty.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                            End If
                        End If
                    End If
                    'SubTotal += Me.ObjCD.SetDecimal(SubTotal1, Me.DefaultPrecision)
                    SubTotal += SubTotal1
                Next

                If IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso Me.txtPrimaryCurrency.Text = 1 Then
                    'kalo IDR rounddown ke nol rupiah, sesuai dengan ketentuan perpajakan yang baru
                    Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(SubTotal, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefaultPrecision) 'Me.ObjCD.DigitGrouping(SubTotal, DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                Else
                    Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
                End If

                Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(Me.TxtShipment.Text, DefaultPrecision)
                'hitung DP
                Me.txtDownPayment.Text = "0"
                For Each row In Me.dgvDP.Rows
                    Me.txtDownPayment.Text = CDec(Me.txtDownPayment.Text) + CDec(row.Cells(3).Value)
                Next
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtDownPayment.Text, DefaultPrecision))

                '------------------------------------------------------------------------------------------------------------------------'
                '------------------------------------------------------------------------------------------------------------------------'
                'diganti berdasarkan complain error ci fitri no 50 - error sales smess
                'Keputusan : with or without tax, shipment tetap dihitung, tapi saat perhitungan pajak, shipment tidak dihitung
                '2016/07/25
                '------------------------------------------------------------------------------------------------------------------------'

                If IsNumeric(Me.txtPrimaryCurrency.Text) AndAlso Me.txtPrimaryCurrency.Text = 1 Then
                    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.ObjCD.DigitGrouping(SubTotal, 0, ClsCekData.EnumTipeRounding.Middlepoint) + _
                CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                    'If Me.cboShipmentTax.SelectedIndex = 0 Then
                    '    DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.ObjCD.DigitGrouping(SubTotal, 0, ClsCekData.EnumTipeRounding.Middlepoint) + _
                    '    CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                    'Else
                    '    DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.ObjCD.DigitGrouping(SubTotal, 0, ClsCekData.EnumTipeRounding.Middlepoint) - _
                    '    CDec(Me.txtDownPayment.Text), DefaultPrecision))
                    'End If
                    If Me.cboShipmentTax.SelectedIndex = 0 Then
                        DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), 4))
                    Else
                        DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), 4))
                    End If
                Else
                    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                    If Me.cboShipmentTax.SelectedIndex = 0 Then
                        DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), 4))
                    Else
                        DecDasarPerhitunganPajak = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), 4))
                    End If
                End If

                'If Me.cboShipmentTax.SelectedIndex = 0 Then
                '    'shipment with tax
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal + CDec(Me.TxtShipment.Text) - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                'Else
                '    'shipment without tax
                '    Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal - CDec(Me.txtDownPayment.Text), DefaultPrecision))
                'End If

                'hitung pajak
                Dim DefPrecTax As Decimal = 0
                If IsNumeric(Me.txtPrimaryCurrTax.Text) AndAlso CDec(Me.txtPrimaryCurrTax.Text) > -1 Then
                    DefPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Else
                    DefPrecTax = DefaultPrecision
                End If


                If Me.chkTax.Checked = True Then
                    If Me.cboPPNDipungut.SelectedIndex = 0 Then
                        'kalo dibayar dalam rupiah maka pajak tidak perlu dikalikan dengan exchange rate
                        If Me.txtPrimaryCurrency.Text = "1" Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                            DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                        Else
                            If Me.txtPrimaryCurrTax.Text = "1" Then
                                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTaxExchangeRate.Text) * CDec(Me.txtTax.Text) * _
                                DecDasarPerhitunganPajak / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                            Else

                                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(CDec(Me.txtTax.Text) * _
                                DecDasarPerhitunganPajak / 100, DefPrecTax, ClsCekData.EnumTipeRounding.Middlepoint), DefPrecTax)
                            End If
                        End If
                    Else
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                    End If
                Else
                    Me.txtTax.Text = "0.00"
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", DefPrecTax)
                End If
                Me.txtTaxAmount1.Text = Me.txtTaxAmount.Text 'new tax amount1
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, DefPrecTax)
            Else
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.dgvDP.Rows.Clear()
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping("0", Me.ObjCurrency.GetDefaultPrecission(txtPrimaryCurrTax.Text))
                DecDasarPerhitunganPajak = 0
            End If

            '2025 Core Tax Tahap 2.2
            'DPP Nilai Lain untuk 04, sementara 11/12, tapi bisa diedit
            Me.UbahDPPNilaiLain(DecDasarPerhitunganPajak)

            Dim decTotal As Decimal = 0
            'hitung total = sub total + taxj 
            If IsNumeric(Me.txtPrimaryCurrency.Text) Then
                If Me.txtPrimaryCurrency.Text = 1 Then
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
            End If


        Catch ex As Exception
            MsgBox(ex.Message + dummy_index.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

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

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/2050#
        Dim TanggalPerbandingan As Date
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

        Try
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
            KoleksiNamaData.Add("Invoice number")
            KoleksiTextBox.Add(Me.txtInvoiceNo)
            KoleksiNamaData.Add("Customer code")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            KoleksiNamaData.Add("Currency")
            KoleksiTextBox.Add(Me.txtCurrencyCode)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

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
                        Me.dgvItemSR, Me.dgvExchageRates, Me)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
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

            'If Me.dgvDPDummy.Rows.Count > 0 AndAlso Not String.IsNullOrEmpty(Me.dgvDPDummy.Rows(0).Cells(0).Value) Then
            Me.ObjInvoice.IsiDgvDP(ij, Me.dgvDPDummy, Me.txtPrimaryCurrency.Text, Me.txtPrimaryInvoice.Text)
            If Me.dgvDPDummy.Rows.Count > Me.dgvDP.Rows.Count Then
                '2025 Core Tax Tahap 2.1, cek SO Balance nol semua, jika true, harus diisi DP nya
                If Me.ObjInvoice.CekSOBNol(ij) Then
                    MsgBox("There are down payment refering to this invoice and cannot continue", MsgBoxStyle.Information, "Information")
                    Me.llbAddDP.Focus()
                    Return False
                End If

                If MsgBox("There are down payment refering to this invoice, continue and ignore the downpayment?", _
                MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Return False
                End If
            End If
            'End If

            'cek berkaitan dengan tanggal closing
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan tanggal stock out dan tanggal DO
            For Each row In Me.dgvItem.Rows
                'cek stock out]
                If ListNoStockOut.Contains(row.Cells(Me.ColStockOutNo.Index).Value) = False Then
                    TanggalPerbandingan = Me.ObjStockOutSales.GetTanggalTransaksi_ByNoTransaksi(row.Cells(Me.ColStockOutNo.Index).Value)
                    If Format(TanggalPerbandingan.Date, "MMM yyyy") <> Format(dtpDate.Value.Date, "MMM yyyy") Then
                        If StockOutStr = "" Then
                            StockOutStr = row.Cells(Me.ColStockOutNo.Index).Value
                        Else
                            StockOutStr = StockOutStr + ", " + row.Cells(Me.ColStockOutNo.Index).Value
                        End If
                    End If
                    ListNoStockOut.Add(row.Cells(Me.ColStockOutNo.Index).Value)
                End If
            Next
            If StockOutStr <> "" Then
                MsgBox("The stock out sales must be dated on the same month as invoice date. Please change your invoice date or fix following stock out data: " _
                + StockOutStr, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
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
                If Me.txtNoFakturPajak.Text = "" Then
                    Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    If Me.txtNoFakturPajak.Text = "" Then
                        MsgBox("There are no new tax invoice number available. Please specify  new tax invoice number in your global setting.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    'MsgBox("Please enter tax invoice number", MsgBoxStyle.Information, "Information")
                    'Return False

                End If
            End If

            'berkaitan dengan tax exchange rate
            If Me.txtPrimaryCurrency.Text = "1" Then
                If CDec(Me.txtTaxExchangeRate.Text) <> 1 Then
                    MsgBox("For invoice with " + Me.ObjCurrency.GetCurrencyCode(1) + " currency, the tax exchange rate value must be 1.", _
                    MsgBoxStyle.Information, "Information")
                    Me.txtTaxExchangeRate.Focus()
                    Return False
                End If
            Else
                'jika bukan IDR dan tax ratenya bernilai 1 maka berikan peringatan
                If CDec(Me.txtTaxExchangeRate.Text) = 1 Then
                    MsgBox("This invoice use non-" + Me.ObjCurrency.GetCurrencyCode(1) + " currency and the tax exchange rate value = 1. This is wrong data input.", _
                    MsgBoxStyle.Information, "Information")
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
                        If CDec(row_dgv.Cells(Me.ColExchangeRate.Index).Value) = 1 Then
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
                        If CDec(row_dgv.Cells(Me.ColExchangeRate.Index).Value) <> 1 Then
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
                Dim PrimaryCurr As Double = -1
                If row_dgv.Cells(Me.ColDPTipeDP.Index).Value = ClsTransaksiSalesInvoice.EnumTipeDP.InvDP Then
                    PrimaryCurr = Me.ObjInvoice.Get_Primary_Currency(row_dgv.Cells(Me.ColPrimaryInvDP.Index).Value)
                Else
                    PrimaryCurr = Me.ObjInvoice.GetPrimaryCurrency_ManualBalance(row_dgv.Cells(Me.ColPrimaryInvDP.Index).Value)
                End If
                If PrimaryCurr <> Me.txtPrimaryCurrency.Text Then
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
            Me.IsiPaymentTerm()



            'Item DP yang dipilih harus memiliki account payable dan currency tax yang sama
            If dgvDP.Rows.Count > 0 AndAlso Not IsDBNull(dgvDP.Rows(0).Cells(0).Value) Then
                Dim strMsg As String = ""
                Dim strMsg2 As String = ""
                For Each row In Me.dgvDP.Rows
                    Dim PrimaryInvDP As Double = -1
                    Dim PrimaryCurrTax As Double = -1
                    Dim PrimaryAccount As Double = -1
                    PrimaryInvDP = row.Cells(Me.ColPrimaryInvDP.Index).Value
                    If row.Cells(Me.ColDPTipeDP.Index).Value = ClsTransaksiSalesInvoice.EnumTipeDP.InvDP Then
                        PrimaryCurrTax = Me.ObjInvoice.GetPrimaryCurrency(PrimaryInvDP)
                    Else
                        PrimaryCurrTax = Me.ObjInvoice.GetPrimaryCurrency_ManualBalance(row.Cells(Me.ColPrimaryInvDP.Index).Value)
                    End If
                    PrimaryAccount = Me.ObjInvoice.GetPrimaryAccountAR(PrimaryInvDP)

                    If Me.txtPrimaryCurrency.Text <> PrimaryCurrTax Then 'or Me.txtPrimaryAccount.Text <> PrimaryAccount OrElse  Then
                        If strMsg = "" Then
                            strMsg &= row.Cells(Me.CollDPInvNo.Index).Value & " use currency: " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
                        Else
                            strMsg &= vbNewLine & row.Cells(Me.CollDPInvNo.Index).Value & _
                                      " use currency: " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
                        End If
                    End If
                Next

                If strMsg <> "" Then
                    MsgBox("Invoice DP No. " & strMsg & ",  you must use similar currency for this transaction.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            ''Item DP yang dipilih harus memiliki account payable dan currency tax yang sama
            'If Me.dgvSR.Rows.Count > 0 AndAlso Not IsDBNull(dgvSR.Rows(0).Cells(0).Value) Then
            '    Dim strMsg As String = ""
            '    Dim strMsg2 As String = ""
            '    For Each row In Me.dgvSR.Rows
            '        Dim PrimaryPR As Double = -1
            '        Dim PrimaryCurrTax As Double = -1
            '        Dim PrimaryAccount As Double = -1
            '        PrimaryPR = row.Cells(Me.ColPrimarySR_SR.Index).Value
            '        PrimaryCurrTax = Me.ObjSR.GetPrimaryCurrTax(PrimaryPR)
            '        PrimaryAccount = Me.ObjSR.GetPrimaryAccountAP(PrimaryPR)

            '        If Me.txtPrimaryAccount.Text <> PrimaryAccount OrElse Me.txtPrimaryCurrTax.Text <> PrimaryCurrTax Then
            '            If strMsg = "" Then
            '                strMsg &= row.Cells(Me.ColSRNo.Index).Value & " use account receivable " & _
            '                          ObjAkun.GetAccountCode(PrimaryAccount) & "-" & ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAccount) & _
            '                          " and currency tax " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
            '            Else
            '                strMsg &= vbNewLine & row.Cells(Me.ColSRNo.Index).Value & _
            '                          " use account receivable " & ObjAkun.GetAccountCode(PrimaryAccount) & _
            '                          "-" & ObjAkun.GetDescription_ByPrimaryAcc(PrimaryAccount) & _
            '                          " and currency tax " & Me.ObjCurrency.GetCurrencyCode(PrimaryCurrTax)
            '            End If
            '        End If
            '    Next

            '    If strMsg <> "" Then
            '        MsgBox("Sales return No. " & strMsg & vbNewLine & " you just change account receivable and currency " & _
            '        "tax with similiar data for journal needed.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If


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
            If CDec(Me.txtTaxAmount1.Text) > 999999999999 Then 'new tax amount1
                MsgBox("Maximum VAT value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'jika ada shipment cost, pastikan akun shipment costnya ada
            If CDec(Me.TxtShipment.Text) <> 0 Then
                If Me.ObjInvoice.GetPrimaryAkunPendapatanOngkosKirim = -1 Then
                    MsgBox("Please specify income from shipment account in global setting data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
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

            'nilai dp, tidak boleh melebih nilai invoice
            If CDec(Me.txtDownPayment.Text) > CDec(Me.txtSubTotalBeforeDP.Text) Then
                MsgBox("Maximum down payment amount is " + Me.txtSubTotalBeforeDP.Text + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '2025 Core Tax Tahap 2.2
            'kalo pake pajak, maka no seri faktur pajak harus diisikan
            'Me.txtNoSeriFakturPajak.Text = Me.txtNoSeriFakturPajak.Text.Trim
            'If Me.chkTax.Checked = True And Me.txtNoSeriFakturPajak.Text = "" Then
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

            'cek akun duty stamp
            If CDec(Me.txtDutyStamp.Text) > 0 Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify dutystamp expense account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'cek akun income from shipment
            If CDec(Me.TxtShipment.Text) > 0 Then
                Dummy = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, ClsTransaksiSetting.EnumNamaParameter.Akun_ShippingIncome)
                If Dummy = "" Or Dummy = "-1" Then
                    MsgBox("Please specify income from shipment account.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'pastikan semua sales return revisi invoice dibuatkan pada bulan yg sama pada tgl invoice
            Dummy = ""
            For Each row In Me.dgvItemSR.Rows
                If Format(Me.dtpDate.Value, "MMM yyyy") <> Format(Me.ObjSR.GetTanggalTransaksi(row.Cells(Me.ColPrimarySR.Index).Value), "MMM yyyy") Then
                    If Dummy = "" Then
                        Dummy = row.Cells(Me.ColSRNo.Index).Value
                    Else
                        Dummy = Dummy + ", " + row.Cells(Me.ColSRNo.Index).Value
                    End If
                End If
            Next
            If Dummy <> "" Then
                MsgBox("The following sales return must be dated on: " + Format(Me.dtpDate.Value, "MMM yyyy") + ": " + _
                Dummy + ".", MsgBoxStyle.Information, "Information")
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
                If Me.chkTax.Checked = True Then
                    'for quick input
                    If Me.QuickInput = True Then
                        Me.Show()
                    End If
                    If MsgBox("The following code no has amount = 0, this will cause exception in e-Faktur import process. Continue save?" + vbCrLf + Dummy, _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If

            'kalo IDR, pastikan tidak ada decimal
            If Me.txtPrimaryCurrency.Text = "1" Then
                If Me.ObjCD.GetJmlPrecision(Me.txtSubTotalAfterDP.Text) > 0 Then
                    If Me.chkTax.Checked = True Then
                        'for quick input
                        If Me.QuickInput = True Then
                            Me.Show()
                        End If
                        '2025 Core Tax Tahap 2.1
                        If Me.dtpDate.Value < #1/1/2025# Then
                            If MsgBox("The invoice is in " + Me.ObjCurrency.GetCurrencyCode(1) + _
                            " currency and sub total amount after downpayment is in decimal value, this will cause failure in e-Faktur data import. Continue save?", _
                            MsgBoxStyle.YesNo, "Confirm") _
                            <> MsgBoxResult.Yes Then
                                Return False
                            End If
                        End If
                    End If
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
            '    Me.txtNoSeriFakturPajak.Text <> "011" And _
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
            'End If



            '************************************************************************************************************************************
            'PERUBAHAN INSTRUKSI KO IWAN TGL 31 JULI 2016 
            'JIKA 1 SO sama dengan 2 do yg berbeda, lalu buat 1 DP, saat add DP harusnya cm 1 invoice yg bisa add
            '************************************************************************************************************************************
            If Me.ObjInvoice.ApakahAdaDPSudahDipakai(Me.txtPrimaryInvoice.Text, True, Me) = True Then
                Me.llbAddDP.Focus()
                Return False
            End If

            '************************************************************************************************************************************
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

            'terkait ship to
            Dim listPrimarySO As New List(Of Double)
            Dim PrimaryCustShipToTemporary As Double = -1
            Dim CountList As Integer = 0
            For Each rd As DataGridViewRow In Me.dgvItem.Rows
                If listPrimarySO.Contains(rd.Cells(ColPrimarySO.Index).Value) = False Then
                    listPrimarySO.Add(rd.Cells(ColPrimarySO.Index).Value)
                End If
            Next

            Dim DtaSOAddress As New DseTransaksiSOTableAdapters.tbltrans_so3TableAdapter
            Dim TblSOAddress As New DseTransaksiSO.tbltrans_so3DataTable
            Dim RowSOAddress As DseTransaksiSO.tbltrans_so3Row
            DtaSOAddress.Connection.ConnectionString = Core_Module.GetConstring
            For Each PrimarySO As Double In listPrimarySO
                CountList += 1
                If DtaSOAddress.Fill(TblSOAddress, PrimarySO) > 0 Then
                    RowSOAddress = TblSOAddress.Rows(0)
                    With RowSOAddress
                        If CountList = 1 Then
                            PrimaryCustShipToTemporary = .primary_customer_ship_to
                        Else
                            If PrimaryCustShipToTemporary <> .primary_customer_ship_to Then
                                MsgBox("You have to input DO data that has the same customer ship to.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    End With
                End If
            Next
            '-----------------------

            'Me.txtNoNPWP.Text = Me.txtNoNPWP.Text.Trim
            'If Me.ObjCustomer.ApakahFormatNPWPBenar(Me.txtNoNPWP.Text, True) = False Then
            '    Return False
            'End If
            'Comment: Perubahan NPWP 16 Digit
            If chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                '2025 Core Tax Tahap 2.1
                If ObjSetNPWP.SetNPWP_New(Me.dtpDate.Value.Date, Me.txtPrimaryCustomer.Text, Me.txtNoNPWP.Text, _
                ClsTrans_SettingNPWPNo.EnumFromMaster.Customer, Me.txtNoNIKPass.Text, frmLaporanVAT2025_EditCustomer.EnumTransactionType.TransactionSales, _
                Me.txtPrimaryAccount.Text, Me.UserCode) = False Then
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

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim UseTaxSetting, TaxPercentSetting As String

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.SedangRevisi = False
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TabControl1.SelectedIndex = 0
            'Me.dtpDate.Focus()
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
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
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

            Me.chkShowCustCode.Checked = False
            Me.SetTaxRate()
            Me.SetTxtExchangeRate()

            UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.ShipmentTax)
            If UseTaxSetting = "" Or UseTaxSetting = "0" Then
                Me.cboShipmentTax.SelectedIndex = 1
            Else
                Me.cboShipmentTax.SelectedIndex = 0
            End If

            '2025 Core Tax Tahap 2.2
            Me.cboTaxtype.SelectedIndex = Me.ObjInvoice.getLastTaxType()
            Me.txtNoFakturPajak.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage4.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
            ModCekData.KosongkanControl(Me.gbo_EFaktur.Controls, True, False, False)

            Me.txtPrimaryAccount.Text = "-1"
            Me.txtPrimaryCurrTax.Text = "-1"
            Me.txtPrimaryInvoice.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryCustomer.Text = "-1"
            Me.txtRevisiNo.Text = "0"
            Me.txtSubTotalBeforeDP.Text = "0"
            Me.txtOtherCharge.Text = "0"
            Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
            Me.txtTaxExchangeRate.Text = "1.00"
            Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision_DefCur)
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur)
            Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur) 'new tax amount1
            Me.TxtTotal.Text = "0"
            Me.txtDownPayment.Text = "0"
            Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision_DefCur)
            Me.txtSubTotalAfterDP.Text = "0"
            Me.txtExchangeRateToUSD.Text = "1"
            Me.txtExhangeRateToIDR.Text = "1"
            Me.txtPaymentTerm.Text = "0"
            Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping("1", Core_Module.DefPrecExchangeRate)

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

            Me.TxtNamaSigned.Clear()
            Me.TxtJabatanSigned.Clear()

            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.dgvDP.Rows.Clear()
            Me.dgvExchageRates.Rows.Clear()
            Me.dgvItemSR.Rows.Clear()
            Me.dgvSR.Rows.Clear()

      
            Me.cboPrimaryAccount.DataSource = Nothing
            Me.cboPrimaryCurrTax.DataSource = Nothing
            Me.cboAccount.DataSource = Nothing
            Me.cboCurrTax.DataSource = Nothing

            Me.cboPrimaryAccount.Text = ""
            Me.cboAccount.Text = ""
            Me.cboCurrTax.Text = ""

            Me.cboPrimaryAccount.Items.Clear()
            Me.cboPrimaryCurrTax.Items.Clear()
            Me.cboAccount.Items.Clear()
            Me.cboCurrTax.Items.Clear()

            Me.TBLHeaderCost.Rows.Clear()
            Me.TBLitemCost.Rows.Clear()

            Me.chkTax.Checked = True
            Me.cboCustomerAddress.DataSource = Nothing
            Me.cboCustomerAddress.Text = ""
            Me.cboCustomerAddress.Items.Clear()
            PrimaryCustomerShipTo = -1
            IdentifierAddressCustomerShipTo = ""

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
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.txtPrimaryInvoice.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.txtPrimaryInvoice.BackColor, False)

            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoice_EditBillToAddress") = False Then
                Me.cboCustomerAddress.Enabled = False
            End If

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.dgvDP.AllowUserToDeleteRows = True
            Me.dgvItemSR.AllowUserToDeleteRows = True
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvExchageRates.ReadOnly = False

            'Me.llbAddress.Enabled = True
            Me.llbAddDP.Enabled = True
            Me.llbAddSR.Enabled = True
            'Me.llbEditNoFakturPajak.Enabled = True
            'Me.llbEditNoInvoice.Enabled = True
            Me.llbUpDateNPWP.Enabled = True
            'Me.lblOtherCharge.Enabled = True

            If Me.dgvItem.RowCount > 0 Then
                Me.txtCustomerCode.ReadOnly = True
                Me.btnFindCustomer.Enabled = False
                ''    Me.txtCustomerCode.BackColor = Me.TxtTotal.BackColor
            Else
                Me.txtCustomerCode.ReadOnly = False
                Me.btnFindCustomer.Enabled = True
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

            Me.cboAccount.Enabled = True
            Me.cboCurrTax.Enabled = True

            If Me.txtPrimaryCurrTax.Text = -1 Then
                Me.IsiCBOArAkunByTipePerkiraan1()
                Me.IsiCBOCurrencyTax()
            End If


            'set defaul precision untuk idr
            Me.DefaultPrecision_DefCur = Me.ObjCurrency.GetDecimal_DefaultCurrency
            Me.SetTaxRate(False)

            'set llb faktur pajak
            If Me.chkTax.Checked = True Then
                Me.llbEditNoFakturPajak.Enabled = True
            Else
                Me.llbEditNoFakturPajak.Enabled = False
            End If

            '//TR : Tambahan baru Nik/Pas
            If Me.cboNIKPASS.Visible = True Then
                Me.cboNIKPASS.TabStop = True
                '2025 Core Tax Tahap 2.1
                Me.llbUpDateNPWP.Enabled = False
            Else
                Me.cboNIKPASS.TabStop = False
            End If
            Me.dgvItem.ReadOnly = False
            Me.BolehEditHarga = Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoice_OverridePrice")

            '2025 Core Tax Tahap 2.1
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
            'Me.TxtSPPBNo.ReadOnly = True
            Me.txtDPPNilaiLain.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage4.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)

            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit

            Me.dgvDP.AllowUserToDeleteRows = False
            Me.dgvItemSR.AllowUserToDeleteRows = False
            Me.dgvItem.AllowUserToDeleteRows = False
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
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtReadOnly.BackColor, False)

            Me.cboAccount.Enabled = False
            Me.cboCurrTax.Enabled = False

            Me.dgvItem.ReadOnly = True

            Me.txtNoFakturPajak.Enabled = False
            Me.txtRevisedInvoice.Enabled = False
            'Me.llbEditNoFakturPajak.Visible = False
            'Me.txtNoSeriFakturPajak.Visible = False
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
                Me.IsiCboDeliveryOrder2(Me.txtPrimaryCustomer.Text)
                Me.chkTax.Checked = .UseTax
                Me.TxtRemarks.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtNoFakturPajakSederhana.Text = .NoFakturPajakSederhana
                Me.TxtNamaSigned.Text = .NamaSigned
                Me.TxtJabatanSigned.Text = .JabatanSigned
                Me.TxtExchangeRate.Text = ObjCD.DigitGrouping(.ExchangeRate, DefPrecExchangeRate)
                Me.SetTxtExchangeRate()

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

                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotalB4DP, Me.DefaultPrecision))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(.TaxAmount, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax))
                Me.txtTaxAmount1.Text = Me.ObjCD.DigitGrouping(.TaxAmount1, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax)) 'new tax amount1
                Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(.DutyStamp, Me.ObjCurrency.GetDefaultPrecission(.PrimaryCurrTax))
                Me.txtTaxExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxExchangeRate, Me.DefaultPrecision))
                Me.txtDownPayment.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.DownPayment, Me.DefaultPrecision))
                Me.txtExhangeRateToIDR.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExchangRateToIDR, Me.DefaultPrecision))
                Me.txtExchangeRateToUSD.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.ExhangeRateToUSD, Me.DefaultPrecision))
                Me.TxtShipment.Text = Me.ObjCD.DigitGrouping(.OngkosKirim, Me.DefaultPrecision)
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(.TaxPercentage, 2)
                Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.OtherCharge, Me.DefaultPrecision))
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))

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
                'salin data sales return
                IsiDGVReturnItem()

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                '2025 Core Tax Tahap 2.2
                Me.SPPBDate = .SPPBDate
                Me.chkRevisedInvoice.Checked = .TaxRevised
                Me.AddInfo = .PrimaryAddInfo
                Me.FacilityStamp = .PrimaryFacilityStamp
                '2025 Core Tax Tahap 2.2, diupdate ketika semua type bisa dipilih, jadi selectedindex, tanpa percabangan
                'If .TaxType = ClsTransaksiSalesInvoice.EnumTaxType.Tax04 Then
                '    Me.cboTaxType.SelectedIndex = 0
                'Else
                '    Me.cboTaxType.SelectedIndex = 1
                'End If
                Me.cboTaxtype.SelectedIndex = .TaxType
                Me.txtDPPNilaiLain.Text = Me.ObjCD.DigitGrouping(.DPPNilaiLain, Me.DefaultPrecision)
                Me.txtRevisedInvoice.Text = IIf(.Revisi, "01", "00")
            End With
            Me.FormatDGV()
            Me.SetSalesPerson()

            Me.LblPrintStatus.Text = "Print Status: " + Me.ObjInvoice.GetPrintStatus(Me.txtPrimaryInvoice.Text)
            If Me.LblPrintStatus.Text = "Print Status: " Then
                Me.LblPrintStatus.Text = "Print Status: Unprinted"
            End If

            If Me.TBLHeaderCost.Rows.Count > 0 Then
                Me.rowHeaderCost = Me.TBLHeaderCost.Rows(0)
                Me.PrimaryCustomerShipTo = Me.rowHeaderCost.primary_cust_ship_to
                Me.IdentifierAddressCustomerShipTo = Me.rowHeaderCost.identifier_cust_ship_to
            End If

            

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
                .UrutanFakturPajak = 0
                .ExchangeRate = Me.TxtExchangeRate.Text
                .OngkosKirim = Me.TxtShipment.Text
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
                'jika datanya baru, maka update data alamat pajak 1 dan alamat pajak 2
                If Me.baru = True Then
                    Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, .AlamatPajak1, .AlamatPajak2)
                Else
                    'jika ada perubahan data customer dari invoice lama maka update data alamat customer
                    If Me.ObjInvoice.Get_Primary_Cust(Me.txtPrimaryInvoice.Text) <> Me.txtPrimaryCustomer.Text Then
                        Me.ObjCustomer.GetAlamatPajak(Me.txtPrimaryCustomer.Text, .AlamatPajak1, .AlamatPajak2)
                    End If
                End If

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

                '2025 Core Tax Tahap 2.2
                .SPPBDate = Me.SPPBDate
                .TaxRevised = Me.chkRevisedInvoice.Checked
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


    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiSalesInvoice
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice

            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                If Me.ObjInvoice.CariDetail(hj.PrimaryTransaksi, _
                Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me) _
                = True Then
                    Me.SalinKeField()
                    Me.NonAktifkanDetail()
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


    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================

    Public Enum EnumHitungInclude As Integer
        'UnitPriceSO
        'UnitPriceInclude
        NettPriceInclude
    End Enum

    Private Sub FrmTransaksiAOP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ColPrimarySalesPerson.Visible = False
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.CboFakturPajak.SelectedIndex = 0
            Me.cboCalculatesBy.SelectedIndex = 0
            Me.dgvDP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvExchageRates.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemSR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvSR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.chkShowCustCode.Checked = False

            Me.cboShipmentTax.SelectedIndex = 0
            Me.cboPPNDipungut.SelectedIndex = 0

            Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude
            'If Me.UserCode <> "USER" Then
            '    Me.LblPrintStatus.Visible = False
            'End If
            Me.TabControl1.Location = New Point(Me.TabControl1.Left, Me.btnOK.Top - 5 - Me.TabControl1.Height)
            Me.dtpDate.Left = Me.SplitContainer5.Panel1.Width - Me.dtpDate.Width
            Me.llbEditNoInvoice.Left = Me.dtpDate.Left - 5 - Me.llbEditNoInvoice.Width
            Me.txtInvoiceNo.Width = Me.llbEditNoInvoice.Left - 5 - Me.txtInvoiceNo.Left
            Me.TxtCustomerName.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.TxtCustomerName.Width
            Me.btnFindCustomer.Left = Me.TxtCustomerName.Left - 5 - Me.btnFindCustomer.Width
            Me.txtCustomerCode.Width = Me.btnFindCustomer.Left - 5 - Me.txtCustomerCode.Left
            Me.TxtSalesPersonName.Left = Me.TxtCustomerName.Left
            Me.TxtSalesPersonName.Width = Me.TxtCustomerName.Width
            Me.TxtSalesPersonCode.Width = Me.TxtSalesPersonName.Left - 5 - Me.TxtSalesPersonCode.Left

            Me.llbUpDateNPWP.Left = Me.dtpDate.Left + Me.dtpDate.Width - Me.llbUpDateNPWP.Width
            Me.txtNoNPWP.Width = Me.llbUpDateNPWP.Left - 5 - Me.txtNoNPWP.Left

            Me.llbAddress.Left = Me.llbUpDateNPWP.Left
            Me.cboCustomerAddress.Width = Me.txtNoNPWP.Width
            Me.LlbViewDetailTax.Left = Me.llbAddress.Left
            Me.llbEditNoFakturPajak.Left = Me.llbAddress.Left
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
            Me.txtTaxAmount1.Left = Me.txtSubTotalAfterDP.Width + Me.txtSubTotalAfterDP.Left - Me.txtTaxAmount.Width
            Me.cboPPNDipungut.Width = Me.txtTaxAmount.Left - 5 - Me.cboPPNDipungut.Left
            Me.txtOtherCharge.Width = Me.txtSubTotalAfterDP.Width
            Me.TextBox2.Width = Me.txtSubTotalAfterDP.Width
            Me.TextBox2.Top = Me.txtOtherCharge.Top
            Me.txtDutyStamp.Width = Me.txtSubTotalAfterDP.Width
            Me.TxtTotal.Width = Me.txtSubTotalAfterDP.Width
            Me.TxtRemarks.Width = Me.txtSubTotalAfterDP.Width

            Me.dtpDate.CustomFormat = "dd MMM yyyy"
            Me.dgvDP.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvDP.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            Me.dgvItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Me.dgvItem.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

            'for quick input
            If Me.QuickInput = True Then
                Me.CreateInvoiceByQuickInputStockOut(Me.TblQuickInput, Me.DirectlySave, Me.DirectlyPrint, Me.PathPrintCustom)
            Else
                Me.lblProgress.Visible = False
                Me.ProgressBar1.Visible = False
                Me.BtnStop.Visible = False
            End If

            '2025 Core Tax Tahap 2.1
            Me.ColVAT.Visible = False
            Me.ColUnitPriceInclude.Visible = False
            Me.ColIsVAT.Visible = False
            Me.ColVAT2Return.Visible = False
            Me.ColPriceByMultiUnit_Str.Visible = False
            Me.ColAmountIncludeReturn.Visible = False
            Me.ColAmountReturn.Visible = False
            Me.ColVAT2Return.Visible = False
            Me.ColAmount2Return.Visible = False
            Me.colCostAccCode.Visible = False
            Me.colCostAccDesc.Visible = False
            Me.ColDiscount1AmountInclude.Visible = False
            Me.ColDiscount1Include.Visible = False
            Me.ColDiscount1PercInclude.Visible = False
            Me.ColDiscAmountInclude.Visible = False
            Me.ColDiscountInclude.Visible = False
            Me.ColPrimarySalesPerson.Visible = False
            Me.ColDiscountGlobalAmountInclude.Visible = False
            Me.ColDiscountGlobalPercInclude.Visible = False
            Me.ColDiscountGlobalStringInclude.Visible = False
            Me.ColNettPriceInclude.Visible = False
            Me.ColUnitPriceInclude.Visible = False
            Me.ColAmountReturn.Visible = False
            Me.ColAmountInclude.Visible = False

            Me.BackgroundWorker2.RunWorkerAsync()

            '2025 Core Tax Tahap 2.2
            Me.llbEditNoFakturPajak.Visible = False
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
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
                If Me.CboFakturPajak.SelectedIndex = 0 Or Me.chkTax.Checked = False Then
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
        Try
            Me.cboDO.DataSource = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTaxSetting()
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


    Private Sub cboPPNDipungut_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPPNDipungut.SelectedValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub SetTxtExchangeRate()
        Try
            'If Me.txtPrimaryCurrency.Text <> Core_Module.PrimaryBaseCurrency Then
            'Me.TxtExchangeRate.Visible = True
            'Me.TxtExchangeRate.TabStop = True
            'Me.LblExchangeRate.Visible = True
            'Me.btnFindCurrency.Left = Me.LblExchangeRate.Left - 5 - Me.btnFindCurrency.Width
            'Me.txtCurrencyCode.Width = Me.btnFindCurrency.Left - 5 - Me.txtCurrencyCode.Left
            'Else
            Me.TxtExchangeRate.Visible = False
            Me.TxtExchangeRate.TabStop = False
            Me.LblExchangeRate.Visible = False
            'Me.btnFindCurrency.Left = Me.TxtExchangeRate.Left + Me.TxtExchangeRate.Width - Me.btnFindCurrency.Width
            'Me.txtCurrencyCode.Width = Me.btnFindCurrency.Left - 5 - Me.txtCurrencyCode.Left
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ' MsgBox("Jangan lupa betulin batasan tax invoice no dan batasan hapus.".ToUpper, MsgBoxStyle.Critical, "Info")
            Me.baru = True
            If Me.QuickInput = False Then
                Me.NonAktifkanDetail()
            End If
            'Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.ObjInvoice.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice
            Control.CheckForIllegalCrossThreadCalls = False
            Me.DefaultPrecision_DefCur = Me.ObjCurrency.GetDecimal_DefaultCurrency
            Me.Icon = My.Resources.desktop_icon_3_fw
            Me.ShowIcon = True
            Me.txtNoNIKPass.Width = Me.txtNoNPWP.Width
            'Me.ColPriceByMultiUnit.Visible = True

            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.SplitContainer3.TabStop = False
            Me.SplitContainer4.TabStop = False
            Me.SplitContainer5.TabStop = False

            '2025 Core Tax Tahap 2.2
            Me.llbEditNoFakturPajak.Visible = False
            Me.txtNoFakturPajakSederhana.Visible = False
            Me.CboFakturPajak.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()

                'for quick input
                If Me.QuickInput = True Then
                    Dim SimpanQuickInput As Boolean = True
                    If Me.DirectlySave = False Then
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil <> MsgBoxResult.Yes Then
                            SimpanQuickInput = False
                        End If
                    End If

                    If SimpanQuickInput = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        'antrian per transaksi
                        If Me.ObjInvoice.SimpanBaru(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me.UserCode, Me) Then
                            Me.NonAktifkanDetail()
                            Me.SalinKeField()
                            Me.btnFind.Focus()
                            Me.ObjCustomer.Update_EFaktur(Me.txtPrimaryCustomer.Text, Me.UserCode, Me.Txt_E_Faktur_Jalan.Text, _
                            Me.Txt_E_Faktur_Blok.Text, Me.Txt_E_Faktur_No.Text, Me.Txt_E_Faktur_RT.Text, Me.Txt_E_Faktur_RW.Text, _
                            Me.Txt_E_Faktur_Kelurahan.Text, Me.Txt_E_Faktur_Kecamatan.Text, Me.Txt_E_Faktur_Kabupaten.Text, Me.Txt_E_Faktur_Propinsi.Text, _
                            Me.Txt_E_Faktur_KodePos.Text, Me.Txt_E_Faktur_NoTelepon.Text)
                        End If

                        Me.Close()
                        Me.btnExit.PerformClick()
                    End If

                    Exit Sub
                End If

                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.SedangRevisi = True Then
                            If Me.ObjInvoice.SimpanRevisi(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                            Me.dgvExchageRates, Me.UserCode, Me.ObjInvoice.Primary, Me) Then
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
                            Me.dgvExchageRates, Me.UserCode, Me) Then
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
                        Me.dgvItemSR, Me.dgvExchageRates, Me.UserCode, Me) Then
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
            Else
                'for quick input
                If Me.QuickInput = True Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
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
                    'for quick input
                    If Me.QuickInput = True Then
                        Me.Close()
                        Exit Sub
                    End If

                    Me.KosongkanItem()
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        Me.ObjInvoice.CariDetail(Me.txtPrimaryInvoice.Text, Me.dgvItem, _
                        Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me)
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

    Private Sub cboShipmentTax_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboShipmentTax.SelectedValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.ObjInvoice.ApakahBolehEdit(Me.ObjInvoice.Primary, Me.UserCode, True) = True Then
                    If Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = True Then
                        MsgBox("This invoice is not sales invoice. Use quick invoice form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If Me.ObjInvoice.GetTimestamp(Me.ObjInvoice.Primary) <> Me.ObjInvoice.TimeStamp Then
                        Me.ObjInvoice.CariDetail(Me.ObjInvoice.Primary, Me.dgvItem, Me.dgvDP, _
                        Me.dgvItemSR, Me.dgvExchageRates, Me)
                    End If
                    Me.baru = False
                    Me.SedangRevisi = False
                    Me.ObjInvoice.CariDetail(Me.ObjInvoice.Primary, Me.dgvItem, Me.dgvDP, _
                    Me.dgvItemSR, Me.dgvExchageRates, Me)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim Rev As Integer = 0

        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.ObjInvoice.ApakahBolehHapus(Me.ObjInvoice.Primary, True, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    If Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = True Then
                        MsgBox("This invoice is not sales invoice. Use quick invoice form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Dim NoInvLama As String = ""
                    If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                        NoInvLama = Me.txtInvoiceNo.Text.Substring(0, Me.txtInvoiceNo.Text.Length - 1)
                    Else
                        NoInvLama = ""
                    End If
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor

                        If Me.ObjInvoice.Hapus(Me.txtPrimaryInvoice.Text, Me.UserCode, True, False) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            'jika masih ada invoce yang revisi selanjutnya maka tampilkan
                            'If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                            '    Dim PrimaryInvBaru As Double = Me.ObjInvoice.GetPrimaryByNoTransaksiLIKE(NoInvLama)
                            '    If PrimaryInvBaru <> -1 Then
                            '        If MsgBox("Load previous revision?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            '            Me.ObjInvoice.CariDetail(PrimaryInvBaru, Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                            '            Me.dgvExchageRates, Me)
                            '            Me.SalinKeField()
                            '        End If
                            '    Else
                            '        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            '    End If
                            'Else
                                MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                            'End If
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
        Try
            If Me.TxtRemarks.Enabled = True And (Me.txtPrimaryCustomer.Text <> "-1" Or Me.dgvItem.RowCount > 0) Then
                'for quick input
                If Me.QuickInput = True Then
                    Exit Sub
                End If

                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.SedangRevisi = True Then
                                If Me.ObjInvoice.SimpanRevisi(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, _
                                Me.dgvExchageRates, Me.UserCode, Me.ObjInvoice.Primary, Me) = True Then
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
                                Me.dgvExchageRates, Me.UserCode, Me) = True Then
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
                            Me.dgvExchageRates, Me.UserCode, Me) = True Then
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
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim DefaultPrecItem As Double = 0
        Dim AdaBatch, AdaED As Boolean
        Dim IsMultiUnit As Boolean = False
        Dim PriceByUnitDefault As Boolean = False

        Try
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If Me.dgvItem.RowCount > 0 Then
                row = Me.dgvItem.Rows(0)
                DefaultPrecItem = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
                'format dp
                For Each row In Me.dgvDP.Rows
                    row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
                Next
                'yang diformat dgv-nya:
                For Each row In Me.dgvItem.Rows
                    'digit grouping: qty, secondary qty,unit price, nett price, amount
                    If row.Cells(Me.ColLotNumber.Index).Value <> "" Then
                        AdaBatch = True
                        If AdaED = True Then
                            Exit For
                        End If
                    End If
                    If row.Cells(Me.ColExpiry.Index).Value <> "-" Then
                        AdaED = True
                        If AdaBatch = True Then
                            Exit For
                        End If
                    End If
                Next

                If AdaED = True Then
                    Me.ColExpiry.Visible = True
                Else
                    Me.ColExpiry.Visible = False
                End If
                If AdaBatch = True Then
                    Me.ColLotNumber.Visible = True
                Else
                    Me.ColLotNumber.Visible = False
                End If

                'yang diformat dgv-nya:
                Dim TidakAdaSecondary As Boolean = True
                For Each row In Me.dgvItem.Rows
                    '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                    '     secondary unitsnya "-"
                    If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = True Then
                        If row.Cells(Me.ColSecQty.Index).Value <> "-1" Then
                            TidakAdaSecondary = False
                        End If
                    End If
                    If row.Cells(Me.ColSecQty.Index).Value = "-1" Or row.Cells(Me.ColSecUnit.Index).Value = "0" Then
                        row.Cells(Me.ColSecQty.Index).Value = "-"
                        row.Cells(Me.ColSecUnit.Index).Value = "-"
                    End If
                Next
                '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
                '   tidak perlu ditampilkan
                If TidakAdaSecondary = True Then
                    Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = False
                    Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = False
                Else
                    Me.dgvItem.Columns(Me.ColSecQty.Index).Visible = True
                    Me.dgvItem.Columns(Me.ColSecUnit.Index).Visible = True
                End If
            End If

            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnit.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
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

            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    PriceByUnitDefault = True
                    Exit For
                End If
            Next
            Me.ColPriceByMultiUnit_Str.Visible = PriceByUnitDefault

            'untuk exchange rate
            For Each row In Me.dgvExchageRates.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
            Next

            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #10/9/2025# And Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = False Then
                Me.ColUnitPrice.Visible = False
                Me.ColNettPrice.Visible = False
                Me.ColAmount.Visible = False
                Me.ColDiscount.Visible = False
                Me.ColDiscountAmount.Visible = False

                Me.ColVAT2.Visible = True
                Me.ColAmount2.Visible = True
                Me.ColPriceCond.Visible = True

                Me.ColUnitPriceInclude.Visible = True
                Me.ColNettPriceInclude.Visible = True
                Me.ColAmountInclude.Visible = True
                Me.ColDiscAmountInclude.Visible = True
                Me.ColDiscountInclude.Visible = True
            ElseIf Me.dtpDate.Value >= #10/9/2025# And Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = True Then
                Me.ColUnitPrice.Visible = True
                Me.ColNettPrice.Visible = True
                Me.ColAmount.Visible = True
                Me.ColDiscount.Visible = True
                Me.ColDiscountAmount.Visible = True

                Me.ColVAT2.Visible = False
                Me.ColAmount2.Visible = False
                Me.ColPriceCond.Visible = True

                Me.ColUnitPriceInclude.Visible = False
                Me.ColNettPriceInclude.Visible = False
                Me.ColAmountInclude.Visible = False
                Me.ColDiscAmountInclude.Visible = False
                Me.ColDiscountInclude.Visible = False
            ElseIf Me.dtpDate.Value >= #1/1/2025# And Me.dtpDate.Value < #10/9/2025# Then
                Me.ColUnitPrice.Visible = True
                Me.ColNettPrice.Visible = True
                Me.ColAmount.Visible = True
                Me.ColDiscount.Visible = True
                Me.ColDiscountAmount.Visible = True

                Me.ColVAT2.Visible = False
                Me.ColAmount2.Visible = False
                Me.ColPriceCond.Visible = True

                Me.ColUnitPriceInclude.Visible = False
                Me.ColNettPriceInclude.Visible = False
                Me.ColAmountInclude.Visible = False
                Me.ColDiscAmountInclude.Visible = False
                Me.ColDiscountInclude.Visible = False
            Else
                Me.ColUnitPrice.Visible = True
                Me.ColNettPrice.Visible = True
                Me.ColAmount.Visible = True
                Me.ColDiscount.Visible = True
                Me.ColDiscountAmount.Visible = True

                Me.ColVAT2.Visible = False
                Me.ColAmount2.Visible = False
                Me.ColPriceCond.Visible = False

                Me.ColUnitPriceInclude.Visible = False
                Me.ColNettPriceInclude.Visible = False
                Me.ColAmountInclude.Visible = False
                Me.ColDiscAmountInclude.Visible = False
                Me.ColDiscountInclude.Visible = False
            End If

            Me.ColVAT.Visible = False
            Me.ColIsVAT.Visible = False
            Me.ColVAT2Return.Visible = False
            Me.ColPriceByMultiUnit_Str.Visible = False
            Me.ColAmountIncludeReturn.Visible = False
            Me.ColAmountReturn.Visible = False
            Me.ColVAT2Return.Visible = False
            Me.ColAmount2Return.Visible = False
            Me.colCostAccCode.Visible = False
            Me.colCostAccDesc.Visible = False
            Me.ColDiscount1AmountInclude.Visible = False
            Me.ColPrimarySalesPerson.Visible = False
            Me.ColDiscountGlobalAmountInclude.Visible = False
            Me.ColDiscountGlobalPercInclude.Visible = False
            Me.ColDiscountGlobalStringInclude.Visible = False
            Me.ColAmountReturn.Visible = False
            Me.ColDiscount1AmountInclude.Visible = False
            Me.ColDiscount1Include.Visible = False
            Me.ColDiscount1PercInclude.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try

            If e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColAmount.Index Or _
            e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColDiscountAmount.Index Then
                If Me.BolehEditHarga = True Then
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColRemarks.Index Or e.ColumnIndex = Me.ColRemarksItem.Index Then
                Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColRemarks.Index Then
                    If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Length > 150 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Substring(0, 150)
                    End If
                ElseIf e.ColumnIndex = Me.ColUnitPrice.Index Then
                    Me.EndEditUnitPrice(sender, e, Me.dgvItem.Rows(e.RowIndex))
                ElseIf e.ColumnIndex = Me.ColDiscount.Index Then
                    Me.EndEditDiscount(sender, e, Me.dgvItem.Rows(e.RowIndex))
                ElseIf e.ColumnIndex = Me.ColDiscountAmount.Index Then
                    Me.EndEditDiskonAmount(sender, e, Me.dgvItem.Rows(e.RowIndex))
                ElseIf e.ColumnIndex = Me.ColAmount.Index Then
                    Me.EndEditAmount(sender, e, Me.dgvItem.Rows(e.RowIndex))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditAmount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim Dummy As String = ""
        Dim BufferNettPrice, BufferUnitPrice As String

        Try
            With row
                'untuk nilai unit price, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for unit price.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefaultPrecision)

                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    BufferNettPrice = Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value
                    BufferUnitPrice = Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value

                    'hitung nett price
                    Dummy = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColAmount.Index).Value) / _
                    CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecision))
                    If CDec(Dummy) <= 0 Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    Else
                        row.Cells(Me.ColNettPrice.Index).Value = Dummy
                    End If

                    'hitung disc/disc amount
                    If Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "0" Then
                        'disc dari amount
                        'unit price
                        Dummy = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColNettPrice.Index).Value) + _
                        CDec(row.Cells(Me.ColDiscountAmount.Index).Value), DefaultPrecision))
                        If CDec(Dummy) = 0 Then
                            Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value = BufferNettPrice
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        Else
                            row.Cells(Me.ColUnitPrice.Index).Value = Dummy
                        End If
                        'disc %
                        row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                        ((CDec(row.Cells(Me.ColDiscountAmount.Index).Value) / CDec(row.Cells(Me.ColUnitPrice.Index).Value))) * 100, 2))
                    Else
                        'disc dari %
                        'unit price
                        Dummy = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal((CDec(row.Cells(Me.ColNettPrice.Index).Value) * 100) / _
                        (100 - CDec(row.Cells(Me.ColDiscount.Index).Value)), DefaultPrecision))
                        If CDec(Dummy) = 0 Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem.Item(Me.ColNettPrice.Index, e.RowIndex).Value = BufferNettPrice
                            Exit Sub
                        Else
                            row.Cells(Me.ColUnitPrice.Index).Value = Dummy
                        End If
                        'disc amount
                        row.Cells(Me.ColDiscountAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                        ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecision))
                    End If 
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiskonAmount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai disc amount, pastikan nilainya positif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for discount amount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefaultPrecision)

                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'jika unit price < nett price maka kembalikan ke nilai buffer
                    'hanya berlaku klo menentukan diskon berdasarkan amount
                    If CDec(Me.dgvItem.Item(Me.ColDiscountAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'tandai diskon dihitung berdasarkan amount
                    row.Cells(Me.ColDiscFromPerct.Index).Value = "0"
                    'hitung ulang row
                    Me.Hitung_Amount(row)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiscount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
   ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai diskon, pastikan nilainya tidak negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for discount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek nili maksimum untuk diskon adalah 100%
                If CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) >= 100 Then
                    MsgBox("Valid value is <= 100%.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, 2)
                'hanya hitung ulang kalo ada perubahan data
                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'tandai di ubah dari diskon
                    row.Cells(Me.ColDiscFromPerct.Index).Value = "1"
                    'hitung ulang row
                    Me.Hitung_Amount(row)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditUnitPrice(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai unit price, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for unit price.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefaultPrecision)

                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'jika unit price < nett price maka kembalikan ke nilai buffer
                    'hanya berlaku klo menentukan diskon berdasarkan amount
                    If Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "0" Then
                        If CDec(Me.dgvItem.Item(Me.ColDiscountAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    'hitung ulang row
                    Me.Hitung_Amount(row)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Amount(ByVal row As DataGridViewRow)
        Dim DiscPerct As Integer = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= ObjSetNPWP.Tglberlaku Then
                Me.Hitung_Amount_MiddlePoint(row)
                Exit Sub
            End If
            'hitung nett price
            DiscPerct = row.Cells(Me.ColDiscFromPerct.Index).Value
            If DiscPerct = 1 Then
                row.Cells(Me.ColDiscountAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecision))
            Else
                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscountAmount.Index).Value) / CDec(row.Cells(Me.ColUnitPrice.Index).Value))) * 100, 2))
            End If
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(row.Cells(Me.ColUnitPrice.Index).Value) - CDec(row.Cells(Me.ColDiscountAmount.Index).Value) _
            , DefaultPrecision))

            'hitung amount
            row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
            row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub Hitung_Amount_MiddlePoint(ByVal row As DataGridViewRow)
        Dim DiscPerct As Integer = 0
        '\\TR==GlobalDiscount
        Dim netPriceSementara As Decimal = 0
        Try
            'hitung nett price

            DiscPerct = row.Cells(Me.ColDiscount1PercInclude.Index).Value
            If DiscPerct = 1 Then
                row.Cells(Me.ColDiscount1AmountInclude.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) * CDec(row.Cells(Me.ColDiscount1Include.Index).Value))) / 100, DefaultPrecision))
            Else
                row.Cells(Me.ColDiscount1Include.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscount1AmountInclude.Index).Value) / CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value))) * 100, 2))
            End If

            netPriceSementara = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                       CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) - CDec(row.Cells(Me.ColDiscount1AmountInclude.Index).Value), DefaultPrecision))

            'If Me.chkUseGlobalDisc.Checked = False Then
            row.Cells(Me.ColDiscAmountInclude.Index).Value = row.Cells(Me.ColDiscount1AmountInclude.Index).Value
            row.Cells(Me.ColDiscountInclude.Index).Value = row.Cells(Me.ColDiscount1Include.Index).Value
            row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = row.Cells(Me.ColDiscount1AmountInclude.Index).Value
            'End If

            'If Me.txtGlobalIsDiscMulti.Text = "0" Then
            '    row.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(netPriceSementara * CDec(Me.txtGlobalDisc.Text) / 100, DefaultPrecision)
            '    If CDec(row.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value) <> 0 Then
            '        row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = row.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value
            '    End If
            'Else
            '    If row.Cells(Me.ColPrimarySORef.Index).Value = "-1" Then
            '        If Me.GetAmountDiscBertingkat(row.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value, row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value, netPriceSementara, DefaultPrecision, True) = False Then
            '            Me.dgvItem.Rows(row.Index).Selected = True
            '            If dariEndEditDisc1 = True Then
            '                row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = Me.txtGlobalDiscString.Text
            '                If DiscPerct = 1 Then
            '                    row.Cells(Me.ColDiscount1PercInclude.Index).Value = Me.buffer
            '                Else
            '                    row.Cells(Me.ColDiscount1AmountInclude.Index).Value = Me.buffer
            '                End If
            '            Else
            '                Me.txtGlobalDiscString.Text = row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value
            '            End If
            '            dariEndEditDisc1 = False
            '            dariEndEditUnitPrice = False
            '            Exit Sub
            '        Else
            '            If netPriceSementara <> 0 Then
            '                row.Cells(Me.ColDiscountGlobalPercInclude.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value) / netPriceSementara) * 100, 2)
            '            End If
            '        End If
            '    Else
            '        row.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(0, DefaultPrecision)
            '        row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = "0.00"
            '        row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = Me.txtGlobalDiscString.Text
            '    End If
            'End If

            row.Cells(Me.ColNettPriceInclude.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
           CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) - CDec(row.Cells(Me.ColDiscount1AmountInclude.Index).Value) - CDec(row.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value) _
           , DefaultPrecision))

            'karena penambahan fitur global discount, maka discount final discount harus di itung ulang
            'If Me.chkUseGlobalDisc.Checked = True Then
            '    row.Cells(Me.ColDiscAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) - CDec(row.Cells(Me.ColNettPriceInclude.Index).Value), DefaultPrecision)
            '    row.Cells(Me.ColDiscountInclude.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColDiscAmountInclude.Index).Value) / CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value)) * 100, 2)
            'End If

            'hitung amount
            If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                row.Cells(Me.ColAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQtyMulti.Index).Value * _
                    row.Cells(Me.ColNettPriceInclude.Index).Value, DefaultPrecision))
            Else
                row.Cells(Me.ColAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQty.Index).Value * _
                row.Cells(Me.ColNettPriceInclude.Index).Value, DefaultPrecision))
            End If
            'dariEndEditDisc1 = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            'hapus semua item yang sama primary DO-nya
            Dim ListPrimarySO As New List(Of Double)
            Dim row As DataGridViewRow
            Dim i As Integer = 0
            While i <= Me.dgvItem.Rows.Count - 1
                row = Me.dgvItem.Rows(i)
                If row.Cells(Me.ColPrimaryDO.Index).Value = Me.buffer Then
                    If ListPrimarySO.Contains(row.Cells(Me.ColPrimarySO.Index).Value) = False Then
                        ListPrimarySO.Add(row.Cells(Me.ColPrimarySO.Index).Value)
                    End If
                    Me.dgvItem.Rows.Remove(row)
                Else
                    i += 1
                End If
            End While

            'tambahkan data DO ke dalam tabel sumber
            Me.TBL.Rows.Add(Me.buffer, Me.objDO.GetDONO(Me.buffer))
            'hapus data sales return yang merujuk ke DO ybs
            Me.IsiDGVReturn(False)
            'hapus data Down payment yang merujuk pada SO yg bersangkutan
            If Me.dgvDP.RowCount > 0 Then
                Me.HapusDP()
            End If
            'set number
            Me.ObjCD.SetNumber(Me.dgvItem)
            'hitung ulang nilai total
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If

            Me.SetSalesPerson()
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

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtCustomerCode.ReadOnly = False And Me.txtCustomerCode.Enabled = True Then
                    txtCustomerCode.BackColor = BackColorAll
                Else
                    txtCustomerCode.BackColor = Me.TxtTotal.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
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

    Private Function GetJmlDGVSR(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryDO As Double, _
     ByVal vSecQty As Double, ByVal vPrimarySO As Double, ByVal vBatchNumber As String, ByVal vED As Date) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvSR.Rows
                If vED = #1/1/1950# Then
                    If IsDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = False Then
                        If IsNumeric(row.Cells(Me.ColSecQty_SR.Index).Value) = True AndAlso CDec(row.Cells(Me.ColSecQty_SR.Index).Value) > 0 Then
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo _
                            And row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColSecQty_SR.Index).Value = vSecQty And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber Then
                                Jml = Jml + CDec(row.Cells(Me.ColTotalQtymulti_SR.Index).Value)
                            End If
                        Else
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo And _
                            row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber Then
                                Jml = Jml + CDec(row.Cells(Me.ColTotalQtymulti_SR.Index).Value)
                            End If
                        End If
                    End If
                Else
                    If IsDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = True Then
                        If IsNumeric(row.Cells(Me.ColSecQty_SR.Index).Value) = True AndAlso CDec(row.Cells(Me.ColSecQty_SR.Index).Value) > 0 Then
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo _
                            And row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColSecQty_SR.Index).Value = vSecQty And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber And _
                            CDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = vED Then
                                Jml = Jml + CDec(row.Cells(Me.ColTotalQtymulti_SR.Index).Value)
                            End If
                        Else
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo And _
                            row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber And _
                            CDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = vED Then
                                Jml = Jml + CDec(row.Cells(Me.ColTotalQtymulti_SR.Index).Value)
                            End If
                        End If
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGVSR_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryDO As Double, _
   ByVal vSecQty As Double, ByVal vPrimarySO As Double, ByVal vBatchNumber As String, ByVal vED As Date) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0

            For Each row In Me.dgvSR.Rows
                If vED = #1/1/1950# Then
                    If IsDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = False Then
                        If IsNumeric(row.Cells(Me.ColSecQty_SR.Index).Value) = True AndAlso CDec(row.Cells(Me.ColSecQty_SR.Index).Value) > 0 Then
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo _
                            And row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColSecQty_SR.Index).Value = vSecQty And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber Then
                                Jml = Jml + CDec(row.Cells(Me.COlTotalQtySR.Index).Value)
                            End If
                        Else
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo And _
                            row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber Then
                                Jml = Jml + CDec(row.Cells(Me.COlTotalQtySR.Index).Value)
                            End If
                        End If
                    End If
                Else
                    If IsDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = True Then
                        If IsNumeric(row.Cells(Me.ColSecQty_SR.Index).Value) = True AndAlso CDec(row.Cells(Me.ColSecQty_SR.Index).Value) > 0 Then
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo _
                            And row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColSecQty_SR.Index).Value = vSecQty And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber And _
                            CDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = vED Then
                                Jml = Jml + CDec(row.Cells(Me.COlTotalQtySR.Index).Value)
                            End If
                        Else
                            If row.Cells(Me.COlPrimaryCodeNoSR.Index).Value = vPrimaryCodeNo And _
                            row.Cells(Me.ColPrimaryDO_SR.Index).Value = vPrimaryDO And _
                            row.Cells(Me.ColPrimarySO_SR.Index).Value = vPrimarySO And _
                            row.Cells(Me.ColBatchNo_SR.Index).Value = vBatchNumber And _
                            CDate(row.Cells(Me.ColExpiryDate_SR.Index).Value) = vED Then
                                Jml = Jml + CDec(row.Cells(Me.COlTotalQtySR.Index).Value)
                            End If
                        End If
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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

    'Private Sub Hitung_Amount2(ByVal row As DataGridViewRow)
    'Dim DefaultPrecisionRow As Integer

    'Try
    'DefaultPrecisionRow = Me.ObjCurrency.GetDefaultPrecission(row.Cells(Me.ColPrimaryCurrency.Index).Value)
    'hitung nett price
    'row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value * _
    '(100 - row.Cells(Me.ColDiscount.Index).Value) / 100, DefaultPrecisionRow))

    'hitung amount
    'If Me.MetodeHitungInclude = EnumHitungInclude.UnitPriceInclude Then
    'If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
    'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(row.Cells(Me.ColTotalQty.Index).Value * _
    '(row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), _
    'DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
    'DefaultPrecisionRow))
    'Else
    'If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
    'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDec(row.Cells(Me.ColTotalQty.Index).Value) * _
    '(row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100) _
    '* CDec(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow, MidpointRounding.AwayFromZero), DefaultPrecisionRow))
    'Else
    'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Math.Round(CDec(row.Cells(Me.ColTotalQty.Index).Value) * _
    '(row.Cells(Me.ColUnitPrice.Index).Value * ((100 - row.Cells(Me.ColDiscount.Index).Value)) / 100), DefaultPrecisionRow, MidpointRounding.AwayFromZero), _
    'DefaultPrecisionRow))
    'End If
    'End If
    'ElseIf Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
    'If IsNumeric(row.Cells(Me.ColSecQty.Index).Value) = False Then
    'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColTotalQty.Index).Value * _
    'row.Cells(Me.ColNettPrice.Index).Value, _
    'DefaultPrecisionRow))
    'Else
    'If row.Cells(Me.ColSecQty.Index).Value <> -1 Then
    'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColTotalQty.Index).Value) * _
    'row.Cells(Me.ColNettPrice.Index).Value _
    '* CDec(row.Cells(Me.ColSecQty.Index).Value), DefaultPrecisionRow))
    'Else
    'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDec(row.Cells(Me.ColTotalQty.Index).Value) * _
    'row.Cells(Me.ColNettPrice.Index).Value, _
    'DefaultPrecisionRow))
    'End If
    'End If
    'End If
    'Catch ex As Exception
    'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    'End Try
    'End Sub

    Private Sub IsiItemDO2(ByVal TBL1 As DseTransaksiDO.tbltrans_do_view3_newDataTable, ByVal vPrimaryDO As Double)
        Dim rowDO As DseTransaksiDO.DORow
        Dim rowDO1 As DseTransaksiDO.tbltrans_do_view3_newRow
        Dim Filter As String = ""
        Dim rowSO As DseTransaksiDO.SORow

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.IsiItemDO2_Update(TBL1, vPrimaryDO)
                Exit Sub
            End If

            'kosongkan tabel
            TBL1.Rows.Clear()
            Me.TBL_DoItem1.Rows.Clear()
            Me.TBL_SoItem1.Rows.Clear()
            Me.ListPrimarySODummy.Clear()
            Me.DTA_SoItem1.ClearBeforeFill = False

            'isikan tabel dummy
            Me.DTA_DoItem1.FillBy_PrimaryDO(Me.TBL_DoItem1, vPrimaryDO)
            For Each rowDO In Me.TBL_DoItem1.Rows
                If Me.ListPrimarySODummy.Contains(rowDO.Primary_SO) = False Then
                    Me.ListPrimarySODummy.Add(rowDO.Primary_SO)
                    Me.DTA_SoItem1.FillBy_PrimarySO(Me.TBL_SoItem1, rowDO.Primary_SO)
                End If
            Next

            'isikan ke dalam tabel asli
            For Each rowDO In Me.TBL_DoItem1.Rows
                rowDO1 = TBL1.Newtbltrans_do_view3_newRow

                With rowDO1
                    .Batch_No = rowDO.Batch_No
                    .DeliveryOrderNo = rowDO.DeliveryOrderNo
                    .Description = rowDO.Description
                    .Expiry_Date = rowDO.Expiry_Date
                    .Master_Unit = rowDO.Master_Unit
                    .Primary_DO = rowDO.Primary_DO
                    .Primary_Master_Units = rowDO.Primary_Master_Units
                    .Primary_Product = rowDO.Primary_Product
                    .Primary_Secondary_Units = rowDO.Primary_Secondary_Units
                    .Primary_SO = rowDO.Primary_SO
                    .Primary_Units = rowDO.Primary_Units
                    .Product_Code = rowDO.Product_Code
                    .Qty = rowDO.Qty
                    .Qty_Master = rowDO.Qty_Master
                    .Qty_Per_Master = rowDO.Qty_Per_Master
                    .Secondary_Qty = rowDO.Secondary_Qty
                    .Secondary_Unit = rowDO.Secondary_Unit
                    .Stock_Out_Number = rowDO.Stock_Out_Number
                    .Total_Qty = rowDO.Total_Qty
                    .Unit_Code = rowDO.Unit_Code
                    ._Customer_s_Code_No = ""
                    ._Customer_s_Desc = ""
                    ._Decimal = rowDO._Decimal
                    .Total_Qty_Multi = rowDO.Total_Qty_Multi
                    .Units_Code_Multi = rowDO.Units_Code_Multi
                    .Primary_Units_Multi = rowDO.Primary_Units_Multi
                    .Ratio_Reference = rowDO.Ratio_Reference
                    .Ratio_Default = rowDO.Ratio_Default

                    Filter = "[Primary SO] =" + .Primary_SO.ToString + " AND [Primary Product] =" + .Primary_Product.ToString
                    Me.TBL_SoItem1.DefaultView.RowFilter = Filter
                    If Me.TBL_SoItem1.DefaultView.Count > 0 Then
                        rowSO = Me.TBL_SoItem1.DefaultView(0).Row
                        .Primary_Customer = rowSO.Primary_Customer
                        .Nett_Price = rowSO.Nett_Price
                        .Sales_Order_No = rowSO.Sales_Order_No
                        .Price_Exclude = rowSO.Price_Exclude
                        .Primary_Currency = rowSO.Primary_Currency
                        .Currency_Code = rowSO.Currency_Code
                        .Cust_PO_No = rowSO.Cust_PO_No
                        .Tax = rowSO.Tax
                        .Default_Precision = rowSO.Default_Precision
                        .Discount = rowSO.Discount
                        .Unit_Price = rowSO.Unit_Price
                        .Unit_Price_Exclude = 0
                        .Disc_Amount = rowSO.Disc_Amount
                        .Disc_From_Perct = rowSO.Disc_From_Perct
                        .price_by_multi_unit = rowSO.price_by_multi_unit
                    Else
                        MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End With

                TBL1.Rows.Add(rowDO1)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '2025 Core Tax Tahap 2.1
    Private Sub IsiItemDO2_Update(ByVal TBL1 As DseTransaksiDO.tbltrans_do_view3_newDataTable, ByVal vPrimaryDO As Double)
        Dim rowDO As DseTransaksiDO.DORow
        Dim rowDO1 As DseTransaksiDO.tbltrans_do_view3_newRow
        Dim Filter As String = ""
        Dim rowSO As DseTransaksiDO.SORow

        Try
            'kosongkan tabel
            TBL1.Rows.Clear()
            Me.TBL_DoItem1.Rows.Clear()
            Me.TBL_SoItem1.Rows.Clear()
            Me.ListPrimarySODummy.Clear()
            Me.DTA_SoItem1.ClearBeforeFill = False

            'isikan tabel dummy
            Me.DTA_DoItem1.FillBy_PrimaryDO(Me.TBL_DoItem1, vPrimaryDO)
            For Each rowDO In Me.TBL_DoItem1.Rows
                If Me.ListPrimarySODummy.Contains(rowDO.Primary_SO) = False Then
                    Me.ListPrimarySODummy.Add(rowDO.Primary_SO)
                    Me.DTA_SoItem1.FillBy_PrimarySO(Me.TBL_SoItem1, rowDO.Primary_SO)
                End If
            Next

            'isikan ke dalam tabel asli
            For Each rowDO In Me.TBL_DoItem1.Rows
                rowDO1 = TBL1.Newtbltrans_do_view3_newRow

                With rowDO1
                    .Batch_No = rowDO.Batch_No
                    .DeliveryOrderNo = rowDO.DeliveryOrderNo
                    .Description = rowDO.Description
                    .Expiry_Date = rowDO.Expiry_Date
                    .Master_Unit = rowDO.Master_Unit
                    .Primary_DO = rowDO.Primary_DO
                    .Primary_Master_Units = rowDO.Primary_Master_Units
                    .Primary_Product = rowDO.Primary_Product
                    .Primary_Secondary_Units = rowDO.Primary_Secondary_Units
                    .Primary_SO = rowDO.Primary_SO
                    .Primary_Units = rowDO.Primary_Units
                    .Product_Code = rowDO.Product_Code
                    .Qty = rowDO.Qty
                    .Qty_Master = rowDO.Qty_Master
                    .Qty_Per_Master = rowDO.Qty_Per_Master
                    .Secondary_Qty = rowDO.Secondary_Qty
                    .Secondary_Unit = rowDO.Secondary_Unit
                    .Stock_Out_Number = rowDO.Stock_Out_Number
                    .Total_Qty = rowDO.Total_Qty
                    .Unit_Code = rowDO.Unit_Code
                    ._Customer_s_Code_No = ""
                    ._Customer_s_Desc = ""
                    ._Decimal = rowDO._Decimal
                    .Total_Qty_Multi = rowDO.Total_Qty_Multi
                    .Units_Code_Multi = rowDO.Units_Code_Multi
                    .Primary_Units_Multi = rowDO.Primary_Units_Multi
                    .Ratio_Reference = rowDO.Ratio_Reference
                    .Ratio_Default = rowDO.Ratio_Default

                    Filter = "[Primary SO] =" + .Primary_SO.ToString + " AND [Primary Product] =" + .Primary_Product.ToString
                    Me.TBL_SoItem1.DefaultView.RowFilter = Filter
                    If Me.TBL_SoItem1.DefaultView.Count > 0 Then
                        rowSO = Me.TBL_SoItem1.DefaultView(0).Row
                        .Primary_Customer = rowSO.Primary_Customer
                        .Nett_Price = rowSO.Nett_Price
                        .Sales_Order_No = rowSO.Sales_Order_No
                        .Price_Exclude = rowSO.Price_Exclude
                        .Primary_Currency = rowSO.Primary_Currency
                        .Currency_Code = rowSO.Currency_Code
                        .Cust_PO_No = rowSO.Cust_PO_No
                        .Tax = rowSO.Tax
                        .Default_Precision = rowSO.Default_Precision
                        .Discount = rowSO.Discount
                        .Unit_Price = rowSO.Unit_Price
                        .Unit_Price_Exclude = 0
                        .Disc_Amount = rowSO.Disc_Amount
                        .Disc_From_Perct = rowSO.Disc_From_Perct

                        '\\TR==GlobalDiscount
                        .Unit_Price_Exclude = rowSO.Unit_Price_Exclude
                        .Disc_From_Perct_Exclude = rowSO.Disc_From_Perct
                        .Disc_Perc_Exclude = rowSO.Discount_Amount_Exclude_Perc
                        .Disc_Amount_Exclude = rowSO.Discount_Amount_Exclude
                        .Nett_Price_Exclude = rowSO.Nett_Price_Exclude
                        If rowSO.Used_VAT = False Then
                            .Use_VAT = 0
                        Else
                            .Use_VAT = 1
                        End If
                        If rowSO.Use_Tax = False Then
                            .Use_Tax = 0
                        Else
                            .Use_Tax = 1
                        End If
                        .Primary_SO_Ref = rowSO.Primary_SO_Ref


                        .Disc_Amount = rowSO.Disc_Amount
                        .Disc_From_Perct = rowSO.Disc_From_Perct
                        .price_by_multi_unit = rowSO.price_by_multi_unit

                        'aul
                        .primary_cost_acc = rowSO.Primary_Cost_Acc
                        .cost_acc_code = rowSO.cost_acc_code
                        .cost_acc_desc = rowSO.cost_acc_desc
                    Else
                        MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End With

                TBL1.Rows.Add(rowDO1)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiItemDO1(ByVal vTBL As DseTransaksiDO.DataTable10DataTable, _
  ByVal vPrimaryDO As Double)
        Dim rowDO As DseTransaksiDO.tbltrans_do_item8Row
        Dim rowDO1 As DseTransaksiDO.DataTable10Row
        Dim Filter As String = ""
        Dim rowSO As DseTransaksiDO.tbltrans_so_item1Row

        Try
            'kosongkan data table 10
            vTBL.Rows.Clear()
            Me.ListPrimarySODummy.Clear()

            'isikan data table 8
            Me.DTADO2.Fill(Me.TBLDO2, vPrimaryDO)

            'isikan data so
            Me.TBLDO3.Rows.Clear()
            Me.DTADO3.ClearBeforeFill = False
            For Each rowDO In Me.TBLDO2.Rows
                If Me.ListPrimarySODummy.Contains(rowDO.Primary_SO) = False Then
                    Me.DTADO3.Fill(Me.TBLDO3, rowDO.Primary_SO)
                End If
            Next

            'pindahkan ke data table 10
            For Each rowDO In Me.TBLDO2.Rows
                rowDO1 = vTBL.NewDataTable10Row
                With rowDO1
                    .Primary_DO = rowDO.Primary_DO
                    .Primary_Product = rowDO.Primary_Product
                    .Primary_SO = rowDO.Primary_SO

                    Filter = "[Primary SO] =" + .Primary_SO.ToString + " AND [Primary Product] =" + .Primary_Product.ToString

                    Me.TBLDO3.DefaultView.RowFilter = Filter
                    If Me.TBLDO3.DefaultView.Count > 0 Then
                        rowSO = Me.TBLDO3.DefaultView(0).Row
                        .Price_Exclude = rowSO.Price_Exclude
                    Else
                        MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End With
                vTBL.Rows.Add(rowDO1)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
                    Me.SetExchangeRate()
                    If Me.TxtExchangeRate.Visible = True Then
                        Me.TxtExchangeRate.Focus()
                    Else
                        Me.cboCalculatesBy.Focus()
                    End If
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision()
                        Me.SetExchangeRate()
                        If Me.TxtExchangeRate.Visible = True Then
                            Me.TxtExchangeRate.Focus()
                        Else
                            Me.cboCalculatesBy.Focus()
                        End If
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision()
                    End If
                End If
                '2025 Core Tax Tahap 2.1
                If Me.txtPrimaryCurrency.Text <> "" Or Me.txtPrimaryCurrency.Text <> "-1" Then
                    '\\TR==GlobalDiscount
                    Dim row As DataGridViewRow
                    For Each row In Me.dgvItem.Rows
                        'If Me.txtGlobalIsDiscMulti.Text = "1" Then
                        '    row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = Me.txtGlobalDiscString.Text
                        'Else
                        '    row.Cells(Me.ColDiscountGlobalPercInclude.Index).Value = Me.txtGlobalDisc.Text
                        'End If
                        'Me.GenerateStringDiscBertingkat(row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value, DefaultPrecision)
                        Me.Hitung_Amount(row)
                    Next
                    Me.Hitung_Total()
                    'If Me.txtGlobalIsDiscMulti.Text = "1" Then
                    '    Me.GenerateStringDiscBertingkat(Me.txtGlobalDiscString.Text, DefaultPrecision)
                    'End If
                End If


                Me.SetTaxRate()
                Me.SetCalculateBy()
                Me.SetTxtExchangeRate()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
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
                        Me.SetExchangeRate()
                        If Me.TxtExchangeRate.Visible = True Then
                            Me.TxtExchangeRate.Focus()
                        Else
                            Me.cboCalculatesBy.Focus()
                        End If
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
            '2025 Core Tax Tahap 2.1
            If Me.txtPrimaryCurrency.Text <> "" Or Me.txtPrimaryCurrency.Text <> "-1" Then
                '\\TR==GlobalDiscount
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    'If Me.txtGlobalIsDiscMulti.Text = "1" Then
                    '    row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = Me.txtGlobalDiscString.Text
                    'Else
                    '    row.Cells(Me.ColDiscountGlobalPercInclude.Index).Value = Me.txtGlobalDisc.Text
                    'End If
                    'Me.GenerateStringDiscBertingkat(row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value, DefaultPrecision)
                    Me.Hitung_Amount(row)
                Next
                Me.Hitung_Total()
                'If Me.txtGlobalIsDiscMulti.Text = "1" Then
                '    Me.GenerateStringDiscBertingkat(Me.txtGlobalDiscString.Text, DefaultPrecision)
                'End If
            End If

            Me.SetTaxRate()
            Me.SetCalculateBy()
            Me.SetTxtExchangeRate()
            Me.SetExchangeRate()
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
                Me.SetExchangeRate()
                If Me.TxtExchangeRate.Visible = True Then
                    Me.TxtExchangeRate.Focus()
                Else
                    Me.cboCalculatesBy.Focus()
                End If
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
                    'Me.SetExchangeRate()
                    If Me.TxtExchangeRate.Visible = True Then
                        Me.TxtExchangeRate.Focus()
                    Else
                        Me.cboCalculatesBy.Focus()
                    End If
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetDefaultPrecision()
                End If
            End If
            '2025 Core Tax Tahap 2.1
            If Me.txtPrimaryCurrency.Text <> "" Or Me.txtPrimaryCurrency.Text <> "-1" Then
                '\\TR==GlobalDiscount
                Dim row As DataGridViewRow
                For Each row In Me.dgvItem.Rows
                    'If Me.txtGlobalIsDiscMulti.Text = "1" Then
                    '    row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = Me.txtGlobalDiscString.Text
                    'Else
                    '    row.Cells(Me.ColDiscountGlobalPercInclude.Index).Value = Me.txtGlobalDisc.Text
                    'End If
                    'Me.GenerateStringDiscBertingkat(row.Cells(Me.ColDiscountGlobalStringInclude.Index).Value, DefaultPrecision)
                    Me.Hitung_Amount(row)
                Next
                Me.Hitung_Total()
                'If Me.txtGlobalIsDiscMulti.Text = "1" Then
                '    Me.GenerateStringDiscBertingkat(Me.txtGlobalDiscString.Text, DefaultPrecision)
                'End If
            End If

            Me.SetTaxRate()
            Me.SetCalculateBy()
            Me.SetTxtExchangeRate()
            Me.SetExchangeRate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetExchangeRate()
        Try
            If Me.txtPrimaryCurrency.Text <> Core_Module.PrimaryBaseCurrency Then
                If CDec(Me.TxtExchangeRate.Text) = 1 Then
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.ObjCurrency.GetDefaultExhangeRate(Me.txtPrimaryCurrency.Text), _
                    Core_Module.DefPrecExchangeRate)
                End If
            Else
                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(1, _
                Core_Module.DefPrecExchangeRate)
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
                Me.txtSubTotalBeforeDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotalBeforeDP.Text, DefaultPrecision))
                Me.txtSubTotalAfterDP.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotalAfterDP.Text, DefaultPrecision))
                '  Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxAmount.Text, 0))
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim UseTaxSetting, TaxPercentSetting As String
        Dim deffPrecTax As Integer = 0
        deffPrecTax = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
        Try
            If Me.chkTax.Checked = True Then
                Me.txtTax.ReadOnly = False
                Me.txtTax.TabStop = True
                Me.txtTax.BackColor = Me.txtPrimaryInvoice.BackColor

                UseTaxSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.UseTax)
                If UseTaxSetting = "" Then
                    Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                Else
                    If UseTaxSetting = "0" Then
                        'Me.chkTax.Checked = False
                    Else
                        'Me.chkTax.Checked = True
                        TaxPercentSetting = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Invoice, ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                        If IsNumeric(TaxPercentSetting) = True Then
                            Me.txtTax.Text = Me.ObjCD.DigitGrouping(TaxPercentSetting, 2)
                        End If
                    End If
                End If
                '2025 Core Tax Tahap 2.2
                'txtNoFakturPajakSederhana.Visible = False
                'txtNoSeriFakturPajak.Visible = True
                'Label8.Visible = True
                'txtNoFakturPajak.Visible = True
                'llbEditNoFakturPajak.Visible = True
            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.TabStop = False
                Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
                '2025 Core Tax Tahap 2.2
                'txtNoFakturPajakSederhana.Visible = True
                'txtNoSeriFakturPajak.Visible = False
                'Label8.Visible = False
                'txtNoFakturPajak.Visible = False
                'llbEditNoFakturPajak.Visible = False
            End If

            If Me.TxtRemarks.Enabled = True Then
                Me.SetTaxRate()
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
                Else
                    Me.llbEditNoFakturPajak.Enabled = False
                    '2025 Core Tax Tahap 2.2
                    'Me.txtNoFakturPajak.Text = Me.ObjInvoice.GetNoFakturPajakStandarBaru(Me.dtpDate.Value)
                    'If Me.txtNoFakturPajak.Text = "" Then
                    '    Me.txtNoFakturPajak.Text = ".000-00.00000000"
                    'End If
                    'Me.txtNoFakturPajak.Text = Me.txtNoFakturPajak.Text.Substring(0, 8) + "00000000"
                    Me.cboTaxtype.SelectedIndex = 0
                    Me.cboTaxtype.Enabled = False
                    Me.cboTaxtype.TabStop = False
                    Me.cboPPNDipungut.Enabled = False
                    Me.cboPPNDipungut.TabStop = False
                End If
            End If
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
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                    Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                    End If
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
                    Me.IsiCboDeliveryOrder(Me.txtPrimaryCustomer.Text)
                    Me.cboCustomerAddress.Focus()
                    Me.UpdatePaymentTerm()
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
                        Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                        Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                        '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                        If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                            NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                        End If
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
                        Me.IsiCboDeliveryOrder(Me.txtPrimaryCustomer.Text)
                        Me.cboCustomerAddress.Focus()
                        Me.UpdatePaymentTerm()
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
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtNoSeriFakturPajak.Clear()
                        Me.cboDO.DataSource = Nothing
                        Me.txtCustomerCode.Focus()
                        Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                        Me.Txt_E_Faktur_Jalan.Text = ""

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
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.GetJml(Me.txtCustomerCode.Text) = 1 Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                    Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                    If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        '2025 Core Tax Tahap 2.1
                        Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                        NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                    End If
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
                    Me.IsiCboDeliveryOrder(Me.txtPrimaryCustomer.Text)
                    Me.UpdatePaymentTerm()
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
                        Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                        Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026

                        '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                        If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                            '2025 Core Tax Tahap 2.1
                            Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                            NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                        Else
                            Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                        End If
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
                        Me.IsiCboDeliveryOrder(Me.txtPrimaryCustomer.Text)
                        Me.cboCustomerAddress.Focus()
                        Me.UpdatePaymentTerm()
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtNoSeriFakturPajak.Clear()
                        Me.cboDO.DataSource = Nothing
                        Me.txtCustomerCode.Focus()
                        Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                        Me.Txt_E_Faktur_Jalan.Text = ""
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
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.cboDO.DataSource = Nothing
                Me.txtNoSeriFakturPajak.Clear()
                Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                Me.Txt_E_Faktur_Jalan.Text = ""

                '//TR : Tambahan baru Nik/Pas
                'Me.llbUpDateNPWP.Visible = True
                Me.cboNIKPASS.Visible = False
                Me.cboNIKPASS.TabStop = False
                Me.txtNoNIKPass.Visible = False
                Me.txtNoNPWP.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim i As New FrmCariMasterCustomer
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
                Me.cboCustomerAddress.SelectedIndex = 0
                Me.txtNoSeriFakturPajak.Text = Me.ObjCustomer.GetTaxCode(Me.txtPrimaryCustomer.Text)
                Me.TaxCodeDefaultDariCustomer(Me.txtNoSeriFakturPajak.Text) 'Update SMESS 2026

                '2024-09-13 IM, NPWP16 jika belum tgl berlaku, ambil yg OLD
                If Me.dtpDate.Value.Date >= ObjSetNPWP.Tglberlaku Then
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    '2025 Core Tax Tahap 2.1
                    Me.txtNoNPWP.Text = ObjSetNPWP.format16DigitSpasi(Me.txtNoNPWP.Text)
                    NPWPType = ObjSetNPWP.GetNPWPType_MasterCustomer(Me.txtPrimaryCustomer.Text)
                Else
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWPOld(Me.txtPrimaryCustomer.Text)
                End If
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
                Me.IsiCboDeliveryOrder(Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.Focus()
                Me.UpdatePaymentTerm()
                Me.txtNoNPWP.Visible = True

            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtNoSeriFakturPajak.Clear()
                Me.cboDO.DataSource = Nothing
                Me.txtCustomerCode.Focus()
                Me.ObjCD.KosongkanControl(Me.TableLayoutPanel1.Controls, True, False, False)
                Me.Txt_E_Faktur_Jalan.Text = ""
                '//TR : Tambahan baru Nik/Pas
                Me.cboNIKPASS.Visible = False
                Me.txtNoNIKPass.Visible = False
                Me.cboNIKPASS.TabStop = False
                Me.txtNoNPWP.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub UpdatePaymentTerm()
        Dim PT As Decimal = 0

        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtPrimaryCustomer.Text = "-1" Then
                    Me.txtPaymentTerm.Text = "0"
                Else
                    PT = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)

                    If CDec(Me.txtPaymentTerm.Text) = 0 Then
                        If Me.baru = True Then
                            Me.txtPaymentTerm.Text = PT
                        Else
                            If MsgBox("Update payment term?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.txtPaymentTerm.Text = PT
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub txtTax_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Enter
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTax, sender, e)
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
                If CDec(Me.txtTax.Text) > 100 Then
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
            If Me.txtPrimaryCustomer.Text <> "-1" And Me.txtPrimaryCustomer.Text <> "" Then
                If Me.TxtCustomerName.Text <> "" Then
                    Me.Text = "Sales Invoice: " + Me.txtCustomerCode.Text + " - " + Me.TxtCustomerName.Text
                Else
                    Me.Text = "Sales Invoice"
                End If
            Else
                Me.Text = "Sales Invoice"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
            If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                Me.txtCustomerCode.Focus()
                Exit Sub
            End If

            Dim ij As New Core_Module.FrmDetailAddress
            ij.vAddIdentifier = Me.txtCustomerAddIdentifier.Text
            ij.PrimaryCust = Me.txtPrimaryCustomer.Text

            If Me.TxtRemarks.Enabled = False Or Me.cboCustomerAddress.Enabled = False Then
                ij.cboCustAddCode.Enabled = False
            Else
                ij.cboCustAddCode.Enabled = True
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiCboDeliveryOrder2(ByVal vPrimaryCustomer As Double) As Boolean
        Dim row2 As DseTransaksiSalesInvoice.DataTable7Row
        Dim row1 As DseTransaksiDO.tbltrans_do_view2Row

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data yang akan digunakan sebagai acuan
            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()

            'ambil delivery order yang belum ada invoice
            Me.objDO.IsiTBLDeliveryOrderBelumAdaInvoice(Me.txtPrimaryCustomer.Text, Me.TBL)
            'tambahkan dengan delivery order yang sudah disimpan dalam object lama
            Me.ObjInvoiceItem.IsiTBLDO(Me.txtPrimaryInvoice.Text, Me.TBL2)
            While Me.TBL2.Rows.Count > 0
                row2 = Me.TBL2.Rows(0)
                Me.TBL.Rows.Add(row2.Primary_DO, row2.Do_Number)
                Me.TBL2.Rows.Remove(row2)
            End While
            'kurangkan dengan yang ada didalam dgv
            Me.IsiTBLDO(Me.TBL2)
            Dim Ketemu As Boolean = False
            While Me.TBL2.Rows.Count > 0
                Ketemu = False
                row2 = Me.TBL2.Rows(0)
                For Each row1 In Me.TBL.Rows
                    If row1.Primary_DO = row2.Primary_DO And row1.DO_Number = row2.Do_Number Then
                        Me.TBL.Rows.Remove(row1)
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    'tidak mungkin
                    MsgBox("Logic error, contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Exit Function
                End If
                Me.TBL2.Rows.Remove(row2)
            End While

            Me.cboDO.DataSource = Me.TBL.DefaultView
            Me.TBL.DefaultView.Sort = "[Do Number] ASC"
            Me.cboDO.DisplayMember = "Do Number"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function IsiCboDeliveryOrder(ByVal vPrmaryCustomer As Double) As Boolean
        Dim row2 As DseTransaksiSalesInvoice.DataTable7Row
        Dim row1 As DseTransaksiDO.tbltrans_do_view2Row

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'kosongkan data yang akan digunakan sebagai acuan
            Me.TBL.Rows.Clear()
            Me.TBL2.Rows.Clear()
            Me.dgvItem.Rows.Clear()

            'ambil delivery order yang belum ada invoice
            Me.objDO.IsiTBLDeliveryOrderBelumAdaInvoice(Me.txtPrimaryCustomer.Text, Me.TBL)
            'tambahkan dengan delivery order yang sudah disimpan dalam object lama
            Me.ObjInvoiceItem.IsiTBLDO(Me.txtPrimaryInvoice.Text, Me.TBL2)
            While Me.TBL2.Rows.Count > 0
                row2 = Me.TBL2.Rows(0)
                Me.TBL.Rows.Add(row2.Primary_DO, row2.Do_Number)
                Me.TBL2.Rows.Remove(row2)
            End While
            'kurangkan dengan yang ada didalam dgv
            Me.IsiTBLDO(Me.TBL2)
            Dim Ketemu As Boolean = False
            While Me.TBL2.Rows.Count > 0
                Ketemu = False
                row2 = Me.TBL2.Rows(0)
                For Each row1 In Me.TBL.Rows
                    If row1.Primary_DO = row2.Primary_DO And row1.DO_Number = row2.Do_Number Then
                        Me.TBL.Rows.Remove(row1)
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    'tidak mungkin
                    MsgBox("Logic error, contact your vendor if you see this message", _
                    MsgBoxStyle.Information, "Information")
                    Exit Function
                End If
                Me.TBL2.Rows.Remove(row2)
            End While

            Me.cboDO.DataSource = Me.TBL.DefaultView
            Me.TBL.DefaultView.Sort = "[Do Number] ASC"
            Me.cboDO.DisplayMember = "Do Number"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Private Sub IsiTBLDO(ByVal TBL As DseTransaksiSalesInvoice.DataTable7DataTable)
        Dim row As DataGridViewRow
        Dim PrimaryDummy As Double = -1

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryDO.Index).Value <> PrimaryDummy Then
                    PrimaryDummy = row.Cells(Me.ColPrimaryDO.Index).Value
                    TBL.Rows.Add(row.Cells(Me.ColPrimaryDO.Index).Value, Me.objDO.GetDONO(row.Cells(Me.ColPrimaryDO.Index).Value))
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetSalesPerson()
        Dim Tbl As New DseTransaksiSalesInvoice.DataTable51DataTable
        Dim row As DseTransaksiSalesInvoice.DataTable51Row
        Dim Dummy As String = ""
        Dim rowDGV As DataGridViewRow
        Dim SetAutosize As Boolean = False
        Dim ListPrimarySalesP As New List(Of Double)
        Dim PrimarySalesPerson As Double = -1
        Dim SalesPersonCode As String = ""
        Dim SalesPersonName As String = ""

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowDGV.Cells(Me.ColPrimarySalesPerson.Index).Value) = True Then
                    rowDGV.Cells(Me.ColPrimarySalesPerson.Index).Value = -1
                End If
                PrimarySalesPerson = rowDGV.Cells(Me.ColPrimarySalesPerson.Index).Value

                Tbl.DefaultView.RowFilter = "[Primary SO] = " + rowDGV.Cells(Me.ColPrimarySO.Index).Value.ToString + _
                " AND [Primary DO] = " + rowDGV.Cells(Me.ColPrimaryDO.Index).Value.ToString
                If Tbl.DefaultView.Count = 0 Then
                    row = Tbl.NewDataTable51Row
                    row.Primary_SO = rowDGV.Cells(Me.ColPrimarySO.Index).Value
                    row.Primary_DO = rowDGV.Cells(Me.ColPrimaryDO.Index).Value
                    If rowDGV.Cells(Me.ColPrimarySalesPerson.Index).Value <> -1 Then
                        row.Primary_Sales_Person = rowDGV.Cells(Me.ColPrimarySalesPerson.Index).Value
                    Else
                        row.Primary_Sales_Person = Me.ObjSO.GetPrimarySalesPerson(row.Primary_SO)
                    End If
                    If row.Primary_Sales_Person <> -1 Then
                        row.Sales_Person_Code = Me.ObjSalesPerson.GetCode(row.Primary_Sales_Person)
                        row.Sales_Person_Name = Me.ObjSalesPerson.GetName(row.Primary_Sales_Person)
                        Tbl.Rows.Add(row)
                    End If
                End If
            Next
            'For Each rowDGV In Me.dgvItem.Rows
            '    Tbl.DefaultView.RowFilter = "[Primary SO] = " + rowDGV.Cells(Me.ColPrimarySO.Index).Value.ToString
            '    If Tbl.DefaultView.Count = 0 Then
            '        row = Tbl.NewDataTable51Row
            '        row.Primary_SO = rowDGV.Cells(Me.ColPrimarySO.Index).Value
            '        row.Primary_Sales_Person = Me.ObjSO.GetPrimarySalesPerson(row.Primary_SO)
            '        If row.Primary_Sales_Person <> -1 Then
            '            row.Sales_Person_Code = Me.ObjSalesPerson.GetCode(row.Primary_Sales_Person)
            '            row.Sales_Person_Name = Me.ObjSalesPerson.GetName(row.Primary_Sales_Person)
            '            Tbl.Rows.Add(row)
            '        End If
            '    End If
            'Next

            If Tbl.Rows.Count = 1 Then
                row = Tbl.Rows(0)
                Me.ColSalesPerson.Visible = False
                Me.TxtSalesPersonCode.Text = row.Sales_Person_Code
                Me.TxtSalesPersonName.Text = row.Sales_Person_Name
            ElseIf Tbl.Rows.Count = 0 Then
                Me.ColSalesPerson.Visible = False
                Me.TxtSalesPersonCode.Text = "-"
                Me.TxtSalesPersonName.Text = "-"
            Else
                For Each row In Tbl.Rows
                    If ListPrimarySalesP.Contains(row.Primary_Sales_Person) = False Then
                        ListPrimarySalesP.Add(row.Primary_Sales_Person)
                        If Dummy = "" Then
                            Dummy = row.Sales_Person_Name
                        Else
                            Dummy = Dummy + ", " + row.Sales_Person_Name
                        End If
                    End If
                Next
                If ListPrimarySalesP.Count > 1 Then
                    Me.TxtSalesPersonCode.Text = "Multiple"
                    Me.TxtSalesPersonName.Text = Dummy

                    Me.ColSalesPerson.Visible = True
                Else
                    Me.ColSalesPerson.Visible = False
                    row = Tbl.Rows(0)
                    Me.TxtSalesPersonCode.Text = row.Sales_Person_Code
                    Me.TxtSalesPersonName.Text = row.Sales_Person_Name
                End If
            End If

            If Me.dgvItem.AutoSizeColumnsMode <> DataGridViewAutoSizeColumnsMode.None Then
                SetAutosize = True
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            End If
            For Each rowDGV In Me.dgvItem.Rows
                Tbl.DefaultView.RowFilter = "[Primary SO] = " + rowDGV.Cells(Me.ColPrimarySO.Index).Value.ToString + _
                " AND [Primary DO] = " + rowDGV.Cells(Me.ColPrimaryDO.Index).Value.ToString
                If Tbl.DefaultView.Count = 1 Then
                    row = Tbl.DefaultView(0).Row
                    rowDGV.Cells(Me.ColPrimarySalesPerson.Index).Value = row.Primary_Sales_Person
                    rowDGV.Cells(Me.ColSalesPerson.Index).Value = row.Sales_Person_Code + " - " + row.Sales_Person_Name
                End If
            Next
            Me.ColSalesPerson.Visible = True
            If SetAutosize = True Then
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Tbl.Dispose()
            ListPrimarySalesP = Nothing
        End Try
    End Sub

    Private Function ApakahShipToSama(ByVal vPrimaryCustomer As Double, ByVal vPrimaryCustShipTo As Double) As Boolean
        Try
            'Me.PrimaryCustomerShipTo()
            If vPrimaryCustomer <> -1 And vPrimaryCustShipTo = -1 Then
                vPrimaryCustShipTo = vPrimaryCustomer
            End If
            If vPrimaryCustShipTo <> Me.PrimaryCustomerShipTo Then 'Or vPrimaryCustomer <> Me.PrimaryCustomerShipTo Then
                If Me.dgvItem.Rows.Count > 0 Then
                    MsgBox("You have to input DO data that has the same ship to customer.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        If Me.cboDO.Items.Count > 0 Then
            Dim TBL1 As New DseTransaksiDO.tbltrans_do_view3_newDataTable
            Dim DTA1 As New DseTransaksiDOTableAdapters.tbltrans_do_view3_newTableAdapter()
            Dim rowItemDO As DseTransaksiDO.tbltrans_do_view3_newRow
            Dim rowDoNo As DseTransaksiDO.tbltrans_do_view2Row
            Dim Dataview As DataRowView
            Dim NetPriceExclude As Double = 0
            Dim Amount As Double = 0
            Dim CurrencyAcuan As String = ""
            Dim rowDGV As DataGridViewRow
            Dim DiscPerct As Integer = 0

            Try
                '2025 Core Tax Tahap 2.1
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                If Me.dtpDate.Value >= #1/1/2025# Then
                    Me.ButtonInput_MiddlePoint()
                    Exit Sub
                End If
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                If Me.txtPrimaryCurrency.Text = "-1" Or Me.txtPrimaryCurrency.Text = "" Then
                    MsgBox("Please specify the currency for the invoice.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCurrencyCode.Focus()
                    Exit Sub
                End If
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Dataview = Me.cboDO.SelectedItem
                rowDoNo = Dataview.Row
                'cek periode do dan periode stock out harus sama dengan periode invoice
                If Format(Me.objDO.GetTanggalStockOut(rowDoNo.Primary_DO), "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                    MsgBox("You can only enter DO with stock out dated from " + Format(Me.dtpDate.Value.Date, "MMM yyyy"), MsgBoxStyle.Information, _
                    "Information")
                    Exit Sub
                End If

                If Me.dgvItem.Rows.Count > 0 And Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoice_InputDifferentBillTo") = False Then
                    Dim DtaBillTo As New DseTransaksiDOTableAdapters.tbltrans_doTableAdapter
                    DtaBillTo.Connection.ConnectionString = Core_Module.GetConstring
                    If DtaBillTo.ScalarQuery_GetIdentifierAddressBillTo(Me.dgvItem.Item(Me.ColPrimaryDO.Index, 0).Value) <> DtaBillTo.ScalarQuery_GetIdentifierAddressBillTo(rowDoNo.Primary_DO) Then
                        MsgBox("You have to input DO data that has the same bill address.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If

                'cek tanggal stock out --> ga usah, karena tanggal stock out juga dibatasi harus sama bulannya dengan tanggal
                'DO

                'isikan tabel acuan
                'DTA1.FillBy_PrimaryDO(TBL1, rowDoNo.Primary_DO)
                Me.IsiItemDO2(TBL1, rowDoNo.Primary_DO)
                'cek currency
                If TBL1.Rows.Count > 0 And Me.dgvItem.RowCount > 0 Then
                    CurrencyAcuan = Me.dgvItem.Item(Me.ColCurrency.Index, 0).Value
                    rowItemDO = TBL1.Rows(0)
                    If rowItemDO.Currency_Code <> CurrencyAcuan Then
                        MsgBox("To enter 2 DO in one invoice the currency code must be similar.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If

                'set currency code jika, masih kosong dgvnya
                If Me.dgvItem.RowCount = 0 Then
                    If TBL1.Rows.Count > 0 Then
                        rowItemDO = TBL1.Rows(0)
                        Me.txtPrimaryCurrency.Text = Me.ObjSO.GetPrimaryCurrency(rowItemDO.Primary_SO)
                        Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetTaxRate()
                        Me.SetDefaultPrecision()
                    End If
                End If

                'untuk setiap baris dalam tabel
                For Each rowItemDO In TBL1.Rows
                    'hitung ulang exclude price jika perlu
                    With rowItemDO
                        If .Price_Exclude = True Then
                            .Unit_Price_Exclude = .Unit_Price
                            NetPriceExclude = .Nett_Price
                            If .Secondary_Qty <= 0 Then
                                If .price_by_multi_unit = 1 Then
                                    Amount = Me.ObjCD.SetDecimal(NetPriceExclude * .Total_Qty_Multi, .Default_Precision)
                                Else
                                    Amount = Me.ObjCD.SetDecimal(NetPriceExclude * .Total_Qty, .Default_Precision)
                                End If
                            Else
                                Amount = Me.ObjCD.SetDecimal(NetPriceExclude * .Total_Qty_Multi * .Secondary_Qty, .Default_Precision)
                            End If
                        Else
                            If Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                                .Unit_Price_Exclude = Me.ObjCD.SetDecimal(.Unit_Price * 100 / (100 + .Tax), .Default_Precision)
                                DiscPerct = .Disc_From_Perct

                                If DiscPerct = 1 Then
                                    .Disc_Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                                                                    ((CDec(.Unit_Price_Exclude) * CDec(.Discount))) / 100, DefaultPrecision))
                                    NetPriceExclude = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                                    CDec(.Unit_Price_Exclude) - CDec(.Disc_Amount) _
                                    , DefaultPrecision))
                                    .Nett_Price = .Unit_Price_Exclude - .Disc_Amount
                                    NetPriceExclude = Me.ObjCD.SetDecimal(.Unit_Price_Exclude * (100 - .Discount) / 100, .Default_Precision)
                                Else
                                    NetPriceExclude = Me.ObjCD.SetDecimal(.Nett_Price * 100 / (100 + .Tax), .Default_Precision)
                                    .Disc_Amount = .Unit_Price_Exclude - NetPriceExclude
                                    .Nett_Price = .Unit_Price_Exclude - .Disc_Amount
                                End If

                                If .Secondary_Qty <= 0 Then
                                    If .price_by_multi_unit = 1 Then
                                        Amount = Me.ObjCD.SetDecimal(.Total_Qty_Multi * NetPriceExclude, .Default_Precision)
                                    Else
                                        Amount = Me.ObjCD.SetDecimal(.Total_Qty * NetPriceExclude, .Default_Precision)
                                    End If
                                Else
                                    Amount = Me.ObjCD.SetDecimal(.Total_Qty_Multi * .Secondary_Qty * _
                                    NetPriceExclude, .Default_Precision)
                                End If
                            End If
                        End If

                        ''update bill to address based on selected SO
                        ''update ship to address based on selected SO and it's readonly
                        Dim DtaSOAddress As New DseTransaksiSOTableAdapters.tbltrans_so3TableAdapter
                        Dim TblSOAddress As New DseTransaksiSO.tbltrans_so3DataTable
                        Dim RowSOAddress As DseTransaksiSO.tbltrans_so3Row
                        DtaSOAddress.Connection.ConnectionString = Core_Module.GetConstring
                        If DtaSOAddress.Fill(TblSOAddress, .Primary_SO) > 0 Then
                            RowSOAddress = TblSOAddress.Rows(0)
                            'terkait ship to
                            If Me.ApakahShipToSama(Me.txtPrimaryCustomer.Text, RowSOAddress.primary_customer_ship_to) = False Then
                                Exit Sub
                            End If

                            If Me.cboCustomerAddress.Items.Count > 0 Then
                                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(RowSOAddress.Indentifier_Alamat_Customer)
                            End If

                            If RowSOAddress.primary_customer_ship_to = -1 Then
                                Me.PrimaryCustomerShipTo = Me.txtPrimaryCustomer.Text
                                Me.IdentifierAddressCustomerShipTo = RowSOAddress.identifier_address_ship_to
                            Else
                                Me.PrimaryCustomerShipTo = RowSOAddress.primary_customer_ship_to
                                Me.IdentifierAddressCustomerShipTo = RowSOAddress.identifier_address_customer_ship_to
                            End If
                            For Each rowHeaderCost In TBLHeaderCost
                                If RowSOAddress.primary_customer_ship_to = -1 Then
                                    rowHeaderCost.primary_cust_ship_to = Me.txtPrimaryCustomer.Text
                                    rowHeaderCost.identifier_cust_ship_to = RowSOAddress.identifier_address_ship_to
                                Else
                                    rowHeaderCost.primary_cust_ship_to = RowSOAddress.primary_customer_ship_to
                                    rowHeaderCost.identifier_cust_ship_to = RowSOAddress.identifier_address_customer_ship_to
                                End If
                            Next
                        End If

                        'tambahkan ke dalam dgv
                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        rowDGV.Cells(Me.ColNo0.Index).Value = "0"
                        rowDGV.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(Me.ColDescription.Index).Value = .Description
                        rowDGV.Cells(Me.ColCustCodeNo.Index).Value = ._Customer_s_Code_No
                        rowDGV.Cells(Me.ColCustDesc.Index).Value = ._Customer_s_Desc
                        rowDGV.Cells(Me.ColSONO.Index).Value = .Sales_Order_No
                        rowDGV.Cells(Me.ColCustPONo.Index).Value = .Cust_PO_No
                        rowDGV.Cells(Me.ColDONO.Index).Value = .DeliveryOrderNo
                        rowDGV.Cells(Me.ColStockOutNo.Index).Value = .Stock_Out_Number
                        rowDGV.Cells(Me.ColQtyMaster.Index).Value = .Qty_Master
                        rowDGV.Cells(Me.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDGV.Cells(Me.ColMasterUnits.Index).Value = .Master_Unit
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rowDGV.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty, ._Decimal)
                        rowDGV.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(Me.ColSecQty.Index).Value = Me.ObjCD.DigitGrouping(.Secondary_Qty)
                        rowDGV.Cells(Me.ColSecUnit.Index).Value = .Secondary_Unit
                        rowDGV.Cells(Me.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(.Unit_Price_Exclude, .Default_Precision)
                        rowDGV.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(.Discount, 2)
                        rowDGV.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(NetPriceExclude, .Default_Precision)
                        rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, .Default_Precision)
                        rowDGV.Cells(Me.ColRemarksItem.Index).Value = ""
                        rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(Me.ColPrimarySO.Index).Value = .Primary_SO
                        rowDGV.Cells(Me.ColPrimaryDO.Index).Value = .Primary_DO
                        rowDGV.Cells(Me.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Units
                        rowDGV.Cells(Me.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        rowDGV.Cells(Me.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDGV.Cells(Me.ColLotNumber.Index).Value = .Batch_No
                        rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        rowDGV.Cells(Me.ColDiscountAmount.Index).Value = Me.ObjCD.DigitGrouping(.Disc_Amount, .Default_Precision)
                        If .Expiry_Date = #1/1/1950# Then
                            rowDGV.Cells(Me.ColExpiry.Index).Value = "-"
                        Else
                            rowDGV.Cells(Me.ColExpiry.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If

                        rowDGV.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi))
                        rowDGV.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(Me.ColUnitsCodeMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi))
                        rowDGV.Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, ._Decimal)
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)) _
                            & " " & .Units_Code_Multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, ._Decimal) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(Me.ColUnitsCodeMulti.Index).Value
                        Else
                            rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(Me.ColUnit.Index).Value
                        End If
                    End With
                Next

                'tambahkan exchange rate data
                Me.SetCalculateBy()
                'penomoran
                Me.ObjCD.SetNumber(Me.dgvItem)
                Me.ObjCD.SetNumber(Me.dgvExchageRates)
                Me.FormatDGV()
                'hilangkan do dari item
                Me.TBL.Rows.Remove(rowDoNo)
                'text customer sudah tidak bisa diganti
                Me.txtCustomerCode.ReadOnly = True
                Me.btnFindCustomer.Enabled = False

                Me.SetSalesPerson()

                If Me.dgvItem.RowCount > 0 Then
                    Me.AddDP_Procedure()
                    IsiDGVReturn()
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                DTA1.Dispose()
                TBL1.Dispose()
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Windows.Forms.Cursor.Current = Cursors.Arrow
            End Try
        End If
    End Sub

    '2025 Core Tax Tahap 2.1, fungsi button input diubah
    Private Sub ButtonInput_MiddlePoint()
        If Me.cboDO.Items.Count > 0 Then
            Dim TBL1 As New DseTransaksiDO.tbltrans_do_view3_newDataTable
            Dim DTA1 As New DseTransaksiDOTableAdapters.tbltrans_do_view3_newTableAdapter()
            Dim rowItemDO As DseTransaksiDO.tbltrans_do_view3_newRow
            Dim rowDoNo As DseTransaksiDO.tbltrans_do_view2Row
            Dim Dataview As DataRowView
            Dim NetPriceExclude As Decimal = 0
            Dim Amount As Double = 0
            Dim CurrencyAcuan As String = ""
            Dim rowDGV As DataGridViewRow
            Dim DiscPerct As Integer = 0
            '\\TR==GlobalDiscount
            Dim GlobalDiscPerc As Decimal = 0
            Dim GlobalDiscAmount As Decimal = 0
            Dim Disc1Perc As Decimal = 0
            Dim Disc1Amount As Decimal = 0
            Dim UseTax As Integer = 1
            Dim UnitPriceExclude As Decimal = 0
            Try
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                If Me.txtPrimaryCurrency.Text = "-1" Or Me.txtPrimaryCurrency.Text = "" Then
                    MsgBox("Please specify the currency for the invoice.", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCurrencyCode.Focus()
                    Exit Sub
                End If
                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Dataview = Me.cboDO.SelectedItem
                rowDoNo = Dataview.Row
                'cek periode do dan periode stock out harus sama dengan periode invoice
                If Format(Me.objDO.GetTanggalStockOut(rowDoNo.Primary_DO), "MMM yyyy") <> Format(Me.dtpDate.Value.Date, "MMM yyyy") Then
                    MsgBox("You can only enter DO with stock out dated from " + Format(Me.dtpDate.Value.Date, "MMM yyyy"), MsgBoxStyle.Information, _
                    "Information")
                    Exit Sub
                End If

                If Me.dgvItem.Rows.Count > 0 And Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoice_InputDifferentBillTo") = False Then
                    Dim DtaBillTo As New DseTransaksiDOTableAdapters.tbltrans_doTableAdapter
                    DtaBillTo.Connection.ConnectionString = Core_Module.GetConstring
                    If DtaBillTo.ScalarQuery_GetIdentifierAddressBillTo(Me.dgvItem.Item(Me.ColPrimaryDO.Index, 0).Value) <> DtaBillTo.ScalarQuery_GetIdentifierAddressBillTo(rowDoNo.Primary_DO) Then
                        MsgBox("You have to input DO data that has the same bill address.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If

                'cek tanggal stock out --> ga usah, karena tanggal stock out juga dibatasi harus sama bulannya dengan tanggal
                'DO

                'isikan tabel acuan
                'DTA1.FillBy_PrimaryDO(TBL1, rowDoNo.Primary_DO)
                Me.IsiItemDO2(TBL1, rowDoNo.Primary_DO)
                'cek currency
                If TBL1.Rows.Count > 0 And Me.dgvItem.RowCount > 0 Then
                    CurrencyAcuan = Me.dgvItem.Item(Me.ColCurrency.Index, 0).Value
                    '\\TR==GlobalDiscount
                    UseTax = Me.dgvItem.Item(Me.ColUseTax.Index, 0).Value
                    rowItemDO = TBL1.Rows(0)
                    If rowItemDO.Currency_Code <> CurrencyAcuan Then
                        MsgBox("To enter 2 DO in one invoice the currency code must be similar.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    ''\\TR==GlobalDiscount
                    'If rowItemDO.Use_Tax <> UseTax Then
                    '    '   a. % pajak yang digunakan sama
                    '    MsgBox("To input some DO in one invoice, the tax must be identical", _
                    '    MsgBoxStyle.Information, "Information")
                    '    Exit Sub
                    'End If
                End If

                'set currency code jika, masih kosong dgvnya
                If Me.dgvItem.RowCount = 0 Then
                    If TBL1.Rows.Count > 0 Then
                        rowItemDO = TBL1.Rows(0)
                        Me.txtPrimaryCurrency.Text = Me.ObjSO.GetPrimaryCurrency(rowItemDO.Primary_SO)
                        Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetTaxRate()
                        Me.SetDefaultPrecision()
                    End If
                End If

                'untuk setiap baris dalam tabel
                For Each rowItemDO In TBL1.Rows
                    'hitung ulang exclude price jika perlu
                    With rowItemDO
                        '\\TR==GlobalDiscount
                        Disc1Perc = .Discount
                        Disc1Amount = .Disc_Amount
                        NetPriceExclude = .Nett_Price
                        UnitPriceExclude = .Unit_Price

                        '2025-07-16 dikomen, belum masuk global discount
                        'If Me.chkUseGlobalDisc.Checked = True Then
                        '    If Me.txtGlobalIsDiscMulti.Text = "0" Then
                        '        GlobalDiscAmount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                        '                           ((NetPriceExclude * CDec(Me.txtGlobalDisc.Text))) / 100, DefaultPrecision))
                        '        GlobalDiscPerc = Me.txtGlobalDisc.Text
                        '    Else
                        '        If .Primary_SO_Ref = -1 Then
                        '            If Me.GetAmountDiscBertingkat(GlobalDiscAmount, Me.txtGlobalDiscString.Text, NetPriceExclude, DefaultPrecision, True) = False Then
                        '                'Me.dgvItem.Rows(row.Index).Selected = True
                        '                'Me.txtGlobalDiscStrnig.Text = row.Cells(Me.ColDiscountGlobalString.Index).Value
                        '                Exit Sub
                        '            Else
                        '                If NetPriceExclude <> 0 Then
                        '                    GlobalDiscPerc = Me.ObjCD.DigitGrouping((GlobalDiscAmount / NetPriceExclude) * 100, 2)
                        '                End If
                        '            End If
                        '        Else
                        '            GlobalDiscAmount = 0
                        '            GlobalDiscPerc = 0
                        '        End If
                        '    End If

                        '    NetPriceExclude = NetPriceExclude - GlobalDiscAmount
                        '    .Nett_Price = NetPriceExclude
                        '    .Disc_Amount = Me.ObjCD.DigitGrouping(UnitPriceExclude - NetPriceExclude, DefaultPrecision)
                        '    .Discount = Me.ObjCD.DigitGrouping((CDec(.Disc_Amount) / UnitPriceExclude) * 100, 2)
                        'End If

                        If .Secondary_Qty <= 0 Then
                            If .price_by_multi_unit = 1 Then
                                Amount = Me.ObjCD.SetDecimal(NetPriceExclude * .Total_Qty_Multi, .Default_Precision)
                            Else
                                Amount = Me.ObjCD.SetDecimal(NetPriceExclude * .Total_Qty, .Default_Precision)
                            End If
                        Else
                            Amount = Me.ObjCD.SetDecimal(NetPriceExclude * .Total_Qty_Multi * .Secondary_Qty, .Default_Precision)
                        End If

                        '2025-07-16 dikomen, beda source code
                        ''update bill to address based on selected SO
                        ''update ship to address based on selected SO and it's readonly
                        Dim DtaSOAddress As New DseTransaksiSOTableAdapters.tbltrans_so3TableAdapter
                        Dim TblSOAddress As New DseTransaksiSO.tbltrans_so3DataTable
                        Dim RowSOAddress As DseTransaksiSO.tbltrans_so3Row
                        DtaSOAddress.Connection.ConnectionString = Core_Module.GetConstring
                        If DtaSOAddress.Fill(TblSOAddress, .Primary_SO) > 0 Then
                            RowSOAddress = TblSOAddress.Rows(0)
                            'terkait ship to
                            If Me.ApakahShipToSama(Me.txtPrimaryCustomer.Text, RowSOAddress.primary_customer_ship_to) = False Then
                                Exit Sub
                            End If

                            If Me.cboCustomerAddress.Items.Count > 0 Then
                                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(RowSOAddress.Indentifier_Alamat_Customer)
                            End If

                            If RowSOAddress.primary_customer_ship_to = -1 Then
                                Me.PrimaryCustomerShipTo = Me.txtPrimaryCustomer.Text
                                Me.IdentifierAddressCustomerShipTo = RowSOAddress.identifier_address_ship_to
                            Else
                                Me.PrimaryCustomerShipTo = RowSOAddress.primary_customer_ship_to
                                Me.IdentifierAddressCustomerShipTo = RowSOAddress.identifier_address_customer_ship_to
                            End If
                            For Each rowHeaderCost In TBLHeaderCost
                                If RowSOAddress.primary_customer_ship_to = -1 Then
                                    rowHeaderCost.primary_cust_ship_to = Me.txtPrimaryCustomer.Text
                                    rowHeaderCost.identifier_cust_ship_to = RowSOAddress.identifier_address_ship_to
                                Else
                                    rowHeaderCost.primary_cust_ship_to = RowSOAddress.primary_customer_ship_to
                                    rowHeaderCost.identifier_cust_ship_to = RowSOAddress.identifier_address_customer_ship_to
                                End If
                            Next
                        End If

                        If Me.dtpDate.Value < #10/9/2025# Or Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = True Then
                            If .Price_Exclude = True Then
                                .Unit_Price_Exclude = .Unit_Price
                                NetPriceExclude = .Nett_Price
                                If .Secondary_Qty <= 0 Then
                                    Amount = NetPriceExclude * .Total_Qty_Multi
                                Else
                                    Amount = NetPriceExclude * .Total_Qty_Multi * .Secondary_Qty
                                End If
                            Else
                                If Me.MetodeHitungInclude = EnumHitungInclude.NettPriceInclude Then
                                    '.Unit_Price_Exclude = Me.ObjCD.SetDecimal(.Unit_Price * 100 / (100 + .Tax), .Default_Precision)
                                    'NetPriceExclude = Me.ObjCD.SetDecimal(.Unit_Price_Exclude * (100 - .Discount) / 100, .Default_Precision)
                                    'If .Secondary_Qty <= 0 Then
                                    'Amount = Me.ObjCD.SetDecimal(.Total_Qty * (.Unit_Price_Exclude * _
                                    '((100 - .Discount) / 100)), .Default_Precision)
                                    'Else
                                    'Amount = Me.ObjCD.SetDecimal(.Total_Qty * .Secondary_Qty * _
                                    '(.Unit_Price_Exclude * ((100 - .Discount) / 100)), .Default_Precision)
                                    'End If
                                    'hitung nett price, karena ini harga include, untuk disc from perct, dianggap 1
                                    .Unit_Price_Exclude = Me.ObjCD.SetDecimal(.Unit_Price * 100 / (100 + .Tax), .Default_Precision)
                                    DiscPerct = .Disc_From_Perct
                                    If DiscPerct = 1 Then
                                        .Disc_Amount = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                                                                        ((CDec(.Unit_Price_Exclude) * CDec(.Discount))) / 100, DefaultPrecision))
                                        NetPriceExclude = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                                        CDec(.Unit_Price_Exclude) - CDec(.Disc_Amount) _
                                        , DefaultPrecision))

                                        .Nett_Price = .Unit_Price_Exclude - .Disc_Amount
                                        NetPriceExclude = Me.ObjCD.SetDecimal(.Unit_Price_Exclude * (100 - .Discount) / 100, .Default_Precision)
                                    Else
                                        NetPriceExclude = Me.ObjCD.SetDecimal(.Nett_Price * 100 / (100 + .Tax), .Default_Precision)
                                        .Disc_Amount = .Unit_Price_Exclude - NetPriceExclude
                                        .Nett_Price = .Unit_Price_Exclude - .Disc_Amount
                                    End If

                                    If .Secondary_Qty <= 0 Then
                                        Amount = Me.ObjCD.SetDecimal(.Total_Qty_Multi * NetPriceExclude, .Default_Precision)
                                    Else
                                        Amount = Me.ObjCD.SetDecimal(.Total_Qty_Multi * .Secondary_Qty * _
                                        NetPriceExclude, .Default_Precision)
                                    End If
                                    'Else
                                    '    .Unit_Price_Exclude = Me.ObjCD.SetDecimal(.Unit_Price * 100 / (100 + .Tax), .Default_Precision)
                                    '    NetPriceExclude = Me.ObjCD.SetDecimal(.Unit_Price_Exclude * (100 - .Discount) / 100, .Default_Precision)
                                    '    If .Secondary_Qty <= 0 Then
                                    '        Amount = Me.ObjCD.SetDecimal(.Total_Qty * .Unit_Price * 100 / (100 + .Tax) * (100 - .Discount) / 100, .Default_Precision)
                                    '    Else
                                    '        Amount = Me.ObjCD.SetDecimal(.Total_Qty * .Secondary_Qty * .Unit_Price * 100 / (100 + .Tax) * (100 - .Discount) / 100, .Default_Precision)
                                    '    End If
                                End If
                            End If
                        End If

                        'tambahkan ke dalam dgv
                        Me.dgvItem.Rows.Add()
                        rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        rowDGV.Cells(Me.ColNo0.Index).Value = "0"
                        rowDGV.Cells(Me.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(Me.ColDescription.Index).Value = .Description
                        rowDGV.Cells(Me.ColCustCodeNo.Index).Value = ._Customer_s_Code_No
                        rowDGV.Cells(Me.ColCustDesc.Index).Value = ._Customer_s_Desc
                        rowDGV.Cells(Me.ColSONO.Index).Value = .Sales_Order_No
                        rowDGV.Cells(Me.ColCustPONo.Index).Value = .Cust_PO_No
                        rowDGV.Cells(Me.ColDONO.Index).Value = .DeliveryOrderNo
                        rowDGV.Cells(Me.ColStockOutNo.Index).Value = .Stock_Out_Number
                        rowDGV.Cells(Me.ColQtyMaster.Index).Value = .Qty_Master
                        rowDGV.Cells(Me.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDGV.Cells(Me.ColMasterUnits.Index).Value = .Master_Unit
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty, ._Decimal)
                        rowDGV.Cells(Me.ColTotalQty.Index).Value = Me.ObjCD.DigitGrouping(.Total_Qty, ._Decimal)
                        rowDGV.Cells(Me.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(Me.ColSecQty.Index).Value = Me.ObjCD.DigitGrouping(.Secondary_Qty)
                        rowDGV.Cells(Me.ColSecUnit.Index).Value = .Secondary_Unit
                        rowDGV.Cells(Me.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(Me.ColUnitPriceInclude.Index).Value = Me.ObjCD.DigitGrouping(UnitPriceExclude, .Default_Precision)
                        rowDGV.Cells(Me.ColDiscountInclude.Index).Value = Me.ObjCD.DigitGrouping(Disc1Perc, 2)
                        rowDGV.Cells(Me.ColNettPriceInclude.Index).Value = Me.ObjCD.DigitGrouping(NetPriceExclude, .Default_Precision)
                        rowDGV.Cells(Me.ColAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(Amount, .Default_Precision)

                        If Me.dtpDate.Value < #10/9/2025# Or Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = True Then
                            rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(.Unit_Price_Exclude, .Default_Precision)
                            rowDGV.Cells(Me.ColDiscount.Index).Value = .Discount
                            rowDGV.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(NetPriceExclude, .Default_Precision)
                            rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, .Default_Precision)
                            rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                            rowDGV.Cells(Me.ColDiscountAmount.Index).Value = Me.ObjCD.DigitGrouping(.Disc_Amount, .Default_Precision)
                        End If

                        rowDGV.Cells(Me.ColRemarksItem.Index).Value = ""
                        rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(Me.ColPrimarySO.Index).Value = .Primary_SO
                        rowDGV.Cells(Me.ColPrimaryDO.Index).Value = .Primary_DO
                        rowDGV.Cells(Me.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Units
                        rowDGV.Cells(Me.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        rowDGV.Cells(Me.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDGV.Cells(Me.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDGV.Cells(Me.ColLotNumber.Index).Value = .Batch_No

                        rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        rowDGV.Cells(Me.ColDiscAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(.Disc_Amount, .Default_Precision)
                        If .Expiry_Date = #1/1/1950# Then
                            rowDGV.Cells(Me.ColExpiry.Index).Value = "-"
                        Else
                            rowDGV.Cells(Me.ColExpiry.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        End If

                        rowDGV.Cells(Me.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi))
                        rowDGV.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(Me.ColUnitsCodeMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(Me.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi))
                        rowDGV.Cells(Me.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, ._Decimal)
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(.Ratio_Reference, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)) _
                            & " " & .Units_Code_Multi & " = " & Me.ObjCD.DigitGrouping(.Ratio_Default, ._Decimal) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(Me.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(Me.ColUnitsCodeMulti.Index).Value
                        Else
                            rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(Me.ColUnit.Index).Value
                        End If

                        '\\TR==GlobalDiscount
                        rowDGV.Cells(Me.ColDiscountGlobalStringInclude.Index).Value = "" 'Me.txtGlobalDiscString.Text
                        rowDGV.Cells(Me.ColDiscount1AmountInclude.Index).Value = Me.ObjCD.DigitGrouping(Disc1Amount, .Default_Precision)
                        rowDGV.Cells(Me.ColDiscount1Include.Index).Value = Me.ObjCD.DigitGrouping(Disc1Perc, 2)
                        rowDGV.Cells(Me.ColDiscount1PercInclude.Index).Value = .Disc_From_Perct

                        '2025-07-16 dikomen, belum masuk global discount
                        'If Me.chkUseGlobalDisc.Checked = True Then
                        '    rowDGV.Cells(Me.ColDiscountGlobalPercInclude.Index).Value = Me.ObjCD.DigitGrouping(GlobalDiscPerc, 2)
                        '    rowDGV.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value = Me.ObjCD.DigitGrouping(GlobalDiscAmount, .Default_Precision)
                        '    rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = "0"
                        'Else
                        rowDGV.Cells(Me.ColDiscountGlobalPercInclude.Index).Value = "0"
                        rowDGV.Cells(Me.ColDiscountGlobalAmountInclude.Index).Value = "0"
                        rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = .Disc_From_Perct
                        'End If

                        '\\TR==GlobalDiscount
                        If .Use_VAT = "1" Then
                            rowDGV.Cells(Me.ColIsVAT.Index).Value = Boolean.TrueString
                        Else
                            rowDGV.Cells(Me.ColIsVAT.Index).Value = Boolean.FalseString
                        End If
                        rowDGV.Cells(Me.ColUseTax.Index).Value = .Use_Tax

                        If .Price_Exclude = True Then
                            rowDGV.Cells(Me.ColPriceCond.Index).Value = "Exclude"
                        Else
                            rowDGV.Cells(Me.ColPriceCond.Index).Value = "Include"
                        End If
                        rowDGV.Cells(Me.ColPrimarySORef.Index).Value = -1

                        'aul
                        rowDGV.Cells(Me.colPrimaryCostAcc.Index).Value = -1
                        rowDGV.Cells(Me.colCostAccCode.Index).Value = ""
                        rowDGV.Cells(Me.colCostAccDesc.Index).Value = ""
                    End With
                    Me.Hitung_Amount(rowDGV)
                Next

                'tambahkan exchange rate data
                Me.SetCalculateBy()
                'penomoran
                Me.ObjCD.SetNumber(Me.dgvItem)
                Me.ObjCD.SetNumber(Me.dgvExchageRates)
                Me.FormatDGV()
                'hilangkan do dari item
                Me.TBL.Rows.Remove(rowDoNo)
                'text customer sudah tidak bisa diganti
                Me.txtCustomerCode.ReadOnly = True
                Me.btnFindCustomer.Enabled = False
                Me.txtCustomerCode.BackColor = Me.TxtReadOnly.BackColor

                Me.SetSalesPerson()

                '2025-07-16 dikomen
                'If Me.dgvItem.RowCount > 0 Then
                '    'Me.AddDP_Procedure()
                '    IsiDGVReturn()
                'End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Finally
                DTA1.Dispose()
                TBL1.Dispose()
                Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Windows.Forms.Cursor.Current = Cursors.Arrow
            End Try
        End If
    End Sub

    '\\TR==GlobalDiscount     '2025 Core Tax Tahap 2.1
    Private Sub HitungVATItemWithReturn(ByVal row As DataGridViewRow, ByVal exchangRate As Decimal, ByVal JmlReturn As Decimal)
        'Update SMESS 2026
        Dim rawAmount As Decimal = 0
        Dim rawUnitPrice As Decimal = 0
        Dim rawNettPrice As Decimal = 0
        Try

            If row.Cells(Me.ColPriceCond.Index).Value = "Include" Then
                row.Cells(Me.ColAmount2.Index).Value = row.Cells(Me.ColAmountInclude.Index).Value
                If row.Cells(Me.ColIsVAT.Index).Value = Boolean.TrueString Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmount2.Index).Value) / (100 + CDec(Me.txtTax.Text))) * 100, DefaultPrecision)
                    rawAmount = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmount2.Index).Value) / (100 + CDec(Me.txtTax.Text))) * 100, 2) 'Update SMESS 2026
                Else
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmount2.Index).Value) / (100)) * 100, DefaultPrecision)
                    rawAmount = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmount2.Index).Value) / (100)) * 100, DefaultPrecision + 2) 'Update SMESS 2026
                End If

                row.Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAmount2.Index).Value) - CDec(row.Cells(Me.ColAmount.Index).Value), DefaultPrecision)
                If row.Cells(Me.ColPrimaryUnit.Index).Value <> row.Cells(Me.ColPrimaryUnitsMulti.Index).Value And row.Cells(Me.ColPriceByMultiUnit.Index).Value = "0" Then
                    row.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColVAT2.Index).Value) / CDec(row.Cells(Me.ColTotalQty.Index).Value), DefaultPrecision)
                    row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAmount.Index).Value) / CDec(row.Cells(Me.ColTotalQty.Index).Value), DefaultPrecision)
                    rawNettPrice = Me.ObjCD.DigitGrouping(CDec(rawAmount) / CDec(row.Cells(Me.ColTotalQty.Index).Value), DefaultPrecision + 2) 'Update SMESS 2026
                Else
                    row.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColVAT2.Index).Value) / CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecision)
                    row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAmount.Index).Value) / CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecision)
                    rawNettPrice = Me.ObjCD.DigitGrouping(CDec(rawAmount) / CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecision + 2)  'Update SMESS 2026
                End If
                If row.Cells(Me.ColIsVAT.Index).Value = Boolean.TrueString Then
                    row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) / (100 + CDec(Me.txtTax.Text))) * 100, DefaultPrecision)
                    rawUnitPrice = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) / (100 + CDec(Me.txtTax.Text))) * 100, DefaultPrecision + 2) 'Update SMESS 2026
                Else
                    row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) / (100)) * 100, DefaultPrecision)
                    rawUnitPrice = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColUnitPriceInclude.Index).Value) / (100 + CDec(Me.txtTax.Text))) * 100, DefaultPrecision + 2) 'Update SMESS 2026
                End If
                If CDec(row.Cells(Me.ColDiscAmountInclude.Index).Value) = 0 Then
                    If CDec(row.Cells(Me.ColUnitPrice.Index).Value) < CDec(row.Cells(Me.ColNettPrice.Index).Value) Then
                        row.Cells(Me.ColUnitPrice.Index).Value = row.Cells(Me.ColNettPrice.Index).Value
                    End If
                End If
                row.Cells(Me.ColDiscountAmount.Index).Value = CDec(row.Cells(Me.ColUnitPrice.Index).Value) - CDec(row.Cells(Me.ColNettPrice.Index).Value)
                'Update SMESS 2026
                'row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColDiscountAmount.Index).Value) / CDec(row.Cells(Me.ColUnitPrice.Index).Value)) * 100, 2)
                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(((rawUnitPrice - rawNettPrice) / rawUnitPrice) * 100, 2)

                If JmlReturn = 0 Then
                    row.Cells(Me.ColAmountIncludeReturn.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAmountInclude.Index).Value) * exchangRate, DefaultPrecision)
                Else
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmountIncludeReturn.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                        row.Cells(Me.ColNettPriceInclude.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    Else
                        row.Cells(Me.ColAmountIncludeReturn.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                        row.Cells(Me.ColNettPriceInclude.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    End If
                End If
                row.Cells(Me.ColAmount2Return.Index).Value = row.Cells(Me.ColAmountIncludeReturn.Index).Value
                If row.Cells(Me.ColIsVAT.Index).Value = Boolean.TrueString Then
                    row.Cells(Me.ColAmountReturn.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmount2Return.Index).Value) / (100 + CDec(Me.txtTax.Text))) * 100, DefaultPrecision)
                Else
                    row.Cells(Me.ColAmountReturn.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmount2Return.Index).Value) / 100) * 100, DefaultPrecision)
                End If

                row.Cells(Me.ColVAT2Return.Index).Value = CDec(row.Cells(Me.ColAmount2Return.Index).Value) - CDec(row.Cells(Me.ColAmountReturn.Index).Value)
            Else
                row.Cells(Me.ColUnitPrice.Index).Value = row.Cells(Me.ColUnitPriceInclude.Index).Value
                row.Cells(Me.ColDiscount.Index).Value = row.Cells(Me.ColDiscountInclude.Index).Value
                row.Cells(Me.ColDiscountAmount.Index).Value = row.Cells(Me.ColDiscAmountInclude.Index).Value
                row.Cells(Me.ColNettPrice.Index).Value = row.Cells(Me.ColNettPriceInclude.Index).Value
                row.Cells(Me.ColAmount.Index).Value = row.Cells(Me.ColAmountInclude.Index).Value


                If row.Cells(Me.ColIsVAT.Index).Value = Boolean.TrueString Then
                    row.Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmount.Index).Value) * CDec(Me.txtTax.Text)) / 100, DefaultPrecision)
                Else
                    row.Cells(Me.ColVAT2.Index).Value = Me.ObjCD.DigitGrouping(0, DefaultPrecision)
                End If
                If row.Cells(Me.ColPrimaryUnit.Index).Value <> row.Cells(Me.ColPrimaryUnitsMulti.Index).Value And row.Cells(Me.ColPriceByMultiUnit.Index).Value = "0" Then
                    row.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColVAT2.Index).Value) / CDec(row.Cells(Me.ColTotalQty.Index).Value), DefaultPrecision)
                Else
                    row.Cells(Me.ColVAT.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColVAT2.Index).Value) / CDec(row.Cells(Me.ColTotalQtyMulti.Index).Value), DefaultPrecision)
                End If
                row.Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColVAT2.Index).Value) + CDec(row.Cells(Me.ColAmount.Index).Value), DefaultPrecision)
                If JmlReturn = 0 Then
                    row.Cells(Me.ColAmountIncludeReturn.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColAmountInclude.Index).Value) * exchangRate, DefaultPrecision)
                Else
                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                        row.Cells(Me.ColAmountIncludeReturn.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQtyMulti.Index).Value - JmlReturn) * _
                        row.Cells(Me.ColNettPriceInclude.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    Else
                        row.Cells(Me.ColAmountIncludeReturn.Index).Value = Me.ObjCD.DigitGrouping((row.Cells(Me.ColTotalQty.Index).Value - JmlReturn) * _
                        row.Cells(Me.ColNettPriceInclude.Index).Value * exchangRate, Me.DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                    End If
                End If

                row.Cells(Me.ColAmountReturn.Index).Value = row.Cells(Me.ColAmountIncludeReturn.Index).Value
                If row.Cells(Me.ColIsVAT.Index).Value = Boolean.TrueString Then
                    row.Cells(Me.ColVAT2Return.Index).Value = Me.ObjCD.DigitGrouping((CDec(row.Cells(Me.ColAmountReturn.Index).Value) * CDec(Me.txtTax.Text)) / 100, DefaultPrecision)
                Else
                    row.Cells(Me.ColVAT2Return.Index).Value = Me.ObjCD.DigitGrouping(0, DefaultPrecision)
                End If

                row.Cells(Me.ColAmount2.Index).Value = Me.ObjCD.DigitGrouping(CDec(row.Cells(Me.ColVAT2.Index).Value) + CDec(row.Cells(Me.ColAmount.Index).Value), DefaultPrecision)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '\\TR==GlobalDiscount

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

            If Me.txtPaymentTerm.Text <> Me.ObjCD.DigitGrouping(Minimum) Then
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

    Private Sub chkShowCustCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowCustCode.CheckedChanged
        Try
            If Me.dgvItem.Columns.Count > 0 Then
                If Me.chkShowCustCode.Checked = True Then
                    Me.dgvItem.Columns(3).Visible = True
                    Me.dgvItem.Columns(4).Visible = True
                Else
                    Me.dgvItem.Columns(3).Visible = False
                    Me.dgvItem.Columns(4).Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Me.buffer = e.Row.Cells(Me.ColPrimaryDO.Index).Value
    End Sub

    Private Sub SetCalculateBy()
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
        Me.ObjCD.Number_KeyUp(Me.txtDutyStamp, sender, e)
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
                    If Me.TxtRemarks.Enabled = True Then
                        Me.Hitung_Total()
                    End If
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
        Try
            Me.AddDP_Procedure()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddDP_Procedure()
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

            'Me.ColPrimaryInvDP.Visible = True
            Me.ObjInvoice.IsiDgvDP(ij, Me.dgvDP, Me.txtPrimaryCurrency.Text, Me.txtPrimaryInvoice.Text)
            Me.ObjCD.SetNumber(Me.dgvDP)
            Me.FormatDGV()
            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
            End If
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

    Private Sub IsiDGVReturn(Optional ByVal vHitungTotal As Boolean = True)
        Dim ListPrimary As New List(Of Double)
        Dim row As DataGridViewRow
        Dim rowSR As dseTransaksiSalesReturn.DataTable3Row

        Try
            'kosongkan baris yang diberikan
            Me.dgvItemSR.Rows.Clear()
            Me.dgvSR.Rows.Clear()

            'cari data SR, berdasarkan data do yang dituliskan
            For Each row In Me.dgvItem.Rows
                If ListPrimary.Contains(row.Cells(Me.ColPrimaryDO.Index).Value) = False Then
                    ListPrimary.Add(row.Cells(Me.ColPrimaryDO.Index).Value)
                End If
            Next
            Me.ObjSR.IsiTBLSRByPrimaryDO(ListPrimary, TBLSR)
            For Each rowSR In Me.TBLSR.Rows
                Me.dgvItemSR.Rows.Add("", rowSR.Sales_Return_No, rowSR.Delivery_Order_Number, "", rowSR.Primary)
            Next
            Me.IsiDGVReturnItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub HapusDP()
        Try

            Dim ij As New List(Of Double)
            Dim row As DataGridViewRow

            'tambahkan item DP
            Me.dgvDP.Rows.Clear()
            For Each row In Me.dgvItem.Rows
                If ij.Contains(Me.dgvItem.Item(Me.ColPrimarySO.Index, row.Index).Value) = False Then
                    ij.Add(Me.dgvItem.Item(Me.ColPrimarySO.Index, row.Index).Value)
                End If
            Next

            'Me.ColPrimaryInvDP.Visible = True
            Me.ObjInvoice.IsiDgvDP(ij, Me.dgvDP, Me.txtPrimaryCurrency.Text, Me.txtPrimaryInvoice.Text)
            Me.ObjCD.SetNumber(Me.dgvDP)
            Me.FormatDGV()

            If Me.TxtRemarks.Enabled = True Then
                Me.Hitung_Total()
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

    Private Sub IsiDGVReturnItem(Optional ByVal vHitungTotal As Boolean = True)
        Dim ListPrimary As New List(Of Double)
        Dim row As DataGridViewRow

        Try
            'untuk setiap data SR yang ada, masukkan ke dalam dgv item
            For Each row In Me.dgvItemSR.Rows
                If ListPrimary.Contains(row.Cells(4).Value) = False Then
                    ListPrimary.Add(row.Cells(4).Value)
                End If
            Next
            Me.ObjSRItem.IsiDGVByPrimarySR(ListPrimary, Me.dgvSR, Me.ObjCD)
            Me.FormatDGVSR()

            'hitung ulang sub total dengan memperhitungkan SR
            'For Each row In Me.dgvItem.Rows
            'Me.Hitung_Amount(row)
            'Next
            If vHitungTotal = True Then
                Me.Hitung_Total()
            End If
            Me.ObjCD.SetNumber(Me.dgvItemSR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimary = Nothing
        End Try
    End Sub

    Private Sub llbAddSR_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddSR.LinkClicked
        Try
            If Me.dgvItem.RowCount > 0 Then
                Me.IsiDGVReturn()
            Else
                MsgBox("Please insert invoice item first", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGVSR()
        Dim row As DataGridViewRow

        Try
            Dim AdaSec As Boolean = False
            Dim AdaExpiry As Boolean = False
            Dim AdaBatchNo As Boolean = False
            Dim IsMultiUnit As Boolean = False

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

            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If .Cells(Me.ColPrimaryUnit.Index).Value <> .Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                        IsMultiUnit = True
                        Exit For
                    End If
                End With
            Next
            If IsMultiUnit = False Then
                Me.COlTotalQtySR.Visible = False
                Me.SR_ColUnits.Visible = False
                Me.ColRatioDesc_SR.Visible = False
            Else
                Me.COlTotalQtySR.Visible = True
                Me.SR_ColUnits.Visible = True
                Me.ColRatioDesc_SR.Visible = True
            End If

            Me.ObjCD.SetNumber(Me.dgvSR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSR_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemSR.UserDeletedRow
        Try
            Me.IsiDGVReturnItem()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnRevisi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRevisi.Click
        Try
            If Me.txtInvoiceNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtInvoiceNo.Text <> "" Then
                If Me.ObjInvoice.ApakahBolehRevisi(Me.ObjInvoice.Primary, Me.UserCode, True) = True Then
                    If Me.ObjInvoice.ApakahQuickInvoice(Me.ObjInvoice.Primary) = True Then
                        MsgBox("This invoice is not sales invoice. Use quick invoice form", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    'set parameter untuk revisi
                    Me.baru = True
                    Me.SedangRevisi = True
                    Me.ObjInvoice.CariDetail(Me.ObjInvoice.Primary, Me.dgvItem, Me.dgvDP, _
                    Me.dgvItemSR, Me.dgvExchageRates, Me)
                    Me.SalinKeField()
                    Me.txtRevisiNo.Text = CDec(Me.txtRevisiNo.Text) + 1

                    If CDec(Me.txtRevisiNo.Text) = 1 Then
                        Me.txtInvoiceNo.Text = Me.txtInvoiceNo.Text + "A"
                    Else
                        Dim Abjad As String = ""
                        Select Case CDec(Me.txtRevisiNo.Text)
                            Case 2
                                Abjad = "B"
                                Exit Select
                            Case 3
                                Abjad = "C"
                                Exit Select
                            Case 4
                                Abjad = "D"
                                Exit Select
                            Case 5
                                Abjad = "E"
                                Exit Select
                            Case 6
                                Abjad = "F"
                                Exit Select
                            Case 7
                                Abjad = "G"
                                Exit Select
                            Case 8
                                Abjad = "H"
                                Exit Select
                            Case 9
                                Abjad = "I"
                                Exit Select
                            Case 10
                                Abjad = "J"
                                Exit Select
                            Case 11
                                Abjad = "K"
                                Exit Select
                            Case Else
                                MsgBox("This software only allows 11 revisions", MsgBoxStyle.Information, "Information")
                                Me.ObjInvoice.CariDetail(Me.ObjInvoice.Primary, Me.dgvItem, Me.dgvDP, _
                                Me.dgvItemSR, Me.dgvExchageRates, Me)
                                Me.SedangRevisi = False
                                Me.NonAktifkanDetail()
                                Exit Sub
                        End Select
                        Me.txtInvoiceNo.Text = Me.txtInvoiceNo.Text.Substring(0, Me.txtInvoiceNo.Text.Length - 1) + Abjad
                    End If
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    Me.txtCustomerCode.Enabled = False
                    Me.btnFindCustomer.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
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

    Private Sub txtDutyStamp_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDutyStamp.Leave
        Try
            If Me.TxtRemarks.Enabled = True Then
                ' Me.txtDutyStamp.Text = Me.ObjCD.DigitGrouping(Me.txtDutyStamp.Text, Me.ObjCurrency.GetDecimal_DefaultCurrency)
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDP_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDP.CellContentClick

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
                Me.txtNoFakturPajak.TabStop = False
            Else
                'faktur pajak standar
                Me.txtNoFakturPajakSederhana.Visible = False
                Me.txtNoFakturPajakSederhana.TabStop = False
                Me.txtNoFakturPajakSederhana.SendToBack()
                Me.txtNoSeriFakturPajak.TabStop = True
                Me.txtNoFakturPajak.TabStop = True

                '2025 Core Tax Tahap 2.2
                'If Me.baru = False And Me.txtNoFakturPajak.Text = "" Then 
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

    Private Sub txtTaxExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTaxExchangeRate.Enter
        Try
            If Me.txtTaxExchangeRate.Enabled = True Then
                Me.ObjCD.Number_Enter(Me.txtTaxExchangeRate, sender, e)
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
            Me.ShowDetailFakturPajak()
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
            If Me.TxtRemarks.Enabled = True OrElse IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso CDec(Me.txtPrimaryInvoice.Text) > 0 Then


                If Me.dtpDate.Value < Obj.BatasStartDate Then
                    i.mtb_Number.Mask = "00000000"
                    i.mtb_Number.Text = Me.txtNoFakturPajak.Text.Substring(3)
                    i.MinLength = 8
                    i.Text = "Edit Tax Invoice Number"

                    If Me.TxtRemarks.Enabled = False Then
                        i.TxtRevision.Enabled = False
                        i.mtb_Number.Enabled = False
                    End If
                    If i.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Dummy = Me.txtNoFakturPajak.Text.Substring(0, 3) + i.mtb_Number.Text
                        PrimaryInvoice = Me.ObjInvoice.GetPrimaryInvoiceByNoFakturPajak(Dummy)
                        If PrimaryInvoice = -1 Then
                            Me.txtNoFakturPajak.Text = Dummy
                        Else
                            If CDec(Me.txtPrimaryInvoice.Text) = PrimaryInvoice Then
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

                    'copy data
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


                    If Me.TxtRemarks.Enabled = False Then
                        'copy data
                        j.mtb_Number.Enabled = False
                        j.cboJenisFaktur.Enabled = False
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
    '        If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso CDec(Me.txtPrimaryInvoice.Text)) Then
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
    '                i.TxtRevision.Enabled = False
    '                i.mtb_Number.Enabled = False
    '            End If

    '            If i.ShowDialog = Windows.Forms.DialogResult.OK Then
    '                Dummy = Me.txtInvoiceNo.Text.Replace(Dummy1, i.mtb_Number.Text)
    '                Dummy = Dummy.Substring(0, JmlRujukan2) + i.TxtRevision.Text

    '                PrimaryInvoice = Me.ObjInvoice.GetPrimaryInvoice(Dummy)
    '                If PrimaryInvoice = -1 Then
    '                    Me.txtInvoiceNo.Text = Dummy
    '                Else
    '                    If CDec(Me.txtPrimaryInvoice.Text) = PrimaryInvoice Then
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
            If Me.TxtRemarks.Enabled = True OrElse (IsNumeric(Me.txtPrimaryInvoice.Text) AndAlso CDec(Me.txtPrimaryInvoice.Text)) Then
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
                    i.TxtRevision.Enabled = False
                    i.mtb_Number.Enabled = False
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
                        If CDec(Me.txtPrimaryInvoice.Text) = PrimaryInvoice Then
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
                Dim IntDeffPrec As Integer = 0
                IntDeffPrec = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrTax.Text)
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text, IntDeffPrec)
                If Me.TxtRemarks.Enabled = True Then
                    Me.Hitung_Total()
                End If

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
            'decimal 2, sesuai KMK
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
            '//TR : Tambahan baru Nik/Pas
            If Me.txtPrimaryCustomer.Text <> "" And Me.txtPrimaryCustomer.Text <> "-1" Then
                If Me.cboNIKPASS.Visible = True Then
                    Me.txtNoNIKPass.Text = Me.ObjCustomer.GetNoNikPass2(Me.txtPrimaryCustomer.Text, Me.OptionNIKPASS)
                Else
                    If MsgBox("Update NPWP number from customer data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
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

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.setDtpDateInvoice()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub setDtpDateInvoice()
        Try
            If Me.dtpDate.Value >= Me.ObjTaxInvoiceSetting.BatasStartDate Then
                Label8.Visible = False
            Else
                Label8.Visible = True
            End If

            If Me.TxtRemarks.Enabled = True Then
                If Me.SedangRevisi = False Then
                    If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                        If Me.baru = False Then
                            If Me.ObjInvoice.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjInvoice.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
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

                '2025 Core Tax Tahap 2.2, di comment
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
    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        ObjCD.UpdateThreading()
        Threading.Thread.Sleep(100)
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Dim Middle As Double = 0

        Try
            'jika bukan auto number, maka btnrevisi tidak tampil
            '2019-09-25: iwan
            'fungsi revisi tidak ada lagi, langsung geser2 aja
            'If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = False Then
            Me.btnExit.Left = Me.btnPrint.Left
            Me.btnPrint.Left = Me.btnDelete.Left
            Me.btnDelete.Left = Me.BtnFaktur.Left
            Me.BtnFaktur.Left = Me.btnRevisi.Left
            Me.btnRevisi.Visible = False
            Me.btnRevisi.TabStop = False

            'tengahin posisi btn faktur pajak
            Middle = Me.btnOK.Left + Me.btnOK.Width + 6 + Me.GroupBox1.Left
            Middle = Math.Floor(Me.Width / 2)
            Me.btnDelete.Left = Middle - Math.Floor(Me.btnDelete.Width / 2)
            Me.btnPrint.Left = Me.btnDelete.Left + Me.btnDelete.Width + 6
            Me.btnExit.Left = Me.btnPrint.Left + Me.btnPrint.Width + 6
            Me.btnEdit.Left = Me.btnDelete.Left - Me.btnEdit.Width - 6
            Me.btnNew.Left = Me.btnEdit.Left - Me.btnNew.Width - 6
            Me.btnFind.Left = Me.btnNew.Left - Me.btnFind.Width - 6
            'End If

            Me.btnExit.Left = Me.btnPrint.Left + Me.btnPrint.Width + 6
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoSeriFakturPajak_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoSeriFakturPajak.Enabled = True Then
                    If Me.txtNoSeriFakturPajak.ReadOnly = True Then
                        Me.txtNoSeriFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtNoSeriFakturPajak.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoSeriFakturPajak_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoSeriFakturPajak.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.txtNoSeriFakturPajak.Enabled = True Then
                    If Me.txtNoSeriFakturPajak.ReadOnly = True Then
                        Me.txtNoSeriFakturPajak.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtNoSeriFakturPajak.BackColor = Me.txtPrimaryCurrency.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.GotFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExchangeRate.ReadOnly = True Then
                    Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtExchangeRate.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.LostFocus
        Try
            If Me.TxtRemarks.Enabled = True Then
                If Me.TxtExchangeRate.ReadOnly = True Then
                    Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtExchangeRate.BackColor = Me.txtPrimaryInvoice.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtExchangeRate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtExchangeRate, sender, e)
    End Sub

    Private Sub TxtExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtExchangeRate, sender, e)
    End Sub

    Private Sub TxtExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Leave
        Try
            'decimal 2, sesuai KMK
            Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Me.TxtExchangeRate.Text, Core_Module.DefPrecExchangeRate)

            If Me.TxtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotalBeforeDP.Text) Then
                Me.Hitung_Total()
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
                Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage4.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.txtPrimaryInvoice.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.txtPrimaryInvoice.BackColor, False)
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

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.TxtRemarks.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.TxtRemarks.Enabled = False Then
                If Me.ObjInvoice.CariDetail(vPrimary, _
               Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me) _
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

                I.PrimaryCurrency = Me.txtPrimaryCurrency.Text
                I.PrimaryCustomer = Me.txtPrimaryCustomer.Text
                I.TBLHeaderCost = Me.TBLHeaderCost
                I.rowHeaderCost = Me.rowHeaderCost
                I.TBLitemCost = Me.TBLitemCost
                I.rowitemCost = Me.rowitemCost
                I.OtherCharge = Me.txtOtherCharge.Text
                I.OtherChargeDummy = Me.txtOtherCharge.Text

                If Me.TxtRemarks.Enabled = True Then
                    I.Aktif = True
                Else
                    I.Aktif = False
                End If

                I.PrimaryCustomerShipTo = Me.PrimaryCustomerShipTo
                I.IdentifierAddressCustomerShipTo = Me.IdentifierAddressCustomerShipTo
                I.ShowDialog()

                If Me.TxtRemarks.Enabled = True Then
                    If IsNumeric(I.OtherCharge) Then
                        Me.txtOtherCharge.Text = Me.ObjCD.DigitGrouping(I.OtherCharge, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                    Else
                        Me.txtOtherCharge.Text = 0
                    End If
                    Me.Hitung_Total()
                End If
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

    Public Sub FromCutOff_SalesInvoice_Create(ByVal vPrimaryTransaksi As Double, ByVal ObjCutOff_Core As Core_Module.ClsTrans_CutOffDB_Core, _
    ByVal vTblPrimaryDO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimarySO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimaryInvDP As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable, _
    ByVal vTblPrimarySR As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable)
        Dim TblTempDO, TblTempSO As New Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable
        Dim PrimaryUpdate As Double = -1

        Try
            Me.baru = False
            Me.ObjInvoice.ResetObject()
            Me.ObjInvoice.TipeInvoice = ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice

            Me.ObjInvoice.CariDetail(vPrimaryTransaksi, Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me)
            Me.UserCode = Me.ObjInvoice.CreatedBy

            For Each rowD As DataGridViewRow In Me.dgvItem.Rows
                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimaryDO.Index).Value) = False And rowD.Cells(Me.ColPrimaryDO.Index).Value <> "-1" Then
                    TblTempDO.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimaryDO.Index).Value)
                    If TblTempDO.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryDO.Index).Value), vTblPrimaryDO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryDO.Index).Value), PrimaryUpdate, TblTempDO)
                        rowD.Cells(Me.ColPrimaryDO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimaryDO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimaryDO.Index).Value), TblTempDO)
                    End If
                End If

                If String.IsNullOrEmpty(rowD.Cells(Me.ColPrimarySO.Index).Value) = False And rowD.Cells(Me.ColPrimarySO.Index).Value <> "-1" Then
                    TblTempSO.DefaultView.RowFilter = "[Primary Old] = " & CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value)
                    If TblTempSO.DefaultView.Count = 0 Then
                        PrimaryUpdate = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), vTblPrimarySO)
                        ObjCutOff_Core.AddPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), PrimaryUpdate, TblTempSO)
                        rowD.Cells(Me.ColPrimarySO.Index).Value = PrimaryUpdate
                    Else
                        rowD.Cells(Me.ColPrimarySO.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD.Cells(Me.ColPrimarySO.Index).Value), TblTempSO)
                    End If
                End If
            Next

            If Me.dgvDP.RowCount > 0 Then
                For Each rowD_DP As DataGridViewRow In Me.dgvDP.Rows
                    'Update Primary Sales Invoice DP
                    If String.IsNullOrEmpty(rowD_DP.Cells(Me.ColPrimaryInvDP.Index).Value) = False And rowD_DP.Cells(Me.ColPrimaryInvDP.Index).Value <> "-1" Then
                        rowD_DP.Cells(Me.ColPrimaryInvDP.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD_DP.Cells(Me.ColPrimaryInvDP.Index).Value), vTblPrimaryInvDP)
                    End If
                Next
            End If

            If Me.dgvItemSR.RowCount > 0 Then
                For Each rowD_PR As DataGridViewRow In Me.dgvItemSR.Rows
                    'Update Primary Sales Return
                    If String.IsNullOrEmpty(rowD_PR.Cells(Me.ColPrimarySR.Index).Value) = False And rowD_PR.Cells(Me.ColPrimarySR.Index).Value <> "-1" Then
                        rowD_PR.Cells(Me.ColPrimarySR.Index).Value = ObjCutOff_Core.GetPrimaryUpdate(CDbl(rowD_PR.Cells(Me.ColPrimarySR.Index).Value), vTblPrimarySR)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_SalesInvoice_Save() As Double
        Try
            Me.baru = True
            Me.ObjInvoice.SimpanBaru1(Me.dgvItem, Me.dgvDP, Me.dgvItemSR, Me.dgvExchageRates, Me.UserCode, Me, False, True)
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
                Me.OptionNIKPASS = clsMasterCustomer.EnumOptionNikPass.ByNIK
            ElseIf Me.cboNIKPASS.SelectedIndex = 1 Then
                Me.OptionNIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByPASSPORT
            Else '2025 Core Tax Tahap 2.1
                Me.OptionNIKPASS = Core_Module.clsMasterCustomer.EnumOptionNikPass.ByOtherID
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    'for quick input
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.AddItemByQuickInput.Enabled = ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSalesInvoice_QuickByStockOut")
            Me.ContextMenuStrip1.Show(LinkLabel1, New Point(0, 15))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'for quick input
    Private Sub AddItemByQuickInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddItemByQuickInput.Click
        Dim i As New FrmTransaksiSalesInvoice_QuickByStockOut
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Try
            i.UserCode = Me.UserCode
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    'for quick input
    Private Sub BtnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnStop.Click
        Try
            If MsgBox("Stop process?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.StopQuickInput = True
                Me.NonAktifkanDetail()
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'for quick input
    Public Function CreateInvoiceByQuickInputStockOut(ByVal tblCreateToInv As DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceDataTable, _
    ByVal vDirectlySave As Boolean, ByVal vDirectlyPrint As Boolean, ByVal vPrintCustomPath As String) As Boolean
        Dim rowCreateToInv As DseTransaksiSalesInvoice_QuickByStockOut.tbl_create_invoiceRow
        Dim rowDoNo As DseTransaksiDO.tbltrans_do_view2Row
        Dim Dataview As DataRowView
        Dim SysObj As System.Object
        Dim SysLlb As System.Windows.Forms.LinkLabelLinkClickedEventArgs

        Try
            If tblCreateToInv.Rows.Count > 0 Then
                Me.btnNew_Click(SysObj, SysLlb)

                Me.dtpDate.Value = Me.SetDtpDate
                Me.setDtpDateInvoice()

                'get customer
                rowCreateToInv = tblCreateToInv.Rows(0)
                Me.txtCustomerCode.Focus()
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(rowCreateToInv.primary_cust)
                Me.txtCustomerCode_Leave(SysObj, SysLlb)
                Me.TxtCustomerName.Focus()
                'get currency
                Me.txtCurrencyCode.Focus()
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(rowCreateToInv.primary_curr)
                Me.txtCurrencyCode_Leave(SysObj, SysLlb)
                Me.txtSubTotalBeforeDP.Focus()

                'get tax
                Me.chkTax.Checked = rowCreateToInv.use_tax
                If Me.chkTax.Checked = True Then
                    Me.txtTax.Focus()
                    Me.txtTax.Text = rowCreateToInv.tax_perc
                    Me.TxtTotal.Focus()
                    Me.txtNoSeriFakturPajak.Text = "010"
                End If

                Dim UrutanDO As Integer = 0
                While Me.cboDO.Items.Count > UrutanDO
                    Me.cboDO.SelectedIndex = UrutanDO
                    Dataview = Me.cboDO.SelectedItem
                    rowDoNo = Dataview.Row
                    tblCreateToInv.DefaultView.RowFilter = "[primary do] = " & rowDoNo.Primary_DO
                    If tblCreateToInv.DefaultView.Count > 0 Then
                        Me.BtnInput_Click(SysObj, SysLlb)
                        'Me.BtnInput.PerformClick()
                    Else
                        UrutanDO += 1
                    End If
                End While

                If Me.dgvItem.Rows.Count > 0 Then
                    If vDirectlySave = True Then
                        Me.btnOK.PerformClick()

                        If vDirectlyPrint = True Then
                            If Me.txtInvoiceNo.Text <> "" Then
                                Dim ij As New FrmCetakInvoice
                                If Me.ObjInvoice.CetakInvoice(ij.crv1, Me.txtPrimaryInvoice.Text, Me.dtpDate.Value, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text), Me.UserCode) = True Then
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
                                    End If
                                End If
                                ij.Dispose()
                            End If
                        End If
                    End If
                End If
            End If

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
                    If Me.ObjInvoice.SetDontCollect1(Me.cboTaxtype.SelectedIndex, Me.chkTax.Checked, Me.cboPPNDipungut.SelectedIndex) = True Then
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
