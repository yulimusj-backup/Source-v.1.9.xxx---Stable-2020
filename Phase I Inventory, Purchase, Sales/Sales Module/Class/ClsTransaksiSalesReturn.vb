Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesReturn
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mNoUrut As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mRemarks As String = ""
    Private mInvoiceCancellation As Integer = 0

    Private mPrimaryStockInSalesReturn As Double = -1
    Private mPrimaryDO As Double = -1
    Private mCetakNotaRetur As Integer = 0
    Private mReasonOfReturn As String = ""
    Private mNoNotaRetur As String = ""
    Private mTipeSR As EnumTipeSR = EnumTipeSR.Ganti_Barang
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim mPrimaryCurrency As Double = -1
    Dim mSubTotal As Decimal = 0
    Dim mTaxPercent As Decimal = 0
    Dim mUseTax As Boolean = False
    Dim mTaxExchangeRate As Decimal = 1
    Dim mTaxAmount As Decimal = 0
    Dim mDutystamp As Decimal = 0
    Dim mTotal As Decimal = 0
    Dim mPaymentTerm As Double = 0
    Dim mQuickSR As Integer = 0
    Dim mInvoiceNoQuickSR As String = ""
    Dim mTaxInvoiceNoQuickSR As String = ""
    Dim mInvoiceDateQuickSR As Date = Today
    Dim mPrimaryCustomer As Double = -1
    Dim mPrimarySalesPerson As Double = -1

    'update NPWP 16 DIgit
    Dim mNPWPNo As String = ""

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Dim mPrimaryCurrTax As Double = -1
    Dim mPrimaryAccount As Double = -1

    Public Property PrimaryCurrTax() As Double
        Get
            Return Me.mPrimaryCurrTax
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrTax = value
        End Set
    End Property

    Public Property PrimaryAccount() As Double
        Get
            Return Me.mPrimaryAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryAccount = value
        End Set
    End Property

    'update NPWP 16 DIgit
    Public Property NPWPNo() As String
        Get
            Return Me.mNPWPNo
        End Get
        Set(ByVal value As String)
            mNPWPNo = value
        End Set
    End Property

    Public Property InvoiceCancellation() As Integer
        Get
            Return mInvoiceCancellation
        End Get
        Set(ByVal value As Integer)
            mInvoiceCancellation = value
        End Set
    End Property

    Public Overridable Function CariDetail(ByVal vTbl As dseTransaksiSalesReturn.tbltrans_salesreturn_NewDataTable, _
            ByVal vTblItem As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_item_copyDataTable, _
            ByVal vPrimaryTransaksi As Double, _
            ByVal vDgvItem As DataGridView, ByVal ij As FrmTransaksiSalesReturn_Quick, _
            ByVal vTblCustom As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1DataTable) As Boolean
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_NewRow
        Dim ObjItem As New ClsTransaksiSalesReturnItem
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesReturn_Quick_AddByCustomProduct
        Try
            'fil basis data
            If vTbl.Count > 0 Then
                'salin data header
                row = vTbl.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    mPrimaryStockInSalesReturn = .Primary_Stock_In_Sales_Return
                    mPrimaryDO = .Primary_DO
                    mCetakNotaRetur = .Use_Nota_Retur_Pajak
                    mReasonOfReturn = .Reason_Of_Return
                    Me.mNoNotaRetur = .No_Nota_Retur_Pajak
                    Me.mTipeSR = .Tipe_SR

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mSubTotal = .Sub_Total
                    Me.mDutystamp = .Duty_stamp
                    Me.mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxPercent = .Tax_Percent
                    Me.mTotal = .Total
                    Me.mPaymentTerm = .Payment_Term
                    If .Use_Tax = 1 Then
                        Me.mUseTax = True
                    Else
                        Me.mUseTax = False
                    End If
                    'update NPWP 16 Digit
                    Me.mNPWPNo = .No_NPWP

                    Me.mTaxInvoiceNoQuickSR = .Tax_Invoice_No_Quick
                    Me.mInvoiceDateQuickSR = .Invoice_Date_Quick
                    Me.mInvoiceNoQuickSR = .Invoice_No_Quick
                    Me.mPrimaryCustomer = .Primary_Customer_Quick
                    Me.mPrimarySalesPerson = .Primary_Sales_Person
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccount = .primary_ar_account
                End With

                'salin data item 
                ObjItem.IsiDGV(vTblItem, vDgvItem, Me.mPrimary, ij)

                ObjCustom.IsiTblDummyFromDB(Me.mPrimary, vTblCustom)
                'nilai balikan

                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView, ByVal ij As FrmTransaksiSalesReturn) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_NewDataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_NewRow
        Dim ObjItem As New ClsTransaksiSalesReturnItem

        Try
            'fil basis data
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    mPrimaryStockInSalesReturn = .Primary_Stock_In_Sales_Return
                    mPrimaryDO = .Primary_DO
                    mCetakNotaRetur = .Use_Nota_Retur_Pajak
                    mReasonOfReturn = .Reason_Of_Return
                    Me.mNoNotaRetur = .No_Nota_Retur_Pajak
                    Me.mTipeSR = .Tipe_SR
                    Me.mInvoiceCancellation = .Invoice_Cancellation

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mSubTotal = .Sub_Total
                    Me.mDutystamp = .Duty_stamp
                    Me.mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxPercent = .Tax_Percent
                    Me.mTotal = .Total
                    Me.mPaymentTerm = .Payment_Term
                    If .Use_Tax = 1 Then
                        Me.mUseTax = True
                    Else
                        Me.mUseTax = False
                    End If
                    'update NPWP 16 Digit
                    Me.mNPWPNo = .No_NPWP
                    Me.mPrimarySalesPerson = .Primary_Sales_Person
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccount = .primary_ar_account
                End With

                'salin data item 
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary, ij)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
   ByVal vDgvItem As DataGridView, ByVal ij As FrmTransaksiSalesReturn_Quick, _
   ByVal vTblCustom As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1DataTable) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_NewDataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_NewRow
        Dim ObjItem As New ClsTransaksiSalesReturnItem
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesReturn_Quick_AddByCustomProduct
        Try
            'fil basis data
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    mPrimaryStockInSalesReturn = .Primary_Stock_In_Sales_Return
                    mPrimaryDO = .Primary_DO
                    mCetakNotaRetur = .Use_Nota_Retur_Pajak
                    mReasonOfReturn = .Reason_Of_Return
                    Me.mNoNotaRetur = .No_Nota_Retur_Pajak
                    Me.mTipeSR = .Tipe_SR
                    Me.mInvoiceCancellation = .Invoice_Cancellation

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mSubTotal = .Sub_Total
                    Me.mDutystamp = .Duty_stamp
                    Me.mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxPercent = .Tax_Percent
                    Me.mTotal = .Total
                    Me.mPaymentTerm = .Payment_Term
                    If .Use_Tax = 1 Then
                        Me.mUseTax = True
                    Else
                        Me.mUseTax = False
                    End If

                    'update NPWP 16 Digit
                    Me.mNPWPNo = .No_NPWP

                    Me.mTaxInvoiceNoQuickSR = .Tax_Invoice_No_Quick
                    Me.mInvoiceDateQuickSR = .Invoice_Date_Quick
                    Me.mInvoiceNoQuickSR = .Invoice_No_Quick
                    Me.mPrimaryCustomer = .Primary_Customer_Quick
                    Me.mPrimarySalesPerson = .Primary_Sales_Person
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccount = .primary_ar_account
                End With

                'salin data item 
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary, ij)
                ObjCustom.IsiTblDummyFromDB(Me.mPrimary, vTblCustom)
                ObjCustom.IsiTblInSalesReturn(Me.mPrimary, ij.TblDummyCustom, ij.TblCustomInOutDummy) 'new
                'nilai balikan

                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesReturn) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSalesReturn_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode, vFrm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiSalesReturn, Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA2 As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim ObjItem As New ClsTransaksiSalesReturnItem
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim UseTax_Int As Integer
        Dim ObjDO As ClsTransaksiDO = vFrm.ObjDO
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryInvoice As Double = -1

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now
        Dim TblCustomInOutDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

        Try
            If Me.ApakahDataBenar(True, False, True, -1, vdgvItem, TblCustomInOutDummy, vCekDataBenar) = True Then
                If Me.mUseTax = True Then
                    UseTax_Int = 1
                Else
                    UseTax_Int = 0
                End If
                If Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak And Me.QuickSR = False Then
                    PrimaryInvoice = ObjDO.GetPrimaryInvoice(Me.mPrimaryDO)
                End If

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA2.InsertQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mPrimaryStockInSalesReturn, _
                Me.mPrimaryDO, Me.mTanggalTransaksi, Me.mRemarks, Me.mReasonOfReturn, _
                Me.mCetakNotaRetur, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mNoNotaRetur, Me.mTipeSR, Me.mPrimaryCurrency, Me.mSubTotal, _
                UseTax_Int, Me.mTaxPercent, Me.mTaxAmount, Me.mTaxExchangeRate, Me.mDutystamp, Me.mTotal, _
                Me.mPaymentTerm, PrimaryInvoice, Me.mQuickSR, "", #1/1/1950#, "", -1, -1, Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mInvoiceCancellation, Me.mNPWPNo) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 6))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, vFrm)
                    If vCekDataBenar = True Then
                        'cek berkaitan nomor sama untuk beda bulan
                        Me.CekNomor(DTA)
                    End If

                    'simpan data jurnal
                    If Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                        '**************************************************************************************************************************************
                        '**************************************************************************************************************************************
                        'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                        'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                        'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                        'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                        'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL
                        Dim ReturRupiah As Decimal = 0
                        Dim TotalRetur As Decimal = 0
                        Dim AmountTanpaSales As Decimal = 0
                        Dim DefPrecBase As Integer = 0
                        DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                        TotalRetur = Me.ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefPrecBase)
                        If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase)
                        End If
                        ReturRupiah = Me.ObjCD.DigitGrouping(TotalRetur - AmountTanpaSales, DefPrecBase)
                        '**************************************************************************************************************************************
                        Me.TransferJurnal_SR(Me.mPrimary, vFrm.ObjDO.GetPrimaryCust(Me.mPrimaryDO), _
                        ReturRupiah, _
                        Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), _
                        ObjInv.GetNoInvoice(ObjInv.GetPrimaryInvoice_ByPrimaryDO(Me.mPrimaryDO)), Me.mSubTotal, _
                        Me.mTaxExchangeRate, Me.mPrimaryCurrency, Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mTaxAmount, Me.mDutystamp)
                    End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vdgvItem, vUserCode, vFrm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            ObjItem = Nothing
            ObjInv = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesReturn_Quick, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1DataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSalesReturn_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode, vFrm, vTblCustom)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesReturn_Quick, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1DataTable, _
    Optional ByVal NoStockInSR As String = "", Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vKoreksiStock As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim DTA2 As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter
        Dim ObjItem As New ClsTransaksiSalesReturnItem
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim UseTax_Int As Integer
        Dim ObjDO As ClsTransaksiDO = vFrm.ObjDO
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryInvoice As Double = -1

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now
        Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard

        Try
            If Me.ApakahDataBenar(True, True, True, -1, vdgvItem, vFrm.TblCustomInOutDummy, vCekDataBenar) = True Then
                'simpan stock in SalesReturn 
                If Me.BuatStockInSalesReturn(Me.mPrimaryStockInSalesReturn, vFrm, vdgvItem, vUserCode, NoStockInSR, vCekDataBenar, vKoreksiStock, vFromCutOff) = False Then
                    MsgBox("Can't create stock in sales return, process cancelled", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If Me.mUseTax = True Then
                    UseTax_Int = 1
                Else
                    UseTax_Int = 0
                End If
                If Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak And Me.QuickSR = False Then
                    PrimaryInvoice = ObjDO.GetPrimaryInvoice(Me.mPrimaryDO)
                End If

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA2.InsertQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mPrimaryStockInSalesReturn, _
                Me.mPrimaryDO, Me.mTanggalTransaksi, Me.mRemarks, Me.mReasonOfReturn, _
                Me.mCetakNotaRetur, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mNoNotaRetur, Me.mTipeSR, Me.mPrimaryCurrency, Me.mSubTotal, _
                UseTax_Int, Me.mTaxPercent, Me.mTaxAmount, Me.mTaxExchangeRate, Me.mDutystamp, Me.mTotal, _
                Me.mPaymentTerm, PrimaryInvoice, Me.mQuickSR, Me.mInvoiceNoQuickSR, _
                Me.mInvoiceDateQuickSR, Me.mTaxInvoiceNoQuickSR, Me.mPrimaryCustomer, _
                Me.mPrimarySalesPerson, Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mInvoiceCancellation, Me.mNPWPNo) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 6))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    ObjSC.UpdateDocNo(Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_SalesReturn, Me.mPrimaryStockInSalesReturn, Me.mNoTransaksi)
                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, vFrm, vTblCustom)
                    If vFromCutOff = False Then
                        'cek berkaitan nomor sama untuk beda bulan
                        Me.CekNomor(DTA)
                    End If

                    'simpan data jurnal
                    If Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                        '**************************************************************************************************************************************
                        '**************************************************************************************************************************************
                        'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                        'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                        'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                        'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                        'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL
                        Dim ReturRupiah As Decimal = 0
                        Dim TotalRetur As Decimal = 0
                        Dim AmountTanpaSales As Decimal = 0
                        Dim DefPrecBase As Integer = 0
                        DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                        TotalRetur = Me.ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefPrecBase)
                        If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase)
                        End If
                        ReturRupiah = Me.ObjCD.DigitGrouping(TotalRetur - AmountTanpaSales, DefPrecBase)
                        '**************************************************************************************************************************************

                        Me.TransferJurnal_SR(Me.mPrimary, Me.mPrimaryCustomer, _
                        ReturRupiah, _
                        Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), _
                        ObjInv.GetNoInvoice(ObjInv.GetPrimaryInvoice_ByPrimaryDO(Me.mPrimaryDO)), Me.mSubTotal, _
                        Me.mTaxExchangeRate, Me.mPrimaryCurrency, Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mTaxAmount, Me.mDutystamp)
                    End If
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vdgvItem, vUserCode, vFrm, vTblCustom)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjInv = Nothing
            ObjCur = Nothing
            ObjSC = Nothing
        End Try
    End Function


    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesReturn) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSalesReturn_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(dgvItem, vUserCode, vFrm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesReturn) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim DTA2 As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter
        Dim ObjItem As New ClsTransaksiSalesReturnItem
        Dim UseTax_Int As Integer
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim PrimaryInv As Double = -1
        Dim ObjDO As New ClsTransaksiDO
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim TblCustomInOutDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
        Dim ObjSRCust As New CustomSMESS.ClsTransaksiSalesReturn

        Try
            If Me.ApakahDataBenar(False, False, True, Me.mPrimary, dgvItem, TblCustomInOutDummy) = True Then
                If Me.mUseTax = True Then
                    UseTax_Int = 1
                Else
                    UseTax_Int = 0
                End If
                If Me.QuickSR = False And Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                    PrimaryInv = ObjDO.GetPrimaryInvoice(Me.mPrimaryDO)
                End If

                If DTA2.UpdateQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mPrimaryStockInSalesReturn, _
                Me.mPrimaryDO, Me.mTanggalTransaksi, Me.mRemarks, Me.mReasonOfReturn, _
                Me.mCetakNotaRetur, vUserCode, Now, Now, Me.mNoNotaRetur, Me.mTipeSR, Me.mPrimaryCurrency, UseTax_Int, Me.mPaymentTerm, _
                PrimaryInv, Me.mQuickSR, Me.mSubTotal, Me.mTaxPercent, Me.mTaxAmount, Me.mTaxExchangeRate, Me.mDutystamp, Me.mTotal, _
                Me.mPrimarySalesPerson, Me.mInvoiceNoQuickSR, Me.mInvoiceDateQuickSR, Me.mTaxInvoiceNoQuickSR, _
                -1, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mInvoiceCancellation, Me.mNPWPNo, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mNoTransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary, Me)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, vFrm)
                    'simpan data jurnal
                    If Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                        '**************************************************************************************************************************************
                        '**************************************************************************************************************************************
                        'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                        'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                        'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                        'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                        'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL
                        Dim ReturRupiah As Decimal = 0
                        Dim TotalRetur As Decimal = 0
                        Dim AmountTanpaSales As Decimal = 0
                        Dim DefPrecBase As Integer = 0
                        DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                        TotalRetur = Me.ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefPrecBase)
                        If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase)
                        End If
                        ReturRupiah = Me.ObjCD.DigitGrouping(TotalRetur - AmountTanpaSales, DefPrecBase)
                        '**************************************************************************************************************************************

                        Me.TransferJurnal_SR(Me.mPrimary, Me.mPrimaryCustomer, _
                        ReturRupiah, _
                        Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), _
                        ObjInv.GetNoInvoice(ObjInv.GetPrimaryInvoice_ByPrimaryDO(Me.mPrimaryDO)), Me.mSubTotal, _
                        Me.mTaxExchangeRate, Me.mPrimaryCurrency, Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mTaxAmount, Me.mDutystamp)

                    Else
                        Me.Hapus_Jurnal(Me.mPrimary)
                    End If

                    ObjSRCust.SimpanPerubahanCustom(Me.mPrimary, Me.mNoTransaksi)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan(dgvItem, vUserCode, vFrm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            ObjItem = Nothing
            ObjInv = Nothing
            ObjDO = Nothing
            ObjCur = Nothing
            ObjSRCust = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesReturn_Quick, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1DataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSalesReturn_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(dgvItem, vUserCode, vFrm, vTblCustom)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesReturn_Quick, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiSalesReturn_Quick_AddByCustomProduct.tbltrans_custom_salesreturn_item1DataTable) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim DTA2 As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter
        Dim ObjItem As New ClsTransaksiSalesReturnItem
        Dim UseTax_Int As Integer
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim PrimaryInv As Double = -1
        Dim ObjDO As New ClsTransaksiDO
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSRCust As New CustomSMESS.ClsTransaksiSalesReturn

        Try
            If Me.ApakahDataBenar(False, True, True, Me.mPrimary, dgvItem, vFrm.TblCustomInOutDummy) = True Then
                'hapuskan data stock in sales return lama
                Me.HapusStockInSalesReturn(Me.GetPrimaryStockInSR(Me.mPrimary), vUserCode)

                'simpan data stock in sales return
                If Me.BuatStockInSalesReturn(Me.mPrimaryStockInSalesReturn, vFrm, dgvItem, vUserCode, "", False) = False Then
                    MsgBox("Can't create stock in sales return, process cancelled", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If Me.mUseTax = True Then
                    UseTax_Int = 1
                Else
                    UseTax_Int = 0
                End If
                If Me.QuickSR = False And Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                    PrimaryInv = ObjDO.GetPrimaryInvoice(Me.mPrimaryDO)
                End If

                If DTA2.UpdateQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mPrimaryStockInSalesReturn, _
                Me.mPrimaryDO, Me.mTanggalTransaksi, Me.mRemarks, Me.mReasonOfReturn, _
                Me.mCetakNotaRetur, vUserCode, Now, Now, Me.mNoNotaRetur, Me.mTipeSR, Me.mPrimaryCurrency, UseTax_Int, Me.mPaymentTerm, _
                PrimaryInv, Me.mQuickSR, Me.mSubTotal, Me.mTaxPercent, Me.mTaxAmount, Me.mTaxExchangeRate, Me.mDutystamp, Me.mTotal, _
                Me.mPrimarySalesPerson, Me.mInvoiceNoQuickSR, Me.mInvoiceDateQuickSR, Me.mTaxInvoiceNoQuickSR, _
                Me.mPrimaryCustomer, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mInvoiceCancellation, Me.mNPWPNo, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mNoTransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary, Me)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, vFrm, vTblCustom)
                    'simpan data jurnal
                    If Me.mTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                        '**************************************************************************************************************************************
                        '**************************************************************************************************************************************
                        'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                        'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                        'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                        'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                        'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL
                        Dim ReturRupiah As Decimal = 0
                        Dim TotalRetur As Decimal = 0
                        Dim AmountTanpaSales As Decimal = 0
                        Dim DefPrecBase As Integer = 0
                        DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                        TotalRetur = Me.ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefPrecBase)
                        If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                            AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase)
                        End If
                        ReturRupiah = Me.ObjCD.DigitGrouping(TotalRetur - AmountTanpaSales, DefPrecBase)
                        '**************************************************************************************************************************************

                        Me.TransferJurnal_SR(Me.mPrimary, Me.mPrimaryCustomer, _
                        ReturRupiah, _
                        Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.ObjCD.DigitGrouping(Me.mDutystamp * Me.mTaxExchangeRate, DefPrecBase), _
                        Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), _
                        ObjInv.GetNoInvoice(ObjInv.GetPrimaryInvoice_ByPrimaryDO(Me.mPrimaryDO)), Me.mSubTotal, _
                        Me.mTaxExchangeRate, Me.mPrimaryCurrency, Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mTaxAmount, Me.mDutystamp)
                    Else
                        Me.Hapus_Jurnal(Me.mPrimary)
                    End If

                    ObjSRCust.SimpanPerubahanCustom(Me.mPrimary, Me.mNoTransaksi)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan(dgvItem, vUserCode, vFrm, vTblCustom)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            ObjItem = Nothing
            ObjInv = Nothing
            ObjDO = Nothing
            ObjCur = Nothing
            ObjSRCust = Nothing
        End Try
    End Function

    Public Function GetPrimaryAccountAP(ByVal vPrimary As Double) As Double
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryAccount(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAccount(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrTax(ByVal vPrimary As Double) As Double
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_NewTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCurrTax(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrTax(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================
    Public Enum EnumTipeSR As Integer
        Ganti_Barang
        Revisi_Invoice
        Nota_Retur_Pajak
    End Enum

    'elemen yang digunakan pada automatic SR transfer
    Public Enum EnumTipeSRBalance As Integer
        Automatic
        Manual
    End Enum
    Public Enum EnumTipeJournal As Integer
        GeneralJournal
        AdjustmenJournal
        ClosingJournal
        SalesJournal
        SalesReturnJournal
        SalesDownPayment
        PurchaseJournal
        PurchaseReturnJournal
        PurchaseDownPaymentJournal
        SalesReceivementJournal
        SalesReturnPaymentJournal
        PurchasePaymentJournal
        PurchaseReturnReceivement
    End Enum
    Public Enum Enum_Tipe_Setting
        None
        Debet
        Kredit
    End Enum

    Public Property PrimarySalesPerson() As Double
        Get
            Return Me.mPrimarySalesPerson
        End Get
        Set(ByVal value As Double)
            mPrimarySalesPerson = value
        End Set
    End Property

    Public Property PrimaryCustomer() As Double
        Get
            Return Me.mPrimaryCustomer
        End Get
        Set(ByVal value As Double)
            mPrimaryCustomer = value
        End Set
    End Property

    Public Property InvoiceNoQuickSR() As String
        Get
            Return Me.mInvoiceNoQuickSR
        End Get
        Set(ByVal value As String)
            mInvoiceNoQuickSR = value
        End Set
    End Property

    Public Property TaxInvoiceNoQuickSR() As String
        Get
            Return Me.mTaxInvoiceNoQuickSR
        End Get
        Set(ByVal value As String)
            mTaxInvoiceNoQuickSR = value
        End Set
    End Property

    Public Property InvoiceDateQuickSR() As Date
        Get
            Return Me.mInvoiceDateQuickSR
        End Get
        Set(ByVal value As Date)
            mInvoiceDateQuickSR = value
        End Set
    End Property

    Public Property QuickSR() As Boolean
        Get
            If Me.mQuickSR = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = False Then
                Me.mQuickSR = 0
            Else
                Me.mQuickSR = 1
            End If
        End Set
    End Property

    Public Property PaymentTerm() As Double
        Get
            Return Me.mPaymentTerm
        End Get
        Set(ByVal value As Double)
            mPaymentTerm = value
        End Set
    End Property

    Public Property Total() As Decimal
        Get
            Return Me.mTotal
        End Get
        Set(ByVal value As Decimal)
            mTotal = value
        End Set
    End Property

    Public Property TaxAmount() As Decimal
        Get
            Return Me.mTaxAmount
        End Get
        Set(ByVal value As Decimal)
            mTaxAmount = value
        End Set
    End Property

    Public Property Dutystamp() As Decimal
        Get
            Return Me.mDutystamp
        End Get
        Set(ByVal value As Decimal)
            mDutystamp = value
        End Set
    End Property

    Public Property UseTax() As Boolean
        Get
            Return Me.mUseTax
        End Get
        Set(ByVal value As Boolean)
            mUseTax = value
        End Set
    End Property

    Public Property TaxExchangeRate() As Decimal
        Get
            Return Me.mTaxExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mTaxExchangeRate = value
        End Set
    End Property

    Public Property SubTotal() As Decimal
        Get
            Return Me.mSubTotal
        End Get
        Set(ByVal value As Decimal)
            mSubTotal = value
        End Set
    End Property

    Public Property TaxPercent() As Decimal
        Get
            Return Me.mTaxPercent
        End Get
        Set(ByVal value As Decimal)
            mTaxPercent = value
        End Set
    End Property

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property

    Public Property TipeSR() As EnumTipeSR
        Get
            Return Me.mTipeSR
        End Get
        Set(ByVal value As EnumTipeSR)
            Me.mTipeSR = value
        End Set
    End Property

    Public Property NoNotaRetur() As String
        Get
            Return Me.mNoNotaRetur
        End Get
        Set(ByVal value As String)
            mNoNotaRetur = value
        End Set
    End Property

    Public Property ReturnReason() As String
        Get
            Return Me.mReasonOfReturn
        End Get
        Set(ByVal value As String)
            mReasonOfReturn = value
        End Set
    End Property

    Public Property PrimaryDO() As Double
        Get
            Return Me.mPrimaryDO
        End Get
        Set(ByVal value As Double)
            mPrimaryDO = value
        End Set
    End Property

    Public Property PrimaryStockInSalesReturn() As Integer
        Get
            Return Me.mPrimaryStockInSalesReturn
        End Get
        Set(ByVal value As Integer)
            mPrimaryStockInSalesReturn = value
        End Set
    End Property

    Public Property CetakNotaReturn() As Boolean
        Get
            If Me.mCetakNotaRetur = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mCetakNotaRetur = 1
            Else
                Me.mCetakNotaRetur = 0
            End If
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property NoUrut() As Double
        Get
            Return Me.mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Sub IsiTBLSRByPrimaryDO(ByVal ListDO As List(Of Double), ByVal TBL As dseTransaksiSalesReturn.DataTable3DataTable)
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.DataTable3TableAdapter
        Dim PrimaryDO As Double = 0

        Try
            DTA.ClearBeforeFill = False
            TBL.Rows.Clear()
            For Each PrimaryDO In ListDO
                DTA.FillBy_PrimaryDO(TBL, PrimaryDO)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan no sales return yang merujuk ke primary stock in sales return tertentu
    ''' </summary>
    ''' <param name="vPrimaryStockInPurchase"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNoTransaksiByStockIn(ByVal vPrimaryStockInPurchase As Double, ByVal vPrimarySR As Double) As String
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDonoByPrimaryStockIn(vPrimaryStockInPurchase)) = True Then
                Return ""
            Else
                If Me.GetPrimarySR(DTA.ScalarQuery_GetDonoByPrimaryStockIn(vPrimaryStockInPurchase)) <> vPrimarySR Then
                    Return DTA.ScalarQuery_GetDonoByPrimaryStockIn(vPrimaryStockInPurchase)
                Else
                    Return ""
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimarySR As Double) As Double
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimarySR).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimarySR).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimarySR(ByVal vNoTransaksi As String) As Double
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vNoTransaksi).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
      ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
      Optional ByVal vTanggalSampai As Date = (#1/1/1950#), Optional ByVal IsQuick As Integer = 0) As Long
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                Return DTA.ScalarQuery_GetJmlBySONumberLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDeliveryOrderNo Then
                Return DTA.ScalarQuery_GetJmlByDoNumberLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInSalesReturn Then
                Return DTA.ScalarQuery_GetJmlByStockInSRNo(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                Return DTA.ScalarQuery_GetJmlShiptoCustomerCode(Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                Return DTA.ScalarQuery_GetJmlShiptoCustomerDesc(Criteria + "%", IsQuick)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoSalesReturn(ByVal vNotaReturn As String, ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturnDataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturnRow
        Dim retval As String = ""
        Dim intDummy As Integer = 1

        Try
            DTA.FillBy_NomorNotaRetur(TBL, vNotaReturn, vPrimaryTransaksi)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If retval = "" Then
                        retval = intDummy.ToString + ". " + row.No_Transaksi
                    Else
                        retval = retval + vbCrLf + intDummy.ToString + ". " + row.No_Transaksi
                    End If
                    intDummy += 1
                Next
            End If

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/SR/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahQuickSR(ByVal vPrimary As Double) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            RetVal = DTA.ScalarQuery_GetIsQuick(vPrimary)
            If RetVal.HasValue = True Then
                If RetVal.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As dseTransaksiSalesReturn.tbltrans_salesreturn_view0DataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal IsQuick As Integer, _
   Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
   Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_view0Row

        Dim QueryResult As Double
        Dim StartT As Date

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari, IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                QueryResult = DTA.FillBy_SONumberLIKE(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                QueryResult = DTA.FillBy_CustNameLIKE(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpec(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInSalesReturn Then
                QueryResult = DTA.FillBy_NoStockInSalesReturnLIKE(TBL, Criteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                QueryResult = DTA.FillBy_ShiptoCustomerCodeLIKE(TBL, IsQuick, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                QueryResult = DTA.FillBy_ShiptoCustomerName(TBL, IsQuick, Criteria + "%")
            End If

            If IsQuick = 1 Then
                For Each row In TBL.Rows
                    row.Customer_Code = row.Cust_Code_Quick
                    row.Customer_Name = row.Cust_Name_Quick
                Next
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(9).Visible = False
                DGV.Columns(10).Visible = False
                DGV.Columns(11).Visible = False
                If IsQuick = 1 Then
                    DGV.Columns(3).Visible = False
                    DGV.Columns(8).Visible = False
                End If
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimarySR As Double, ByVal vSalesStaff As String, _
    Optional ByVal TampilkanPesan As Boolean = True, Optional ByVal vStrictSalesStaff As Boolean = True) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim ObjItemInvoice As New ClsTransaksiSalesInvoiceItem
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TanggalSR As Date = #1/1/1950#
        Dim ObjCL As New ClsTransaksiPurchaseClosing
        Dim CreatedBy As String = ""
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim Jurnal As String = ""
        Dim vTipeSR As EnumTipeSR = EnumTipeSR.Ganti_Barang
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesReturn

        Try
            'Suatu transaksi do tidak bisa diedit jika:
            '   b. invoice
            If ObjItemInvoice.GetJmlBySR(vPrimarySR) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("There are sales invoice: " + ObjItemInvoice.GetSalesInvoiceNoBySR(vPrimarySR) + _
                    " refering to this transaction, cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalSR = Me.GetTanggalTransaksi(vPrimarySR)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalSR Then
                MsgBox("The period " + Format(TanggalSR, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. jika bukan yang bersangkutan yang melakukan input data transaksi
            If vStrictSalesStaff = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimarySR)) = False Then
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimarySR)
                    If CreatedBy <> vSalesStaff Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '4. Jika sudah ada transaksi sales return payment untuk transaksi sales return ybs --> khusus
            '   untuk transaksi sales return nota return --> obselete sejak penggabungan sales receivement jurnal dengan sales payment jurnal
            'vTipeSR = Me.GetTipeSR_ByPrimary(vPrimarySR)
            'If vTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
            'Jurnal = ObjJurnal.Get_SalesReturnJurnal(vPrimarySR)
            'If Jurnal <> "" Then
            'If TampilkanPesan = True Then
            'MsgBox("This sales return referred by sales return payment journal: " + Jurnal + _
            '". You cannot edit this data.", MsgBoxStyle.Information, "Information")
            'End If

            'Return False
            'End If
            'End If

            '4.b. Jika sudah ada sales receivement jurnal yang merujuk pada transaksi ini
            vTipeSR = Me.GetTipeSR_ByPrimary(vPrimarySR)
            If vTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                Jurnal = ObjJurnal.Get_SalesInvoiceJurnal_SR(vPrimarySR)
                If Jurnal <> "" Then
                    If TampilkanPesan = True Then
                        MsgBox("This sales return referred by sales receivement journal: " + Jurnal + _
                        ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            End If


            If ObjCustom.ApakahBolehEdit_CustomSalesReturn(vPrimarySR, TampilkanPesan) = False Then
                Return False
            End If

            '5. msgbox("belum")

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItemInvoice = Nothing
            ObjCD = Nothing
            ObjCL = Nothing
            ObjJurnal = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan tipe sr, jika sr yang dimaksud tidak ada, maka secara default akan mengembalikan
    ''' "Ganti Barang"
    ''' </summary>
    ''' <param name="vPrimarySR"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTipeSR_ByPrimary(ByVal vPrimarySR As Double) As EnumTipeSR
        Dim vTipeSR As EnumTipeSR = EnumTipeSR.Ganti_Barang
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            If DTA.ScalarQuery_GetTipeSR(vPrimarySR).HasValue = True Then
                vTipeSR = DTA.ScalarQuery_GetTipeSR(vPrimarySR).Value
            End If

            Return vTipeSR
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'yang dicek hanya item yang tipenya goods
    Public Function ApakahHapusMengakibatkanMinus(ByVal vPrimaryTranskasi As Double, ByVal vTampilkanPesan As Boolean) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item3TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_item3DataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_item3Row
        Dim JmlLama As Double = 0
        Dim JmlBaru As Double = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim Dummy As String = ""
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim ListPrimaryCodeNo As New List(Of Double)

        Try
            DTA.Fill(TBL, vPrimaryTranskasi)

            For Each row In TBL.Rows
                JmlLama = ObjProduct.GetJml2(row.Primary_Product, DTA1, row.Secondary_Qty, _
                row.Primary_Location, row.Batch_No)
                If ObjProduct.ApakahGoods(row.Primary_Product) = True Then
                    If JmlLama <> 0 Then
                        JmlBaru = JmlLama - row.Total_Qty
                        If JmlBaru < 0 Then
                            If ListPrimaryCodeNo.Contains(row.Primary_Product) = False Then
                                ListPrimaryCodeNo.Add(row.Primary_Product)

                                'jumlah setelah hapus transaksi akan menjadi minus
                                If Dummy = "" Then
                                    Dummy = ObjProduct.GetProductCode(row.Primary_Product)
                                Else
                                    Dummy = Dummy + ", " + ObjProduct.GetProductCode(row.Primary_Product)
                                End If
                            End If
                        End If
                    Else
                        If ListPrimaryCodeNo.Contains(row.Primary_Product) = False Then
                            ListPrimaryCodeNo.Add(row.Primary_Product)

                            'tidak ada data digudang, data akan menjadi minus jika item dihapuskan
                            If Dummy = "" Then
                                Dummy = ObjProduct.GetProductCode(row.Primary_Product)
                            Else
                                Dummy = Dummy + ", " + ObjProduct.GetProductCode(row.Primary_Product)
                            End If
                        End If
                    End If
                End If
            Next

            'kalo sampe sini return false, tidak ada data yang minus jika transaksi ini dihapuskan
            If Dummy <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("You cannot delete this data because will cause negatif qty one or several product: " + _
                    Dummy, MsgBoxStyle.Information, "Information")
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProduct = Nothing
            TBL.Dispose()
            ListPrimaryCodeNo = Nothing
        End Try
    End Function

    Public Function ApakahHapusMengakibatkanMinus2(ByVal vPrimarySR As Double, ByVal vTampilkanPesan As Boolean) As Boolean
        Dim DtaSO As New dseTransaksiSalesReturnTableAdapters.tbltrans_so_item2TableAdapter
        Dim DtaDO As New dseTransaksiSalesReturnTableAdapters.tbltrans_do_item2TableAdapter
        Dim DtaSOC As New dseTransaksiSalesReturnTableAdapters.tbltrans_soc_item2TableAdapter
        Dim DtaSR As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item21TableAdapter
        Dim TblSO As New dseTransaksiSalesReturn.tbltrans_so_item2DataTable
        Dim TblDO As New dseTransaksiSalesReturn.tbltrans_do_item2DataTable
        Dim TblSOC As New dseTransaksiSalesReturn.tbltrans_soc_item2DataTable
        Dim TblSR, TblSR2 As New dseTransaksiSalesReturn.tbltrans_salesreturn_item21DataTable
        Dim rowSO As dseTransaksiSalesReturn.tbltrans_so_item2Row
        Dim rowDO As dseTransaksiSalesReturn.tbltrans_do_item2Row
        Dim rowSOC As dseTransaksiSalesReturn.tbltrans_soc_item2Row
        Dim rowSR, rowSR2 As dseTransaksiSalesReturn.tbltrans_salesreturn_item21Row

        Dim Dummy As String = ""
        Dim listSO As New List(Of Double)
        Dim listProduct As New List(Of Double)
        Dim PrimarySOMin As Double = -1
        Dim PrimarySOMax As Double = -1
        Dim qtySO, qtyDO, qtySOC, qtySR, qtyHitung, toleranceUp As Decimal
        Dim SOC_No As String = ""
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSO As New ClsTransaksiSalesOrder

        Try
            DtaSO.Connection.ConnectionString = Core_Module.GetConstring
            DtaDO.Connection.ConnectionString = Core_Module.GetConstring
            DtaSOC.Connection.ConnectionString = Core_Module.GetConstring
            DtaSR.Connection.ConnectionString = Core_Module.GetConstring

            DtaSR.FillBy_PrimarySR(TblSR, vPrimarySR)
            If TblSR.Count > 0 Then
                For Each rowSR In TblSR.Rows
                    With rowSR
                        If listSO.Contains(.Primary_SO) = False Then
                            listSO.Add(.Primary_SO)
                        End If
                    End With
                Next

                listSO.Sort()
                PrimarySOMin = listSO(0)
                PrimarySOMax = listSO(listSO.Count - 1)

                DtaSO.FillBy_PrimarySOMinMax(TblSO, PrimarySOMin, PrimarySOMax)
                DtaDO.FillBy_PrimarySOMinMax(TblDO, PrimarySOMin, PrimarySOMax)
                DtaSOC.FillBy_PrimarySOMinMax(TblSOC, PrimarySOMin, PrimarySOMax)
                DtaSR.FillBy_PrimarySOMinMax(TblSR2, PrimarySOMin, PrimarySOMax)

                If TblSO.Count > 0 Then
                    For Each rowSR In TblSR.Rows
                        qtySO = 0
                        qtyDO = 0
                        qtySOC = 0
                        qtySR = 0
                        toleranceUp = 0

                        TblSO.DefaultView.RowFilter = "[Primary SO] = " & rowSR.Primary_SO.ToString & " AND [Primary Product] = " _
                        & rowSR.Primary_Product.ToString & " AND [Primary Units Multi] = " & rowSR.Primary_Units_Multi
                        If TblSO.DefaultView.Count > 0 Then
                            rowSO = TblSO.DefaultView(0).Row
                            qtySO = rowSO.Qty_Multi
                            toleranceUp = rowSO.Tolerance_Up
                        End If
                        TblDO.DefaultView.RowFilter = "[Primary SO] = " & rowSR.Primary_SO.ToString & " AND [Primary Product] = " _
                        & rowSR.Primary_Product.ToString & " AND [Primary Units Multi] = " & rowSR.Primary_Units_Multi
                        If TblDO.DefaultView.Count > 0 Then
                            For Each dvDO As DataRowView In TblDO.DefaultView
                                rowDO = dvDO.Row
                                qtyDO += rowDO.Total_Qty_Multi
                            Next
                        End If
                        TblSOC.DefaultView.RowFilter = "[Primary SO] = " & rowSR.Primary_SO.ToString & " AND [Primary Product] = " _
                        & rowSR.Primary_Product.ToString & " AND [Primary Units Multi] = " & rowSR.Primary_Units_Multi
                        If TblSOC.DefaultView.Count > 0 Then
                            For Each dvSOC As DataRowView In TblSOC.DefaultView
                                rowSOC = dvSOC.Row
                                qtySOC += rowSOC.Qty_Multi
                            Next
                        End If
                        TblSR2.DefaultView.RowFilter = "[Primary SR] <> " & rowSR.Primary_SR.ToString & " AND [Primary SO] = " _
                        & rowSR.Primary_SO.ToString & " AND [Primary Product] = " & rowSR.Primary_Product.ToString _
                        & " AND [Primary Units Multi] = " & rowSR.Primary_Units_Multi
                        If TblSR2.DefaultView.Count > 0 Then
                            For Each dvSR As DataRowView In TblSR2.DefaultView
                                rowSR2 = dvSR.Row
                                qtySR += rowSR2.Total_Qty_Multi
                            Next
                        End If

                        'tidak mungkin qty so lebih kecil dari qty do, bila iya berarti 
                        'product pada so itu menggunakan tolerance up
                        If qtySO < qtyDO Then
                            qtySO += toleranceUp
                        End If

                        qtyHitung = qtySO - qtyDO - qtySOC + qtySR
                        If qtyHitung < 0 Then
                            If vTampilkanPesan = True Then
                                If Dummy = "" Then
                                    If rowSR.Primary_SO <> -1 Then
                                        Dummy = ObjProd.GetProductCode(rowSR.Primary_Product) & " - SO: " & ObjSO.GetNoSO(rowSR.Primary_SO)
                                    Else
                                        Dummy = ObjProd.GetProductCode(rowSR.Primary_Product)
                                    End If
                                    'listProduct.Add(rowSR.Primary_Product)
                                Else
                                    'If listProduct.Contains(rowSR.Primary_Product) = False Then
                                    'listProduct.Add(rowSR.Primary_Product)
                                    If rowSR.Primary_SO <> -1 Then
                                        Dummy &= ", " & ObjProd.GetProductCode(rowSR.Primary_Product) & " - SO: " & ObjSO.GetNoSO(rowSR.Primary_SO)
                                    Else
                                        Dummy &= ", " & ObjProd.GetProductCode(rowSR.Primary_Product)
                                    End If
                                    'End If
                                End If
                            Else
                                Return True
                            End If
                        End If
                    Next
                End If

                If vTampilkanPesan = True Then
                    If Dummy <> "" Then
                        MsgBox("You cannot delete this data because will cause negatif sales order balance qty for one or several product: " + _
                        Dummy, MsgBoxStyle.Information, "Information")
                        Return True
                    End If
                End If

                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DtaSO.Dispose()
            DtaDO.Dispose()
            DtaSOC.Dispose()
            DtaSR.Dispose()
            TblSO.Dispose()
            TblDO.Dispose()
            TblSOC.Dispose()
            TblSR.Dispose()
            ObjProd = Nothing
            ObjSO = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimarySR As Double, ByVal vSalesStaff As String, _
    Optional ByVal TampilkanPesan As Boolean = True, Optional ByVal vStrictSalesStaff As Boolean = True) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim ObjItemInvoice As New ClsTransaksiSalesInvoiceItem
        Dim ObjCD As New Core_Module.ClsCekData
        Dim TanggalSR As Date = #1/1/1950#
        Dim ObjCL As New ClsTransaksiPurchaseClosing
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim Jurnal As String = ""
        Dim vTipeSR As EnumTipeSR = EnumTipeSR.Ganti_Barang
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesReturn

        Try
            'Suatu transaksi do tidak bisa diedit jika:
            '   b. invoice
            If ObjItemInvoice.GetJmlBySR(vPrimarySR) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("There are sales invoice: " + ObjItemInvoice.GetSalesInvoiceNoBySR(vPrimarySR) + _
                    " refering to this transaction, cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalSR = Me.GetTanggalTransaksi(vPrimarySR)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalSR Then
                MsgBox("The period " + Format(TanggalSR, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
            'hanya sales staff yang bersangkutan yang boleh hapus
            If vStrictSalesStaff = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimarySR)) = False Then
                    CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimarySR)
                    If CreatedBy <> vSalesStaff Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '4. Jika sudah ada transaksi sales return payment untuk transaksi sales return ybs --> khusus
            '   untuk transaksi sales return nota return --> obselete sejak penggabungan sales receivement dengan sales payment jurnal
            'vTipeSR = Me.GetTipeSR_ByPrimary(vPrimarySR)
            'If vTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
            'Jurnal = ObjJurnal.Get_SalesReturnJurnal(vPrimarySR)
            'If Jurnal <> "" Then
            'If TampilkanPesan = True Then
            'MsgBox("This sales return referred by sales return payment journal: " + Jurnal + _
            '". You cannot delete this data.", MsgBoxStyle.Information, "Information")
            'End If

            'Return False
            'End If
            'End If

            '4.b Jika sudah ada transaksi sales receivement untuk transaksi sales return ybs --> khusus
            '   untuk transaksi sales return nota return
            vTipeSR = Me.GetTipeSR_ByPrimary(vPrimarySR)
            If vTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                Jurnal = ObjJurnal.Get_SalesInvoiceJurnal_SR(vPrimarySR)
                If Jurnal <> "" Then
                    If TampilkanPesan = True Then
                        MsgBox("This sales return referred by sales receivement journal: " + Jurnal + _
                        ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            End If

            '4. simulasi stock
            If Me.ApakahQuickSR(vPrimarySR) = True Then
                If Me.ApakahHapusMengakibatkanMinus(vPrimarySR, TampilkanPesan) = True Then
                    Return False
                End If
            Else
                If Me.ApakahHapusMengakibatkanMinus2(vPrimarySR, True) = True Then
                    Return False
                End If
            End If

            If ObjCustom.ApakahBolehHapus_CustomSalesReturn(vPrimarySR, TampilkanPesan) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItemInvoice = Nothing
            ObjCD = Nothing
            ObjCL = Nothing
            ObjJurnal = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimarySR As Double) As Date
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            If DTA.ScalarQuery_GetTanggal(vPrimarySR).HasValue = False Then
                Return #1/1/1950#
            Else
                Return DTA.ScalarQuery_GetTanggal(vPrimarySR).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vQuickSR As Boolean, _
    ByVal TampilkanPesan As Boolean, ByVal vPrimarySR As Double, ByVal vDGVItem As DataGridView, _
    ByVal TblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim Jurnal As String = ""
        Dim vTipeSR As EnumTipeSR = EnumTipeSR.Ganti_Barang
        Dim TanggalSR As Date = #1/1/1950#
        Dim ObjCL As New ClsTransaksiPurchaseClosing
        Dim vFrmSR As New FrmTransaksiSalesReturn_Quick

        'elemen simulasi stock
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item3TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_item3DataTable
        Dim rowInOutLama As dseTransaksiSalesReturn.tbltrans_salesreturn_item3Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date
        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""
        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim TipeItemService As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Services
        Dim DTAService As New Inventory_Module.dseTransaksiStockInOutTableAdapters.tblmaster_productTableAdapter
        Dim TBLService As New Inventory_Module.dseTransaksiStockInOut.tblmaster_productDataTable

        Try
            If vCekDataBenar = False Then
                Return True
            End If

            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                End If
            End If

            '2. simulasikan stock
            If vQuickSR = True Then
                If vBaru = False Then
                    'load database master product qty berdasarkan data product lama dan data product baru
                    TBL.Rows.Clear()
                    DTA.Fill(TBL, vPrimarySR)
                    For Each rowInOutLama In TBL.Rows
                        With rowInOutLama
                            DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                            .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                        End With
                    Next
                End If

                'load data item baru
                DTAService.Fill(TBLService, TipeItemService)
                For Each rowDGV In vDGVItem.Rows
                    PrimaryProduct = rowDGV.Cells(vFrmSR.ColPrimaryCodeNo.Index).Value
                    'jika service maka tidak masuk
                    If IsNothing(TBLService.FindByPrimary(PrimaryProduct)) = True Then
                        BatchNo = rowDGV.Cells(vFrmSR.ColBatchNo.Index).Value
                        PrimaryLocation = rowDGV.Cells(vFrmSR.ColPrimaryLocation.Index).Value
                        If IsNumeric(rowDGV.Cells(vFrmSR.ColSecondaryQty.Index).Value) = True Then
                            Sec_Qty = rowDGV.Cells(vFrmSR.ColSecondaryQty.Index).Value
                        Else
                            Sec_Qty = "-1"
                        End If
                        If IsDate(rowDGV.Cells(vFrmSR.ColExpiryDate.Index).Value) = False Then
                            ExpiryDate = #1/1/1950#
                        Else
                            ExpiryDate = CDate(rowDGV.Cells(vFrmSR.ColExpiryDate.Index).Value)
                        End If
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
                    End If
                Next

                If TBL.Rows.Count > 0 Then
                    For Each rowInOutLama In TBL.Rows
                        With rowInOutLama
                            Me.Koreksi4HapusIn(TBLMasterProductQty, _
                            .Primary_Product, .Qty, .Secondary_Qty, _
                            .Primary_Location, .Batch_No, -1, .Expiry_Date)
                        End With
                    Next
                End If
                'simulasi simpan item baru
                For Each rowDGV In vDGVItem.Rows
                    BatchNo = rowDGV.Cells(vFrmSR.ColBatchNo.Index).Value
                    PrimaryLocation = rowDGV.Cells(vFrmSR.ColPrimaryLocation.Index).Value
                    If IsNumeric(rowDGV.Cells(vFrmSR.ColSecondaryQty.Index).Value) = True Then
                        Sec_Qty = rowDGV.Cells(vFrmSR.ColSecondaryQty.Index).Value
                    Else
                        Sec_Qty = "-1"
                    End If
                    If IsDate(rowDGV.Cells(vFrmSR.ColExpiryDate.Index).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(vFrmSR.ColExpiryDate.Index).Value)
                    End If
                    PrimaryProduct = rowDGV.Cells(vFrmSR.ColPrimaryCodeNo.Index).Value
                    Qty = rowDGV.Cells(vFrmSR.ColQty.Index).Value

                    Me.Koreksi4SimpanIn(TBLMasterProductQty, _
                    PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                    BatchNo, -1, ExpiryDate)
                Next
                ListPrimaryProduct.Clear()
                For Each rowProduct In TBLMasterProductQty.Rows
                    If rowProduct.Qty < 0 Then
                        If ObjProduct.ApakahGoods(rowProduct.Primary_Product) = True Then
                            ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                        End If
                    End If
                Next
                If ListPrimaryProduct.Count > 0 Then
                    While ListPrimaryProduct.Count > 0
                        If Dummy = "" Then
                            Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                        Else
                            Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                        End If
                    End While

                    MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
                    MsgBoxStyle.Information, "Information")

                    Return False
                End If
            End If

            '3. jika sudah ada jurnal sales yg merujuk pada transaksi ini
            vTipeSR = Me.GetTipeSR_ByPrimary(vPrimarySR)
            If vTipeSR = EnumTipeSR.Nota_Retur_Pajak Then
                Jurnal = ObjJurnal.Get_SalesInvoiceJurnal_SR(vPrimarySR)
                If Jurnal <> "" Then
                    If TampilkanPesan = True Then
                        MsgBox("This sales return referred by sales receivement journal: " + Jurnal + _
                        ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            End If

            '4. jika sudah diposting/tanggal transaksi < tanggal kunci terakhir
            TanggalSR = Me.mTanggalTransaksi
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalSR Then
                MsgBox("The period " + Format(TanggalSR, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vQuickSR = True Then
                'cek batasan minus stock inout custom
                '################################CUSTOM INOUT#############################################
                Dim ObjCustInOut As New CustomSMESS.ClsTransaksiStockInOut
                Dim ListProductCustom As New List(Of Double)
                Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
                If ObjCustInOut.ApakahDataBenar_CustomStockInOut(Me.mPrimaryStockInSalesReturn, False, CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In, _
                tblProdQty, TblCustomInOutDummy, ListProductCustom, CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_SalesReturn) = False Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjJurnal = Nothing
            ObjCL = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ListPrimaryProduct = Nothing
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ObjProduct = Nothing
            vFrmSR = Nothing
            DTAService.Dispose()
            TBLService.Dispose()
        End Try
    End Function

    Public Function Koreksi4SimpanIn(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
   ByVal vPrimaryProduct As Double, _
   ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Double = -1, Optional ByVal _
   vPrimaryLocation As Double = -1, Optional ByVal vBatchNo As String = "", _
   Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
   #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                    vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
 ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
 ByVal vPrimaryLocation As Double, ByVal vBatchNo As String) As Decimal
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " _
            + vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.Qty
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Hapus_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
  ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
  ByVal vPrimaryLocation As Double, ByVal vBatchNo As String)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " + _
            vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                TBL.Rows.Remove(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetJml_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
   ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
   ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
   ByVal JmlBaru As Decimal)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " _
            + vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                row.Qty = JmlBaru
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Insert_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
   ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
   ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
   ByVal JmlBaru As Decimal)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Try
            row = TBL.Newtblmaster_productqtyRow

            row.Primary = Me.GetPrimaryMax_TBLQtySimulasi(TBL) + 1
            row.Primary_Product = vPrimaryProduct
            row.Primary_Storage_Area = vPrimaryLocation
            row.Qty_Secondary = vSecQty
            row.Lot_No = vBatchNo
            'expiry date diabaikan dalam simulasi karena tidak berpengaruh ke databse
            'yang sebenarnya
            row.Expiry_Date = #1/1/1950#
            row.Qty = JmlBaru

            TBL.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_TBLQtySimulasi(ByVal TBL As _
  Core_Module.dseMasterProduct.tblmaster_productqtyDataTable) As Double
        Dim Sort As String = "[Primary] DESC"
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Try
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = Sort
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.Primary
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4HapusIn(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
 ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Double = -1, _
 Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
 vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, _
 Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)

            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty * -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub HapusStockInSalesReturn(ByVal vPrimaryStockInSR As Double, ByVal vUserCode As String)
        Dim Obj As New Inventory_Module.clsTransaksiStockInOut
        Dim TblCustomInOutDummy As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
        Dim ij As New FrmTransaksiStockOutSales 'new

        Try
            Obj.CariDetail(vPrimaryStockInSR, ij.dgvItem, TblCustomInOutDummy)
            Obj.Hapus1(vPrimaryStockInSR, vUserCode, TblCustomInOutDummy, False, -1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Function BuatStockInSalesReturn(ByRef vPrimaryStockInSR As Double, ByVal vFrm As FrmTransaksiSalesReturn_Quick, ByVal vDGVItem As DataGridView, _
    ByVal vUserCode As String, Optional ByVal NoStockInSR As String = "", Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vKoreksiStock As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim FrmStockIn As New Inventory_Module.FrmTransaksiStockInSalesReturn2
        Dim rowDGVSR, rowDGVStockIn As DataGridViewRow

        Try
            FrmStockIn.Show()
            FrmStockIn.WindowState = FormWindowState.Normal
            FrmStockIn.Top = (FrmStockIn.Height * -1) - 1

            'persiapkan formnya
            FrmStockIn.NewStockInQuickSR()
            FrmStockIn.txtPrimaryDO.Text = -1
            FrmStockIn.TxtPrimaryPihakAcuan.Text = Me.mPrimaryCustomer
            FrmStockIn.dtpDate.Value = Me.mTanggalTransaksi
            If NoStockInSR <> "" Then
                FrmStockIn.TxtStockInNo.Text = NoStockInSR
            End If

            'isikan item

            For Each rowDGVSR In vDGVItem.Rows
                FrmStockIn.dgvItem.Rows.Add()
                rowDGVStockIn = FrmStockIn.dgvItem.Rows(FrmStockIn.dgvItem.Rows.Count - 1)

                With rowDGVStockIn
                    .Cells(FrmStockIn.Index_BatchNo).Value = rowDGVSR.Cells(vFrm.ColBatchNo.Index).Value
                    .Cells(FrmStockIn.Index_CodeNo).Value = rowDGVSR.Cells(vFrm.ColCodeNo.Index).Value
                    .Cells(FrmStockIn.Index_Desc).Value = rowDGVSR.Cells(vFrm.ColDescription.Index).Value
                    .Cells(FrmStockIn.Index_Expiry).Value = rowDGVSR.Cells(vFrm.ColExpiryDate.Index).Value
                    .Cells(FrmStockIn.Index_Location).Value = rowDGVSR.Cells(vFrm.ColLocation.Index).Value
                    .Cells(FrmStockIn.Index_No).Value = rowDGVSR.Cells(vFrm.ColNo.Index).Value
                    .Cells(FrmStockIn.Index_PONoDONo).Value = rowDGVSR.Cells(vFrm.ColSONO.Index).Value
                    .Cells(FrmStockIn.Index_PrimaryCodeNo).Value = rowDGVSR.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                    .Cells(FrmStockIn.Index_PrimaryLocation).Value = rowDGVSR.Cells(vFrm.ColPrimaryLocation.Index).Value
                    .Cells(FrmStockIn.Index_PrimaryPdO).Value = rowDGVSR.Cells(vFrm.ColPrimaryPdO.Index).Value
                    .Cells(FrmStockIn.Index_PrimarySecUnit).Value = rowDGVSR.Cells(vFrm.ColPrimarySecondaryUnits.Index).Value
                    .Cells(FrmStockIn.Index_PrimarySO).Value = rowDGVSR.Cells(vFrm.ColPrimarySO.Index).Value
                    .Cells(FrmStockIn.Index_PrimaryUnits).Value = rowDGVSR.Cells(vFrm.ColPrimaryUnits.Index).Value
                    .Cells(FrmStockIn.Index_Qty).Value = rowDGVSR.Cells(vFrm.ColQty.Index).Value
                    .Cells(FrmStockIn.Index_Remarks).Value = rowDGVSR.Cells(vFrm.ColRemark.Index).Value
                    .Cells(FrmStockIn.Index_SecQty).Value = rowDGVSR.Cells(vFrm.ColSecondaryQty.Index).Value
                    .Cells(FrmStockIn.Index_SecUnit).Value = rowDGVSR.Cells(vFrm.ColSecondaryUnit.Index).Value
                    .Cells(FrmStockIn.Index_SuppPODONo).Value = rowDGVSR.Cells(vFrm.ColSuppDONo.Index).Value
                    .Cells(FrmStockIn.Index_Unit).Value = rowDGVSR.Cells(vFrm.ColUnit.Index).Value
                    .Cells(FrmStockIn.Index_Warehouse).Value = rowDGVSR.Cells(vFrm.ColWareHouse.Index).Value
                    .Cells(FrmStockIn.Index_QtyMulti).Value = rowDGVSR.Cells(vFrm.ColQtyMulti.Index).Value
                    .Cells(FrmStockIn.Index_UnitsCodeMulti).Value = rowDGVSR.Cells(vFrm.ColUnitsMulti.Index).Value
                    .Cells(FrmStockIn.Index_PrimaryUnitsMulti).Value = rowDGVSR.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value
                    .Cells(FrmStockIn.Index_RatioDefault).Value = rowDGVSR.Cells(vFrm.ColRatioDefault.Index).Value
                    .Cells(FrmStockIn.Index_RatioReference).Value = rowDGVSR.Cells(vFrm.ColRatioReference.Index).Value
                    .Cells(FrmStockIn.Index_Ratio_Desc).Value = rowDGVSR.Cells(vFrm.ColRatioDesc.Index).Value
                End With
            Next

            FrmStockIn.SalinKeObjek()
            If vFromCutOff = True Then
                FrmStockIn.objStockIn.CreatedBy = Me.mCreatedBy
                FrmStockIn.objStockIn.CreatedDate = Me.mCreatedDate
                FrmStockIn.objStockIn.ModifiedBy = Me.mModifiedBy
                FrmStockIn.objStockIn.ModifiedDate = Me.mModifiedDate
                FrmStockIn.objStockIn.TimeStamp = Me.mTimestamp
            End If

            'konfirmasi disimpan atau tidak
            If FrmStockIn.objStockIn.SimpanBaru1(FrmStockIn.dgvItem, vUserCode, vFrm.TblCustomInOutDummy, False, False, -1, vCekDataBenar, vKoreksiStock, vFromCutOff) = True Then
                vPrimaryStockInSR = FrmStockIn.objStockIn.Primary
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmStockIn.Dispose()
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.mNoUrut
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1

        Try
            TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
            If TanggalDari.Date < #5/1/2010# Then
                'tanggan ini dipilih 1 mei 2010 karena yakin pada tanggal tersebut di CMSI
                'tidak ada penomoran bulanan lagi
                TanggalDari = #5/1/2010#
            End If
            TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

            LoopCount = 0
            While LoopGo = True
                If DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                    JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).Value
                Else
                    JmlTrans = 1
                End If
                If JmlTrans <= 1 Then
                    LoopGo = False
                Else
                    If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
                        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari)
                    Else
                        PrimaryMax = -1
                    End If
                    If PrimaryMax = Me.mPrimary Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                        Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 6))
                    End If
                End If
                LoopCount += 1
            End While
            If LoopCount > 1 Then
                DTA.UpdateQuery_NoTrans_NoUrut(Me.mNoTransaksi, Me.mNoUrut, Me.mPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub


    Public Sub GetListPrimarySO(ByVal vPrimarySR As Double, ByVal vListPrimarySO As List(Of Double))
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item2TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_item2DataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_item2Row

        Try
            DTA.Fill(TBL, vPrimarySR)
            vListPrimarySO.Clear()
            For Each row In TBL.Rows
                vListPrimarySO.Add(row.Primary_SO)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Function GetPrimaryStockInSR(ByVal vPrimarySR As Double) As Double
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            RetVal = DTA.ScalarQuery_GetPrimaryStockInSR(vPrimarySR)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiSalesReturn_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, TampilkanPesan)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim ObjSOItem As New ClsTransaksiSalesOrderItem
        Dim ListPrimarySO As New List(Of Double)
        Dim ApakahQuick As Boolean = False
        Dim PrimaryStockIn As Double = -1

        Try
            ApakahQuick = Me.ApakahQuickSR(vPrimaryTransaksi)
            If ApakahQuick = True Then
                PrimaryStockIn = Me.GetPrimaryStockInSR(vPrimaryTransaksi)
            Else
                PrimaryStockIn = -1
            End If
            If Me.ApakahBolehHapus(vPrimaryTransaksi, vUserCode, TampilkanPesan) = True Then
                'hapus data jurnal
                Me.Hapus_Jurnal(vPrimaryTransaksi)
                'hapus data header, data item otomatis dihapuskan oleh server
                Me.GetListPrimarySO(vPrimaryTransaksi, ListPrimarySO)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    ObjSOItem.Set_Order_Balance(ListPrimarySO)

                    'hapus stock in sales return
                    If ApakahQuick = True Then
                        Me.HapusStockInSalesReturn(PrimaryStockIn, vUserCode)
                    End If

                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjSOItem = Nothing
            ListPrimarySO = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimaryStockInSalesReturn = -1
            Me.mCreatedBy = ""
            Me.mCreatedDate = Now
            Me.mModifiedBy = ""
            Me.mModifiedDate = Now
            Me.mNoTransaksi = ""
            Me.mNoUrut = -1
            Me.mPrimary = -1
            Me.mRemarks = ""
            Me.mTanggalTransaksi = Now
            Me.mTimestamp = Now
            mPrimaryDO = -1
            mCetakNotaRetur = 0
            mReasonOfReturn = ""

            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetSRNO(ByVal vPrimarySR As Double) As String
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksi(vPrimarySR)) = False Then
                Return DTA.ScalarQuery_GetNoTransaksi(vPrimarySR)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CetakSalesReturnArtha1(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimarySalesReturn As Double) As Boolean
        Dim DS As New dseTransaksiSalesReturn
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Try
            Dim DTAHeader As New dseTransaksiSalesReturnTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimarySalesReturn(DS.TblCetakHeader, vprimarySalesReturn)
            DTAHeader.Dispose()
            Dim DTAIsi As New dseTransaksiSalesReturnTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.FillBy_PrimarySalesReturn(DS.TblCetakIsi, vprimarySalesReturn)
            DTAIsi.Dispose()

            Dim rowIsi As dseTransaksiSalesReturn.TblCetakIsiRow

            Dim AdaExpiryDate As Boolean = False
            Dim AdaBatchNo As Boolean = False
            Dim AdaQtyPerBatch As Boolean = False
            Dim AdaSecondaryQty As Boolean = False

            For Each rowIsi In DS.TblCetakIsi.Rows
                If rowIsi.Expiry_Date <> #1/1/1950# Then
                    AdaExpiryDate = True
                    Exit For
                End If
            Next

            For Each rowIsi In DS.TblCetakIsi.Rows
                If rowIsi.Batch_No <> "" Then
                    AdaBatchNo = True
                    Exit For
                End If
            Next

            For Each rowIsi In DS.TblCetakIsi.Rows
                If rowIsi.QtyPerBatch >= 0 Then
                    AdaQtyPerBatch = True
                    Exit For
                End If
            Next

            For Each rowIsi In DS.TblCetakIsi.Rows
                If rowIsi.Secondary_Qty >= 0 Then
                    AdaSecondaryQty = True
                    Exit For
                End If
            Next

            If AdaBatchNo = True And AdaExpiryDate = True And AdaQtyPerBatch = True Then
                Me.CetakComplete(CRV, vprimarySalesReturn, DS, FormatDay, FormatMonth, FormatYears, FormatDecimal)
            ElseIf AdaBatchNo = True And AdaExpiryDate = False And AdaQtyPerBatch = True Then
                Me.CetakTanpaED(CRV, vprimarySalesReturn, DS, FormatDay, FormatMonth, FormatYears, FormatDecimal)
            ElseIf AdaBatchNo = False And AdaExpiryDate = False And AdaSecondaryQty = True Then
                Me.CetakTanpaEDNBatchNo(CRV, vprimarySalesReturn, DS, FormatDay, FormatMonth, FormatYears, FormatDecimal)
            ElseIf AdaBatchNo = False And AdaExpiryDate = False And AdaSecondaryQty = False Then
                Me.CetakTanpaEDNBatchNoNQtyPerBatch(CRV, vprimarySalesReturn, DS, FormatDay, FormatMonth, FormatYears, FormatDecimal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

    Private Sub CetakComplete(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryDO As Double, ByVal DS As dseTransaksiSalesReturn, ByVal FormatDay As Integer, _
    ByVal FormatMonth As Integer, ByVal FormatYears As Integer, ByVal FormatDecimal As Integer)
        Dim report As New CrpCetakTransaksiSalesReturnComplete
        Try
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CetakTaxReturn(ByVal vPrimarySalesReturn As Double, ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal DS As dseTransaksiSalesReturn, ByVal vFormatDecimal As Double, ByVal ObjCD As Core_Module.ClsCekData) As Boolean
        Dim ObjSC As New Core_Module.ClsTransaksiSetting
        Dim ObjSR_Custom As New CustomSMESS.ClsTransaksiSalesReturn
        If ObjSC.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            Return ObjSR_Custom.CetakSalesReturn(vCRV, vPrimarySalesReturn)
        End If
        Dim report As New CrpCetakTransaksiNotaRetur
        Dim DTAHeader As New dseTransaksiSalesReturnTableAdapters.TblCetakHeaderTableAdapter
        Dim rowHeader As dseTransaksiSalesReturn.TblCetakHeaderRow
        Dim DTAItem As New dseTransaksiSalesReturnTableAdapters.TblCetakIsiTableAdapter
        Dim rowItem As dseTransaksiSalesReturn.TblCetakIsiRow
        Dim SubTotal, SubTotal2, PPN_Dikembalikan As Double
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim Def_Prec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DTACust As New Core_Module.dseCustomerTableAdapters.tblmaster_customerTableAdapter
        Dim TBLCust As New Core_Module.dseCustomer.tblmaster_customerDataTable
        Dim rowCust As Core_Module.dseCustomer.tblmaster_customerRow
        Dim ExcRate As Decimal = 0
        Dim TaxPercent As Decimal = 0
        'Update NPWP 16 Digit
        Dim NPWP_PKP As String = ""
        Try
            Def_Prec = ObjCur.GetDecimal_DefaultCurrency
            'fill database
            DTAHeader.FillBy_PrimarySalesReturn(DS.TblCetakHeader, vPrimarySalesReturn)
            DTAItem.FillBy_PrimarySalesReturn(DS.TblCetakIsi, vPrimarySalesReturn)
            'hitung amount

            'dapatkan Exc Rate
            rowHeader = DS.TblCetakHeader.Rows(0)
            ExcRate = rowHeader.Tax_Exchange_Rate

            'update NPWP 16 Digit
            Dim objSettingNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If rowHeader.Tgl_SR.Date >= objSettingNPWP.Tglberlaku.Date Then
                NPWP_PKP = ObjCompany.GetCompany_NPWP
            Else
                NPWP_PKP = ObjCompany.GetCompany_NPWP_Old
            End If

            For Each rowItem In DS.TblCetakIsi.Rows
                rowItem.Amount = ObjCD.DigitGrouping(CDec(rowItem.Total_Qty_Multi * rowItem.Nett_Price), Def_Prec)
                SubTotal2 += CDec(rowItem.Amount)
                rowItem.Amount = ObjCD.DigitGrouping(CDec((rowItem.Total_Qty_Multi * rowItem.Nett_Price) * ExcRate), Def_Prec)
                SubTotal += CDec(rowItem.Amount)
                rowItem.Nett_Price = rowItem.Nett_Price * ExcRate
            Next
            If rowHeader.Use_Tax = 1 Then
                TaxPercent = rowHeader.Tax_Percent / 100
            End If
            PPN_Dikembalikan = Math.Round(SubTotal2 * TaxPercent, 1, MidpointRounding.AwayFromZero)
            PPN_Dikembalikan = PPN_Dikembalikan * ExcRate
            'SubTotal = SubTotal * ExcRate

            'set data header jika quick
            rowHeader = DS.TblCetakHeader.Rows(0)
            If rowHeader.Quick_Sales_Return = 1 Then
                DTACust.FillBy_Primary(TBLCust, rowHeader.Primary_Customer_Quick)
                rowCust = TBLCust.Rows(0)
                With rowHeader
                    .Tax_Add_Line2 = rowCust.Tax_Add_Line2
                    .Tax_Add_Line1 = rowCust.Tax_Add_Line1
                    .NPWP_No = rowCust.NPWP_No
                    .No_Invoice = .Invoice_No_Quick
                    .Invoice_Date = .Invoice_Date_Quick
                    .No_Faktur_Pajak = .Tax_Invoice_No_Quick
                    .Customer_Name = rowCust.Customer_Name
                    .Customer_Code = rowCust.Customer_Code
                    .City = rowCust.City
                End With
            End If

            'update NPWP 16 Digit
            If String.IsNullOrEmpty(rowHeader.New_NPWP_No) = False Then
                rowHeader.NPWP_No = rowHeader.New_NPWP_No
            End If

            'set report
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_DefaulPrecision.ParameterFieldName, Def_Prec)
            report.SetParameterValue(report.Parameter_JmlHargaJualDikembalikan.ParameterFieldName, _
            Math.Round(SubTotal, Def_Prec, MidpointRounding.AwayFromZero))
            report.SetParameterValue(report.Parameter_JmlPPnDikembalikan.ParameterFieldName, _
            PPN_Dikembalikan)
            report.SetParameterValue(report.Parameter_CompanyName.ParameterFieldName, ObjCompany.GetCompanyName)
            report.SetParameterValue(report.Parameter_Tax_Add1.ParameterFieldName, ObjCompany.GetCompany_TaxAdd1)
            report.SetParameterValue(report.Parameter_Tax_Add2.ParameterFieldName, ObjCompany.GetCompany_TaxAdd2)
            'Update NPWP 16 digit
            'report.SetParameterValue(report.Parameter_NPWP.ParameterFieldName, ObjCompany.GetCompany_NPWP)
            report.SetParameterValue(report.Parameter_NPWP.ParameterFieldName, NPWP_PKP)


            vCRV.ReportSource = report
            vCRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAHeader.Dispose()
            DTAItem.Dispose()
            ObjCompany = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Private Sub CetakTanpaED(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryDO As Double, ByVal DS As dseTransaksiSalesReturn, ByVal FormatDay As Integer, _
    ByVal FormatMonth As Integer, ByVal FormatYears As Integer, ByVal FormatDecimal As Integer)
        Dim report As New CrpCetakTransaksiSalesReturnTanpaED
        Try
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakTanpaEDNBatchNo(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryDO As Double, ByVal DS As dseTransaksiSalesReturn, ByVal FormatDay As Integer, _
    ByVal FormatMonth As Integer, ByVal FormatYears As Integer, ByVal FormatDecimal As Integer)
        Dim report As New CrpCetakTransaksiSalesReturnTanpaEDNBatchNo
        Try
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakTanpaEDNBatchNoNQtyPerBatch(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryDO As Double, ByVal DS As dseTransaksiSalesReturn, ByVal FormatDay As Integer, _
    ByVal FormatMonth As Integer, ByVal FormatYears As Integer, ByVal FormatDecimal As Integer)
        Dim report As New CrpCetakTransaksiSalesReturnTanpaEDNBatchNoNQtyPerBatch
        Try
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SimpanSalesOrderCancellation(ByVal vDGVSR As DataGridView, ByVal vDGVSOC As DataGridView, _
    ByVal vCancellationReason As String, ByVal vUser As String, ByVal SalesReturnNo As String, _
    ByVal VFrmSR As FrmTransaksiSalesReturn)
        Dim OBJSOC As New ClsTransaksiSOC
        Dim ListPrimarySO As New List(Of Double)
        Dim row As DataGridViewRow
        Dim PrimarySO As Double
        Dim Msg As String = ""
        Dim Jml As Integer = 1
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjCD As New Core_Module.ClsCekData
        Dim row1 As DataGridViewRow

        Try
            'cari no sales order yang ada
            For Each row In vDGVSR.Rows
                If ListPrimarySO.Contains(row.Cells(VFrmSR.ColPrimarySO.Index).Value) = False Then
                    ListPrimarySO.Add(row.Cells(VFrmSR.ColPrimarySO.Index).Value)
                End If
            Next
            With OBJSOC
                .CancellationReason = vCancellationReason
                .NoTransaksi = OBJSOC.GetNoBaru(Today)
                .TanggalTransaksi = Today
            End With

            'simpan data cancellation berdasarkan data sales order
            For Each PrimarySO In ListPrimarySO
                vDGVSOC.Rows.Clear()

                For Each row In vDGVSR.Rows
                    If row.Cells(VFrmSR.ColPrimarySO.Index).Value = PrimarySO Then
                        Dim Index As Integer = Me.GetIndex(row.Cells(VFrmSR.ColPrimaryCodeNo.Index).Value, vDGVSOC)

                        'cari data yang sama persis
                        If Index <> -1 Then
                            'tambahkan nilai qty
                            row1 = vDGVSOC.Rows(Index)
                            row1.Cells(VFrmSR.SOC_ColQtyMulti.Index).Value = CDbl(row1.Cells(VFrmSR.SOC_ColQtyMulti.Index).Value) + CDbl(row.Cells(VFrmSR.ColTotalQtyMulti.Index).Value)
                            row1.Cells(VFrmSR.SOC_ColQty.Index).Value = CDbl(row1.Cells(VFrmSR.SOC_ColQty.Index).Value) + CDbl(row.Cells(VFrmSR.ColTotalQty.Index).Value)
                        Else
                            'tambahkan baris baru
                            vDGVSOC.Rows.Add()
                            row1 = vDGVSOC.Rows(vDGVSOC.RowCount - 1)
                            With row1
                                .Cells(VFrmSR.SOC_ColCodeNo.Index).Value = row.Cells(VFrmSR.ColCodeNo.Index).Value
                                .Cells(VFrmSR.SOC_ColDescription.Index).Value = row.Cells(VFrmSR.ColDescription.Index)
                                .Cells(VFrmSR.SOC_ColNo.Index).Value = vDGVSOC.RowCount
                                .Cells(VFrmSR.SOC_ColPrimaryCodeNo.Index).Value = row.Cells(VFrmSR.ColPrimaryCodeNo.Index).Value
                                .Cells(VFrmSR.SOC_ColPrimaryUnits.Index).Value = row.Cells(VFrmSR.ColPrimaryUnit.Index).Value
                                .Cells(VFrmSR.SOC_ColQty.Index).Value = row.Cells(VFrmSR.ColTotalQty.Index).Value
                                .Cells(VFrmSR.SOC_ColRemark.Index).Value = ""
                                .Cells(VFrmSR.SOC_ColUnit.Index).Value = row.Cells(VFrmSR.ColUnit.Index).Value

                                .Cells(VFrmSR.SOC_ColQtyMulti.Index).Value = row.Cells(VFrmSR.ColTotalQtyMulti.Index).Value
                                .Cells(VFrmSR.SOC_ColUnitsMulti.Index).Value = row.Cells(VFrmSR.ColUnitsMulti.Index).Value
                                .Cells(VFrmSR.SOC_ColPrimaryUnitsMulti.Index).Value = row.Cells(VFrmSR.ColPrimaryUnitsMulti.Index).Value
                                .Cells(VFrmSR.SOC_ColRatioDefault.Index).Value = row.Cells(VFrmSR.ColRatioDefault.Index).Value
                                .Cells(VFrmSR.SOC_ColRatioReference.Index).Value = row.Cells(VFrmSR.ColRatioReference.Index).Value
                                .Cells(VFrmSR.SOC_ColRatio.Index).Value = row.Cells(VFrmSR.ColRatioDesc.Index).Value

                            End With
                        End If
                    End If
                Next

                OBJSOC.PrimarySO = PrimarySO
                OBJSOC.SimpanBaru(vDGVSOC, vUser)

                If Msg = "" Then
                    Msg = CStr(Jml) + ". For SO Number: " + ObjSO.GetNoSO(PrimarySO) + " the SOC Number is: " + OBJSOC.NoTransaksi
                Else
                    Msg = Msg + vbCrLf + CStr(Jml) + ". For SO Number: " + ObjSO.GetNoSO(PrimarySO) + " the SOC Number is: " + OBJSOC.NoTransaksi
                End If
                Jml += 1
            Next

            'beritahukan ke user data sales order yang digunakan, tanyakan apakah perlu dikirimkan email
            'sebagai reminder, jika ya kirimkan email juga
            Msg = "By you request the software have generated sales order cancellation(s) regarding this sales return(" + SalesReturnNo _
            + "). They are:" + vbCrLf + Msg
            If MsgBox(Msg + vbCrLf + "Would you like the software to send notifications email for your notes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.SendSOCNotificationEmail(vUser, vUser, Msg)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBJSOC = Nothing
            ListPrimarySO = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan index baris dalam dgv yang sesuai dengan parameter yang diberikan
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView) As Integer
        Dim row As DataGridViewRow
        Dim i As New FrmTransaksiSalesReturn

        Try
            For Each row In vDGV.Rows
                If row.Cells(i.SOC_ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SendSOCNotificationEmail(ByVal User As String, ByVal UserTujuan As String, ByVal Content As String) As Boolean
        Dim ObjEmail As New E_mail_Module.ClsTransEmail
        Dim ListUser As New List(Of String)

        Try
            With ObjEmail
                .Content = Content + vbCrLf + _
                vbCrLf + vbCrLf + "Regards, " + vbCrLf _
                + vbCrLf + CStr(User)
                .CreatedBy = User
                .TipeEmail = E_mail_Module.ClsTransEmail.EnumTipeEmail.StandarEmail
                .Subject = "Software generated sales order cancellation"
            End With

            ListUser.Add(UserTujuan)

            If ObjEmail.SendEmail(ListUser) = True Then
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahTransferOtomatis(ByVal vTanggalSR As Date) As Boolean
        Dim DTA As New DseTransaksiSalesReturn_TransJurnalTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            'Transfer otomatis jika:
            '1. Tanggal invoice >= convertion date
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA.ScalarQuery_GetConvertionDate.Value.Date <= vTanggalSR.Date Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub IsikanTabelSRAmount(ByVal vTBL As DseTransaksiSalesReturn_TransJurnal.Sales_Return_TransfDataTable, _
    ByVal vPrimarySR As Double, ByVal vDefPrec As Integer, ByVal vTotalSRAmountIDR As Decimal)
        Dim DTAItem As New DseTransaksiSalesReturn_TransJurnalTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim TBLItem As New DseTransaksiSalesReturn_TransJurnal.tbltrans_salesreturn_itemDataTable
        Dim rowItem As DseTransaksiSalesReturn_TransJurnal.tbltrans_salesreturn_itemRow
        Dim Filter As String = ""
        Dim row As DseTransaksiSalesReturn_TransJurnal.Sales_Return_TransfRow
        Dim JmlBaris As Double = 0
        Dim DummyInd As Double = 0
        Dim DummySRAmount As Decimal = 0

        Try
            vTBL.Rows.Clear()

            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAItem.Fill(TBLItem, vPrimarySR)

            'hitung nilai sr amount dalam base currency + masukkan ke dalam tabel sales return
            JmlBaris = TBLItem.Rows.Count
            For Each rowItem In TBLItem.Rows
                DummyInd += 1

                If DummyInd = JmlBaris Then
                    rowItem.SR_Amount = vTotalSRAmountIDR - DummySRAmount
                Else
                    rowItem.SR_Amount = Math.Round(rowItem.Exchange_Rate * rowItem.Nett_Price * _
                    rowItem.Total_Qty, vDefPrec, MidpointRounding.AwayFromZero)
                    DummySRAmount += rowItem.SR_Amount
                End If

                Filter = "[Primary Akun] = " + rowItem.Primary_Sales_Return_Account.ToString
                vTBL.DefaultView.RowFilter = Filter
                If vTBL.DefaultView.Count = 0 Then
                    'tambahkan baris baru
                    row = vTBL.NewSales_Return_TransfRow

                    row.Primary_Akun = rowItem.Primary_Sales_Return_Account
                    row.Amount = rowItem.SR_Amount
                    row.Amount_Original = rowItem.Amount

                    vTBL.Rows.Add(row)
                Else
                    'tambahkan amount dan original amount
                    For Each DV As DataRowView In vTBL.DefaultView
                        row = DV.Row
                        row.Amount += rowItem.SR_Amount
                        row.Amount_Original += rowItem.Amount
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
        End Try
    End Sub

    Private Sub TransferJurnal_SR(ByVal vPrimarySR As Double, ByVal vPrimaryCustomer As Double, ByVal vSalesReturn As Decimal, ByVal vVAT As Decimal, _
    ByVal vDutystamp As Decimal, ByVal vUrutanSR As Double, ByVal vTanggalSR As Date, ByVal vSRNumber As String, _
    ByVal vUserCode As String, ByVal vPaymentTerm As Double, ByVal vRef_Inv_Number As String, _
    ByVal vSalesReturnOriginalAmount As Decimal, ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrency As Double, _
    ByVal vPrimaryCurrTax As Double, ByVal vPrimaryAccount As Double, ByVal vVATOriginal As Decimal, ByVal vDutystampOriginal As Decimal)

        Dim DTA_Balance As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic_SR
        Dim DTA_Jurnal As New DseTransaksiSalesReturn_TransJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiSalesReturn_TransJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.SalesReturnJournal
        Dim RefNumber As String = CStr(vUrutanSR)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit As Double
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim obj As New ClsTransaksi_Kartu_APAR
        Dim MonthStr As String = ""
        Dim PrimaryAkunAR, PrimaryAkunAR_BaseCurrency, PrimaryAkunAR_BaseIDR As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryAkun As Double = -1
        Dim TBLSR As New DseTransaksiSalesReturn_TransJurnal.Sales_Return_TransfDataTable
        Dim DefPrecSR As Integer = 0
        Dim rowSR As DseTransaksiSalesReturn_TransJurnal.Sales_Return_TransfRow
        Dim DefPrecBase As Integer = 0

        Try
            DefPrecSR = ObjCur.GetDefaultPrecission(vPrimaryCurrency)
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimarySR)
            'hapuskan data balance sales lama
            DTA_Balance.DeleteQuery(vPrimarySR, TipeBalance)
            'hapuskan data kartu lama
            obj.Hapus(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis)

            If Me.ApakahTransferOtomatis(vTanggalSR) = True Then
                'buat jurnal baru
                If RefNumber.Length < 6 Then
                    While RefNumber.Length < 6
                        RefNumber = "0" + RefNumber
                    End While
                End If
                MonthStr = vTanggalSR.Month.ToString
                While MonthStr.Length < 2
                    MonthStr = "0" + MonthStr
                End While
                'data header
                DTA_Jurnal.Insert("SRJ-" + RefNumber + "/" + MonthStr + "/" + vTanggalSR.Year.ToString, vTanggalSR, _
                vUrutanSR, TipeJurnal, vPrimarySR, vPrimaryCustomer, "SR: " + vSRNumber _
                + " - " + ObjCust.GetCustomerCode(vPrimaryCustomer) + " - " + ObjCust.GetCustName(vPrimaryCustomer), vUserCode, _
                Now, vUserCode, Now, Now)
                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimarySR).Value

                'data item               
                'AR
                PrimaryAkunAR = vPrimaryAccount
                PrimaryAkunAR_BaseCurrency = ObjCur.GetPrimaryAkunAR(Me.GetPrimaryCurrency(vPrimarySR))
                PrimaryAkunAR_BaseIDR = ObjCur.GetPrimaryAkunAR(Core_Module.PrimaryBaseCurrency)



                Debet = 0
                Credit = 0
                'SR
                Me.IsikanTabelSRAmount(TBLSR, vPrimarySR, DefPrecBase, vSalesReturn)
                Dim decSalesReturn As Decimal = 0
                decSalesReturn = vSalesReturn
                For intA As Integer = 0 To TBLSR.Rows.Count - 1
                    rowSR = TBLSR.Rows(intA)
                    If intA <> TBLSR.Rows.Count - 1 Then
                        decSalesReturn -= rowSR.Amount
                    Else
                        rowSR.Amount = decSalesReturn
                    End If
                Next
                For Each rowSR In TBLSR.Rows
                    DTA_JurnalItem.Insert(PrimaryJurnal, rowSR.Primary_Akun, rowSR.Amount, 0, "", rowSR.Amount, 0, 1)
                Next

                'simpan data vat out-debet
                If vVAT <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunVATOut
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkun, vVAT, 0, "", vVAT, 0, 1)
                        Else
                            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkun, vVATOriginal, 0, "", vVATOriginal, 0, 1)
                        End If
                    End If
                End If

                'simpan data duty stamp-debet
                If vDutystamp <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunDutyStamp
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkun, vDutystamp, 0, "", 0, vDutystamp, 1)
                        Else
                            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkun, vDutystampOriginal, 0, "", 0, vDutystampOriginal, 1)
                        End If
                    End If
                End If

                ''simpan data vat out
                'If vVAT <> 0 Then
                '    PrimaryAkun = Me.GetPrimaryAkunVATOuts
                '    If PrimaryAkun <> -1 Then
                '        DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkun, vVAT, 0, "", vVAT, 0, 1)
                '    End If
                'End If

                ''simpan data duty stamp
                'If vDutystamp <> 0 Then
                '    PrimaryAkun = Me.GetPrimaryAkunDutyStamp
                '    If PrimaryAkun <> -1 Then
                '        DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkun, vDutystamp, 0, "", vDutystamp, 0, 1)
                '    End If
                'End If



                'If PrimaryAkunAR = PrimaryAkunAR_BaseCurrency Then
                '    'idr
                '    Credit = vSalesReturn + vVAT + vDutystamp
                '    DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", Debet, Credit, 1)
                'Else
                '    'non idr
                '    Credit = vSalesReturn
                '    DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", vSalesReturnOriginalAmount, 0, vExchangeRate)
                '    Credit = vVAT + vDutystamp
                '    If Credit <> 0 Then
                '        DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR_BaseCurrency, Debet, Credit, "", Debet, Credit, 1)
                '    End If
                'End If


                Credit = 0
                Debet = 0
                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    Credit = vSalesReturn + vVAT + vDutystamp
                    'masukkan data cost yang ap + cost
                    DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                    Debet, vSalesReturnOriginalAmount + vVATOriginal + vDutystampOriginal, vExchangeRate)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = vSalesReturn + vVAT + vDutystamp
                        DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                        Debet, vSalesReturnOriginalAmount + vVATOriginal + vDutystampOriginal, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = vSalesReturn
                        DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                        Debet, vSalesReturnOriginalAmount, vExchangeRate)


                        'masukkan data cost yang ap + cost ke ap base
                        If vVATOriginal + vDutystampOriginal <> 0 Then
                            Credit = vVATOriginal + vDutystampOriginal
                            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                        End If
                    End If
                End If



                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAR = PrimaryAkunAR_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------

                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        Credit = vSalesReturn
                '        DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                '                                  Debet, vSalesReturnOriginalAmount, vExchangeRate)

                '        'masukkan data vat + dutystamp ke ap base
                '        Credit = vVAT + vDutystamp
                '        DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR_BaseCurrency, Debet, Credit, "", _
                '                                  Debet, vVATOriginal + vDutystampOriginal, vExchangeRate)
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    End If
                'Else
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI NON IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAR = PrimaryAkunAR_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vSalesReturn + vVAT + vDutystamp
                '            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                '            Debet, vSalesReturnOriginalAmount + vVATOriginal + vDutystampOriginal, vExchangeRate)
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vSalesReturn
                '            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                '            Debet, vSalesReturnOriginalAmount, vExchangeRate)


                '            'masukkan data cost yang ap + cost ke ap base
                '            Credit = vVATOriginal + vDutystampOriginal
                '            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vSalesReturn
                '            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                '            Debet, vSalesReturnOriginalAmount, vExchangeRate)

                '            Credit = vVAT + vDutystamp
                '            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR_BaseCurrency, Debet, Credit, "", _
                '            Debet, vVATOriginal + vDutystampOriginal, vExchangeRate)
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vSalesReturn
                '            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", Debet, vSalesReturnOriginalAmount, vExchangeRate)

                '            Credit = vVATOriginal + vDutystampOriginal
                '            DTA_JurnalItem.Insert(PrimaryJurnal, PrimaryAkunAR_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    End If
                'End If


                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                                         vSalesReturn * -1, vVAT * -1, vDutystamp * -1, (vSalesReturn + vVAT + vDutystamp) * -1, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                                         (vSalesReturn + vVAT + vDutystamp) * -1, 1, PrimaryAkunAR)

                    'buat kartu ar baru
                    obj.Simpan1(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                    vPrimaryCustomer, 0, vSalesReturn + vVAT + vDutystamp, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                    vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                    vPrimaryCurrency, 0, vSalesReturn + vVAT + vDutystamp, 1, PrimaryAkunAR)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax Non IDR -------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'sales amount
                        DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                        vSalesReturn * -1, vVAT * -1, vDutystamp * -1, (vSalesReturn + vVAT + vDutystamp) * -1, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        (vSalesReturnOriginalAmount + vVATOriginal + vDutystampOriginal) * -1, vExchangeRate, PrimaryAkunAR)

                        'buat kartu ar baru
                        obj.Simpan1(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                        vPrimaryCustomer, 0, vSalesReturn + vVAT + vDutystamp, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                        vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                        vPrimaryCurrency, 0, vSalesReturnOriginalAmount + vVATOriginal + vDutystampOriginal, vExchangeRate, PrimaryAkunAR)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'sales amount
                        DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                        vSalesReturn * -1, 0, 0, vSalesReturn * -1, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        vSalesReturnOriginalAmount * -1, vExchangeRate, PrimaryAkunAR)

                        If (vVATOriginal + vDutystampOriginal) <> 0 Then
                            'tax amount + duty stamp
                            DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                            0, vVATOriginal * -1, vDutystampOriginal * -1, (vVATOriginal + vDutystampOriginal) * -1, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                            (vVATOriginal + vDutystampOriginal) * -1, 1, PrimaryAkunAR_BaseIDR)
                        End If

                        'buat kartu ar sales amount
                        obj.Simpan1(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                        vPrimaryCustomer, 0, vSalesReturn, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                        vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                        vPrimaryCurrency, 0, vSalesReturnOriginalAmount, vExchangeRate, PrimaryAkunAR)

                        If (vVAT + vDutystamp) <> 0 Then
                            'buat kartu ar tax amount + duty stamp
                            obj.Simpan1(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                            vPrimaryCustomer, 0, vVATOriginal + vDutystampOriginal, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                            vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                            Core_Module.PrimaryBaseCurrency, 0, vVATOriginal + vDutystampOriginal, 1, PrimaryAkunAR_BaseIDR)
                        End If
                    End If
                End If

                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                'Else
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI NON IDR
                '    '========================================================================================================================================
                '    If vPrimaryCurrTax = vPrimaryCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA Tax Non IDR -------------------------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'sales amount
                '        DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                '        vSalesReturn * -1, vVAT * -1, vDutystamp * -1, (vSalesReturn + vVAT + vDutystamp) * -1, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '        (vSalesReturnOriginalAmount + vVATOriginal + vDutystampOriginal) * -1, vExchangeRate, PrimaryAkunAR)

                '        'buat kartu ar baru
                '        obj.Simpan(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                '        vPrimaryCustomer, 0, vSalesReturn + vVAT + vDutystamp, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                '        vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                '        vPrimaryCurrency, 0, vSalesReturnOriginalAmount + vVATOriginal + vDutystampOriginal, PrimaryAkunAR, vExchangeRate)
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'sales amount
                '        DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                '        vSalesReturn * -1, 0, 0, vSalesReturn * -1, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '        vSalesReturnOriginalAmount * -1, vExchangeRate, PrimaryAkunAR)

                '        'tax amount + duty stamp
                '        DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                '        0, vVAT * -1, vDutystamp * -1, (vVAT + vDutystamp) * -1, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '        (vVAT + vDutystamp) * -1, 1, PrimaryAkunAR_BaseIDR)

                '        'buat kartu ar sales amount
                '        obj.Simpan(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                '        vPrimaryCustomer, 0, vSalesReturn, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                '        vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                '        vPrimaryCurrency, 0, vSalesReturnOriginalAmount, PrimaryAkunAR, vExchangeRate)

                '        'buat kartu ar tax amount + duty stamp
                '        obj.Simpan(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                '        vPrimaryCustomer, 0, vVAT + vDutystamp, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                '        vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                '        Core_Module.PrimaryBaseCurrency, 0, vVAT + vDutystamp, PrimaryAkunAR_BaseIDR, 1)
                '    End If
                'End If


                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                '    'sr idr
                '    DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                '    vSalesReturn * -1, vVAT * -1, vDutystamp * -1, (vSalesReturn + vVAT + vDutystamp) * -1, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '    (vSalesReturn + vVAT + vDutystamp) * -1, 1, PrimaryAkunAR)

                '    'buat kartu ar baru
                '    obj.Simpan(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                '    vPrimaryCustomer, 0, vSalesReturn + vVAT + vDutystamp, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                '    vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                '    Core_Module.PrimaryBaseCurrency, 0, vSalesReturn + vVAT + vDutystamp, PrimaryAkunAR_BaseCurrency, 1)
                'Else
                '    'sales amount
                '    DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                '    vSalesReturn * -1, 0, 0, vSalesReturn * -1, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '    vSalesReturnOriginalAmount * -1, vExchangeRate, PrimaryAkunAR)
                '    'tax amount + duty stamp
                '    DTA_Balance.Insert(vPrimaryCustomer, vPrimarySR, TipeBalance, "", #1/1/1950#, "", vTanggalSR + TimeSpan.FromDays(vPaymentTerm), _
                '    0, vVAT * -1, vDutystamp * -1, (vVAT + vDutystamp) * -1, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '    (vVAT + vDutystamp) * -1, 1, PrimaryAkunAR_BaseCurrency)

                '    'buat kartu ar sales amount
                '    obj.Simpan(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                '    vPrimaryCustomer, 0, vSalesReturn, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                '    vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                '    vPrimaryCurrency, 0, vSalesReturnOriginalAmount, PrimaryAkunAR, vExchangeRate)
                '    'buat kartu ar tax amount + duty stamp
                '    obj.Simpan(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, vSRNumber, _
                '    vPrimaryCustomer, 0, vVAT + vDutystamp, "Invoice Number: " + vRef_Inv_Number, ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                '    vTanggalSR, vPrimarySR, vTanggalSR, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis, _
                '    Core_Module.PrimaryBaseCurrency, 0, vVAT + vDutystamp, PrimaryAkunAR_BaseCurrency, 1)
                'End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTA_JurnalItem.Dispose()
            ObjCust = Nothing
            obj = Nothing
            ObjCur = Nothing
            TBLSR = Nothing
        End Try
    End Sub

    Private Function GetPrimaryAkunDutyStamp() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)

            If Dummy = "" Then
                Return -1
            Else
                Return Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ObjSetting = Nothing
        End Try
    End Function

    Private Function GetPrimaryAkunVATOut() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
            If Dummy = "" Then
                Return -1
            Else
                Return Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ObjSetting = Nothing
        End Try
    End Function

    Private Sub Hapus_Jurnal(ByVal vPrimarySR As Double)
        Dim DTA_Balance As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_balance_salesTableAdapter
        'Dim DTA_Balance As New DseTransaksiSalesReturn_TransJurnalTableAdapters.tbltrans_balance_salesreturnTableAdapter
        Dim DTA_Jurnal As New DseTransaksiSalesReturn_TransJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeBalance As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic_SR
        'Dim TipeBalance As EnumTipeSRBalance = EnumTipeSRBalance.Automatic
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.SalesReturnJournal
        Dim Obj As New ClsTransaksi_Kartu_APAR

        Try
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimarySR)
            'hapuskan data balance sales lama
            DTA_Balance.DeleteQuery(vPrimarySR, TipeBalance)
            'DTA_Balance.Delete(vPrimarySR, TipeBalance)
            'hapus data kartu AR
            Obj.Hapus(vPrimarySR, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.SR_Otomatis)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            Obj = Nothing
        End Try
    End Sub
End Class
